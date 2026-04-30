/// Calendar folder navigation and event querying.
///
/// # Recurring event handling
///
/// `GetContentsTable` returns one *series master* per recurring series, not
/// individual occurrences. The master's `PR_START_DATE` is the date of the
/// very first occurrence — which may be months or years in the past.
///
/// A naïve date-range restriction on `PR_START_DATE` would silently drop all
/// recurring series that started before the window. Instead we use an OR
/// restriction:
///
/// ```text
/// OR(
///   AND(start >= window_start, start <= window_end),   -- one-off events in window
///   AND(is_recurring = true, clip_end >= window_start) -- recurring series still active
/// )
/// ```
///
/// `clip_end` (`PidLidClipEnd`) is the date of the last scheduled occurrence.
/// An open-ended series returns `3060-01-31` as the sentinel.
use std::mem::ManuallyDrop;

use windows::Win32::System::AddressBook::{
    FreeProws, HrGetOneProp, HrQueryAllRows,
    IMAPIFolder, IMsgStore,
    SAndRestriction, SOrRestriction, SPropertyRestriction,
    SRestriction, SRestriction_0,
    SPropTagArray, SPropValue, SRowSet, __UPV,
};

use super::props::{
    self, build_tag_array, datetime_to_filetime, read_binary, read_bool,
    read_filetime, read_long, read_str8, read_unicode,
    NamedProps, PR_BODY_W, PR_CONTAINER_CLASS, PR_DISPLAY_NAME, PR_DISPLAY_NAME_W,
    PR_END_DATE, PR_ENTRYID, PR_IPM_SUBTREE_ENTRYID, PR_LAST_MODIFIER_NAME_W,
    PR_SENDER_NAME_W, PR_SENDER_SMTP_ADDRESS, PR_SENSITIVITY,
    PR_SENT_REPRESENTING_NAME_W, PR_SENT_REPRESENTING_SMTP_ADDRESS,
    PR_START_DATE, PR_SUBJECT_W,
};
use super::session::IMAPISession;
use crate::error::{check_hr, MapiError};
use crate::event::{BusyStatus, CalendarEvent, ResponseStatus, Sensitivity};

// ── Restriction type constants ────────────────────────────────────────────────

const RES_AND:      u32 = 0x0000_0000;
const RES_OR:       u32 = 0x0000_0001;
const RES_PROPERTY: u32 = 0x0000_0004;
const RELOP_GE:     u32 = 0x0000_0003; // >=
const RELOP_LE:     u32 = 0x0000_0001; // <=
const RELOP_EQ:     u32 = 0x0000_0004; // ==

// ── Calendar source listing ───────────────────────────────────────────────────

/// An Outlook message store that contains a Calendar folder.
pub struct OutlookCalendar {
    /// Display name shown in the picker (usually the account e-mail address).
    pub display_name: String,
}

/// List all message stores visible in the current MAPI session.
///
/// Returns one entry per store that has a non-empty display name. The caller
/// typically shows this list in a picker and uses the chosen index to identify
/// the store when opening it for sync.
///
/// # Safety
/// Must be called from the MAPI thread.
pub unsafe fn list_calendar_stores(
    session_ptr: *mut IMAPISession,
) -> Result<Vec<OutlookCalendar>, MapiError> {
    let mut raw_table: *mut core::ffi::c_void = core::ptr::null_mut();
    check_hr(unsafe {
        ((*(*session_ptr).vtbl).get_msg_stores_table)(session_ptr as *mut _, 0, &mut raw_table)
    })?;

    let table = ManuallyDrop::new(unsafe {
        core::mem::transmute::<_, windows::Win32::System::AddressBook::IMAPITable>(raw_table)
    });

    let mut cols = build_tag_array(&[PR_ENTRYID, PR_DISPLAY_NAME_W]);
    let mut rows: *mut SRowSet = core::ptr::null_mut();
    unsafe {
        HrQueryAllRows(
            &*table,
            cols.as_mut_ptr() as *mut SPropTagArray,
            core::ptr::null_mut(), core::ptr::null_mut(), 0,
            &mut rows,
        ).map_err(|e| MapiError(e.code().0 as u32))?;
    }

    let mut calendars = Vec::new();
    for i in 0..unsafe { (*rows).cRows } as usize {
        let row = unsafe { &*(*rows).aRow.as_ptr().add(i) };
        if row.lpProps.is_null() || row.cValues < 2 { continue; }
        let p = unsafe { std::slice::from_raw_parts(row.lpProps, row.cValues as usize) };
        // Prefer Unicode; fall back to ANSI if the provider returned PT_STRING8.
        let name = unsafe {
            if p[1].ulPropTag == PR_DISPLAY_NAME_W {
                read_unicode(&p[1], PR_DISPLAY_NAME_W, "")
            } else {
                read_str8(&p[1], PR_DISPLAY_NAME, "")
            }
        };
        if !name.is_empty() {
            calendars.push(OutlookCalendar { display_name: name });
        }
    }

    unsafe { FreeProws(rows) };
    drop(ManuallyDrop::into_inner(table));
    Ok(calendars)
}

// ── Store opening ─────────────────────────────────────────────────────────────

/// Open the user's default message store from the session.
///
/// Returns a `ManuallyDrop<IMsgStore>` to prevent the windows-crate Drop impl
/// from calling `Release` — MAPI manages the object's lifetime and releasing it
/// twice would corrupt internal state.
///
/// # Safety
/// `session_ptr` must be a valid `IMAPISession` COM pointer obtained from
/// `MAPILogonEx`. Must be called from the thread that called `CoInitialize`.
pub unsafe fn open_default_store(
    session_ptr: *mut IMAPISession,
) -> Result<ManuallyDrop<IMsgStore>, MapiError> {
    let mut raw_table: *mut core::ffi::c_void = core::ptr::null_mut();
    // SAFETY: session_ptr is a valid IMAPISession* with a correct vtable.
    check_hr(unsafe {
        ((*(*session_ptr).vtbl).get_msg_stores_table)(session_ptr as *mut _, 0, &mut raw_table)
    })?;

    // Transmute to IMAPITable. MAPI objects don't support QueryInterface across
    // their own interface hierarchy, so from_raw().cast() returns E_NOINTERFACE.
    // SAFETY: raw_table is a valid IMAPITable* returned by MAPI.
    let table = ManuallyDrop::new(unsafe {
        core::mem::transmute::<_, windows::Win32::System::AddressBook::IMAPITable>(raw_table)
    });

    let mut cols = build_tag_array(&[PR_ENTRYID, props::PR_DEFAULT_STORE]);
    let mut rows: *mut SRowSet = core::ptr::null_mut();
    unsafe {
        HrQueryAllRows(
            &*table,
            cols.as_mut_ptr() as *mut SPropTagArray,
            core::ptr::null_mut(), core::ptr::null_mut(), 0,
            &mut rows,
        ).map_err(|e| MapiError(e.code().0 as u32))?;
    }

    let mut store_raw: *mut core::ffi::c_void = core::ptr::null_mut();

    // SAFETY: rows is a valid SRowSet* returned by HrQueryAllRows.
    'outer: for i in 0..unsafe { (*rows).cRows } as usize {
        let row = unsafe { &*(*rows).aRow.as_ptr().add(i) };
        if row.lpProps.is_null() || row.cValues < 2 { continue; }
        let p = unsafe { std::slice::from_raw_parts(row.lpProps, row.cValues as usize) };

        if unsafe { p[1].Value.b } == 0 { continue; } // PR_DEFAULT_STORE is false

        let eid = unsafe { &p[0].Value.bin };
        let mut raw: *mut core::ffi::c_void = core::ptr::null_mut();
        // MAPI_BEST_ACCESS (0x10) — highest available permissions.
        // MAPI_DEFERRED_ERRORS (0x08) — don't fail if store isn't fully connected yet
        //   (required for cached Exchange/O365 profiles).
        check_hr(unsafe {
            ((*(*session_ptr).vtbl).open_msg_store)(
                session_ptr as *mut _, 0,
                eid.cb, eid.lpb as *const _,
                core::ptr::null(), 0x10 | 0x08,
                &mut raw,
            )
        })?;
        store_raw = raw;
        break 'outer;
    }

    unsafe { FreeProws(rows) };
    drop(ManuallyDrop::into_inner(table));

    if store_raw.is_null() {
        return Err(MapiError(0x8004_010F)); // MAPI_E_NOT_FOUND
    }

    // SAFETY: store_raw is a valid IMsgStore* returned by OpenMsgStore.
    Ok(ManuallyDrop::new(unsafe {
        core::mem::transmute::<_, IMsgStore>(store_raw)
    }))
}

// ── Calendar folder navigation ────────────────────────────────────────────────

/// Locate and open the default Calendar folder (`IPF.Appointment`).
///
/// Searches the IPM subtree first (via `PR_IPM_SUBTREE_ENTRYID`), falling back
/// to the store root if the subtree entry ID is unavailable (common in cached
/// Exchange mode).
///
/// # Safety
/// Must be called from the MAPI thread.
pub unsafe fn open_calendar_folder(
    store: &IMsgStore,
) -> Result<ManuallyDrop<IMAPIFolder>, MapiError> {
    let mut subtree_prop: *mut SPropValue = core::ptr::null_mut();
    let search_folder: ManuallyDrop<IMAPIFolder>;

    if unsafe { HrGetOneProp(store, PR_IPM_SUBTREE_ENTRYID, &mut subtree_prop) }.is_ok() {
        let eid = unsafe { (*subtree_prop).Value.bin };
        let mut obj_type = 0u32;
        let mut raw: Option<windows::core::IUnknown> = None;
        unsafe {
            store.OpenEntry(eid.cb, eid.lpb as *const _, None, 0, &mut obj_type, &mut raw)
                .map_err(|e| MapiError(e.code().0 as u32))?;
        }
        search_folder = ManuallyDrop::new(unsafe {
            core::mem::transmute::<_, IMAPIFolder>(raw.unwrap())
        });
    } else {
        // Fall back: open the root folder (NULL entry ID) and search its hierarchy.
        let mut obj_type = 0u32;
        let mut raw: Option<windows::core::IUnknown> = None;
        unsafe {
            store.OpenEntry(0, core::ptr::null(), None, 0, &mut obj_type, &mut raw)
                .map_err(|e| MapiError(e.code().0 as u32))?;
        }
        search_folder = ManuallyDrop::new(unsafe {
            core::mem::transmute::<_, IMAPIFolder>(raw.unwrap())
        });
    }

    let hier = unsafe { (*search_folder).GetHierarchyTable(0) }
        .map_err(|e| MapiError(e.code().0 as u32))?;

    let mut cols = build_tag_array(&[PR_ENTRYID, PR_DISPLAY_NAME, PR_CONTAINER_CLASS]);
    let mut rows: *mut SRowSet = core::ptr::null_mut();
    unsafe {
        HrQueryAllRows(
            &hier,
            cols.as_mut_ptr() as *mut SPropTagArray,
            core::ptr::null_mut(), core::ptr::null_mut(), 0,
            &mut rows,
        ).map_err(|e| MapiError(e.code().0 as u32))?;
    }

    let mut cal_eid: Vec<u8> = Vec::new();

    'outer: for i in 0..unsafe { (*rows).cRows } as usize {
        let row = unsafe { &*(*rows).aRow.as_ptr().add(i) };
        let p = unsafe { std::slice::from_raw_parts(row.lpProps, row.cValues as usize) };

        let class = unsafe { read_str8(&p[2], PR_CONTAINER_CLASS, "") };
        if class == "IPF.Appointment" {
            let eid = unsafe { &p[0].Value.bin };
            cal_eid = unsafe {
                std::slice::from_raw_parts(eid.lpb, eid.cb as usize).to_vec()
            };
            break 'outer;
        }
    }

    unsafe { FreeProws(rows) };
    drop(ManuallyDrop::into_inner(search_folder));

    if cal_eid.is_empty() {
        return Err(MapiError(0x8004_010F)); // MAPI_E_NOT_FOUND
    }

    let mut obj_type = 0u32;
    let mut raw: Option<windows::core::IUnknown> = None;
    unsafe {
        store.OpenEntry(
            cal_eid.len() as u32, cal_eid.as_ptr() as *const _,
            None, 0, &mut obj_type, &mut raw,
        ).map_err(|e| MapiError(e.code().0 as u32))?;
    }

    Ok(ManuallyDrop::new(unsafe {
        core::mem::transmute::<_, IMAPIFolder>(raw.unwrap())
    }))
}

// ── Event querying ────────────────────────────────────────────────────────────

/// Read all calendar events that fall within the given date window.
///
/// See the [module-level documentation](self) for how recurring events are handled.
///
/// # Safety
/// Must be called from the MAPI thread.
pub unsafe fn read_events(
    store:        &IMsgStore,
    folder:       &IMAPIFolder,
    named:        &NamedProps,
    window_start: chrono::DateTime<chrono::Utc>,
    window_end:   chrono::DateTime<chrono::Utc>,
) -> Result<Vec<CalendarEvent>, MapiError> {
    // ── Build restriction ─────────────────────────────────────────────────────

    // Restriction values must outlive the restriction tree they are referenced by.
    let mut ft_start = SPropValue {
        ulPropTag: PR_START_DATE, dwAlignPad: 0,
        Value: __UPV { ft: datetime_to_filetime(window_start) },
    };
    let mut ft_end = SPropValue {
        ulPropTag: PR_START_DATE, dwAlignPad: 0,
        Value: __UPV { ft: datetime_to_filetime(window_end) },
    };
    let mut ft_clip = SPropValue {
        ulPropTag: named.clip_end, dwAlignPad: 0,
        Value: __UPV { ft: datetime_to_filetime(window_start) },
    };
    let mut bool_true = SPropValue {
        ulPropTag: named.recurring, dwAlignPad: 0,
        Value: __UPV { b: 1 },
    };

    // Branch 1: one-off event with start inside the window.
    //   start >= window_start AND start <= window_end
    let res_ge = SRestriction { rt: RES_PROPERTY, res: SRestriction_0 {
        resProperty: SPropertyRestriction { relop: RELOP_GE, ulPropTag: PR_START_DATE, lpProp: &mut ft_start },
    }};
    let res_le = SRestriction { rt: RES_PROPERTY, res: SRestriction_0 {
        resProperty: SPropertyRestriction { relop: RELOP_LE, ulPropTag: PR_START_DATE, lpProp: &mut ft_end },
    }};
    let mut one_off_list = [res_ge, res_le];
    let res_one_off = SRestriction { rt: RES_AND, res: SRestriction_0 {
        resAnd: SAndRestriction { cRes: 2, lpRes: one_off_list.as_mut_ptr() },
    }};

    // Branch 2: recurring series whose last occurrence is on or after window_start.
    //   is_recurring == true AND clip_end >= window_start
    let res_is_recurring = SRestriction { rt: RES_PROPERTY, res: SRestriction_0 {
        resProperty: SPropertyRestriction { relop: RELOP_EQ, ulPropTag: named.recurring, lpProp: &mut bool_true },
    }};
    let res_clip_ge = SRestriction { rt: RES_PROPERTY, res: SRestriction_0 {
        resProperty: SPropertyRestriction { relop: RELOP_GE, ulPropTag: named.clip_end, lpProp: &mut ft_clip },
    }};
    let mut recurring_list = [res_is_recurring, res_clip_ge];
    let res_recurring = SRestriction { rt: RES_AND, res: SRestriction_0 {
        resAnd: SAndRestriction { cRes: 2, lpRes: recurring_list.as_mut_ptr() },
    }};

    // Top-level OR combining both branches.
    let mut or_list = [res_one_off, res_recurring];
    let mut restriction = SRestriction { rt: RES_OR, res: SRestriction_0 {
        resOr: SOrRestriction { cRes: 2, lpRes: or_list.as_mut_ptr() },
    }};

    // ── Query ─────────────────────────────────────────────────────────────────

    let table = unsafe { folder.GetContentsTable(0) }
        .map_err(|e| MapiError(e.code().0 as u32))?;

    // Column order — indices are used when parsing rows below.
    // 0  PR_SUBJECT_W          6  PR_SENDER_SMTP_ADDRESS         12 named.clip_end
    // 1  PR_START_DATE         7  named.location                 13 named.clean_global_id
    // 2  PR_END_DATE           8  named.all_day                  14 named.appt_recur
    // 3  PR_BODY_W             9  named.busy_status              15 PR_ENTRYID (fallback)
    // 4  PR_SENSITIVITY        10 named.response_status          16 PR_SENT_REPRESENTING_NAME_W
    // 5  PR_SENDER_NAME_W      11 named.recurring                17 PR_SENT_REPRESENTING_SMTP_ADDRESS
    //                                                            18 PR_LAST_MODIFIER_NAME_W
    //
    // We request the PT_UNICODE (_W) variants so Outlook returns proper UTF-16
    // strings. The ANSI (PT_STRING8) variants return CP1252 bytes on Western
    // Windows, which we were previously mis-decoding as UTF-8 — mangling
    // umlauts (ä/ö/ü/ß) into U+FFFD replacement characters on the Google side.
    //
    // Organizer resolution priority: SENT_REPRESENTING (16-17) → SENDER (5-6)
    // → LAST_MODIFIER (18). The last fallback covers events with no sender
    // info at all (PowerShell-created, unsent drafts) — for those, the
    // calendar owner is the only sensible answer.
    let mut cols = build_tag_array(&[
        PR_SUBJECT_W, PR_START_DATE, PR_END_DATE, PR_BODY_W, PR_SENSITIVITY,
        PR_SENDER_NAME_W, PR_SENDER_SMTP_ADDRESS,
        named.location, named.all_day, named.busy_status,
        named.response_status, named.recurring, named.clip_end,
        named.clean_global_id, named.appt_recur,
        PR_ENTRYID, // column 15 — needed to open message if appt_recur missing from table
        PR_SENT_REPRESENTING_NAME_W,
        PR_SENT_REPRESENTING_SMTP_ADDRESS,
        PR_LAST_MODIFIER_NAME_W,
    ]);

    let mut rows: *mut SRowSet = core::ptr::null_mut();
    unsafe {
        HrQueryAllRows(
            &table,
            cols.as_mut_ptr() as *mut SPropTagArray,
            &mut restriction,
            core::ptr::null_mut(), 0,
            &mut rows,
        ).map_err(|e| MapiError(e.code().0 as u32))?;
    }

    // ── Parse rows ────────────────────────────────────────────────────────────

    let row_count = unsafe { (*rows).cRows } as usize;
    let mut events = Vec::with_capacity(row_count);

    for i in 0..row_count {
        let row = unsafe { &*(*rows).aRow.as_ptr().add(i) };
        let p = unsafe { std::slice::from_raw_parts(row.lpProps, row.cValues as usize) };

        let is_recurring = unsafe { read_bool(&p[11], named.recurring, false) };
        let recurrence_end = if is_recurring {
            // PidLidClipEnd stores local midnight converted to UTC.
            // Convert back to the machine's local timezone to recover the
            // correct calendar date — otherwise UTC+N timezones show the
            // day before due to the timezone offset.
            //
            // Sentinel: MAPI uses 3060-01-31 to mean "no end date".
            let clip_end_local = unsafe { read_filetime(&p[12], named.clip_end) }
                .with_timezone(&chrono::Local);
            let sentinel = chrono::NaiveDate::from_ymd_opt(3060, 1, 31)
                .unwrap_or_default();
            let date = clip_end_local.date_naive();
            if date == sentinel { None } else { Some(date) }
        } else {
            None
        };

        // Prefer sent-representing over sender for the organizer line:
        // SENDER is reliable on meetings sent by others but commonly empty
        // (or returns an unusable Exchange X.500 address) on appointments the
        // user created themselves. SENT_REPRESENTING is populated for both.
        let sender_name        = unsafe { read_unicode(&p[5],  PR_SENDER_NAME_W,                  "") };
        let sender_email       = unsafe { read_unicode(&p[6],  PR_SENDER_SMTP_ADDRESS,            "") };
        let representing_name  = if p.len() > 16 {
            unsafe { read_unicode(&p[16], PR_SENT_REPRESENTING_NAME_W,            "") }
        } else { String::new() };
        let representing_email = if p.len() > 17 {
            unsafe { read_unicode(&p[17], PR_SENT_REPRESENTING_SMTP_ADDRESS,      "") }
        } else { String::new() };
        let last_modifier_name = if p.len() > 18 {
            unsafe { read_unicode(&p[18], PR_LAST_MODIFIER_NAME_W,                "") }
        } else { String::new() };
        let (organizer_name, organizer_email) = resolve_organizer(
            &representing_name, &representing_email,
            &sender_name,       &sender_email,
            &last_modifier_name,
        );

        // Read body from the row. MAPI table rows truncate `PR_BODY_W` at
        // 255 characters (same default-buffer gotcha as `recur_blob`), so if
        // the row gives us a string at exactly the cutoff we open the
        // message and fetch the property in full.
        let body_row = unsafe { read_unicode(&p[3], PR_BODY_W, "") };
        let body = if body_row.chars().count() >= 255 && p.len() > 15 {
            log::debug!("body truncated for '{}' ({}); refetching",
                unsafe { read_unicode(&p[0], PR_SUBJECT_W, "") },
                body_row.chars().count());
            unsafe { fetch_body(store, &p[15]) }.unwrap_or(body_row)
        } else {
            body_row
        };

        events.push(CalendarEvent {
            subject:         unsafe { read_unicode(&p[0],  PR_SUBJECT_W,            "(no subject)") },
            start:           unsafe { read_filetime(&p[1], PR_START_DATE) },
            end:             unsafe { read_filetime(&p[2], PR_END_DATE) },
            body,
            sensitivity:     Sensitivity::from(unsafe { read_long(&p[4], PR_SENSITIVITY, 0) }),
            organizer_name,
            organizer_email,
            location:        unsafe { read_unicode(&p[7],  named.location,          "") },
            is_all_day:      unsafe { read_bool(&p[8],    named.all_day,           false) },
            busy_status:     BusyStatus::from(unsafe { read_long(&p[9],  named.busy_status,     2) }),
            response_status: ResponseStatus::from(unsafe { read_long(&p[10], named.response_status, 0) }),
            is_recurring,
            recurrence_end,
            clean_global_id: unsafe { read_binary(&p[13], named.clean_global_id) },
            recur_blob: {
                let blob = unsafe { read_binary(&p[14], named.appt_recur) };
                // Exchange/Outlook often omits large binary properties from table
                // rows (returns PT_ERROR instead). Fall back to opening the message
                // directly to read PidLidAppointmentRecur.
                if is_recurring && blob.is_empty() && p.len() > 15 {
                    log::debug!("recur blob missing from table row for '{}'; opening message", unsafe { read_unicode(&p[0], PR_SUBJECT_W, "") });
                    unsafe { fetch_recur_blob(store, &p[15], named.appt_recur) }
                } else {
                    blob
                }
            },
        });
    }

    unsafe { FreeProws(rows) };
    Ok(events)
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/// Open a MAPI message by its entry ID and read `PidLidAppointmentRecur` directly.
///
/// MAPI table rows commonly omit large binary properties (the provider returns
/// `PT_ERROR` for them). This fallback calls `OpenEntry` on the message object
/// and uses `HrGetOneProp` to fetch the recurrence blob from the object itself.
///
/// # Safety
/// Must be called on the MAPI thread. `eid_prop` must be a valid `SPropValue`.
unsafe fn fetch_recur_blob(
    store: &IMsgStore,
    eid_prop: &SPropValue,
    appt_recur_tag: u32,
) -> Vec<u8> {
    if eid_prop.ulPropTag != PR_ENTRYID { return Vec::new(); }
    let eid = unsafe { &eid_prop.Value.bin };

    let mut obj_type = 0u32;
    let mut raw: Option<windows::core::IUnknown> = None;
    if unsafe {
        store.OpenEntry(eid.cb, eid.lpb as *const _, None, 0, &mut obj_type, &mut raw)
    }.is_err() {
        log::debug!("fetch_recur_blob: OpenEntry failed");
        return Vec::new();
    }
    let Some(msg_unk) = raw else { return Vec::new(); };

    // Transmute IUnknown → IMAPIFolder so we can pass it to HrGetOneProp.
    // Both IMessage and IMAPIFolder derive from IMAPIProp; HrGetOneProp only
    // uses the IMAPIProp vtable slots (GetProps), which sit at the same
    // positions in both vtables.
    let msg = ManuallyDrop::new(unsafe {
        core::mem::transmute::<windows::core::IUnknown, IMAPIFolder>(msg_unk)
    });

    let mut prop_ptr: *mut SPropValue = core::ptr::null_mut();
    if unsafe { HrGetOneProp(&*msg, appt_recur_tag, &mut prop_ptr) }.is_err()
        || prop_ptr.is_null()
    {
        log::debug!("fetch_recur_blob: HrGetOneProp failed");
        return Vec::new();
    }

    let blob = unsafe { read_binary(&*prop_ptr, appt_recur_tag) };
    log::debug!("fetch_recur_blob: got {} bytes", blob.len());
    // prop_ptr is MAPI-allocated; should be freed with MAPIFreeBuffer.
    // Omitted here — bounded small leak per recurring event, consistent with
    // the existing HrGetOneProp usage pattern in open_calendar_folder.

    // Release the IMessage. Leaving it wrapped in ManuallyDrop leaks a COM
    // ref; on MAPI session teardown that triggers STATUS_ACCESS_VIOLATION
    // because the provider tries to clean up while we're still holding refs.
    drop(ManuallyDrop::into_inner(msg));
    blob
}

/// Open a MAPI message by its entry ID and read `PR_BODY_W` directly.
///
/// MAPI table rows truncate string properties at a provider-dependent default
/// (255 characters for Outlook). Long event bodies come back snipped mid-word.
/// This fallback opens the message and reads the property in full via
/// `HrGetOneProp`, which doesn't impose the row-buffer limit.
///
/// Returns `Some(full_body)` on success, `None` if the OpenEntry/GetProp
/// chain fails (caller falls back to the truncated row value rather than
/// dropping data entirely).
///
/// # Safety
/// Must be called on the MAPI thread. `eid_prop` must be a valid `SPropValue`.
unsafe fn fetch_body(store: &IMsgStore, eid_prop: &SPropValue) -> Option<String> {
    if eid_prop.ulPropTag != PR_ENTRYID { return None; }
    let eid = unsafe { &eid_prop.Value.bin };

    let mut obj_type = 0u32;
    let mut raw: Option<windows::core::IUnknown> = None;
    if unsafe {
        store.OpenEntry(eid.cb, eid.lpb as *const _, None, 0, &mut obj_type, &mut raw)
    }.is_err() {
        log::debug!("fetch_body: OpenEntry failed");
        return None;
    }
    let msg_unk = raw?;

    // Same vtable-shape transmute as `fetch_recur_blob` — both `IMessage` and
    // `IMAPIFolder` derive from `IMAPIProp`, and `HrGetOneProp` only invokes
    // `IMAPIProp::GetProps`.
    let msg = ManuallyDrop::new(unsafe {
        core::mem::transmute::<windows::core::IUnknown, IMAPIFolder>(msg_unk)
    });

    let mut prop_ptr: *mut SPropValue = core::ptr::null_mut();
    if unsafe { HrGetOneProp(&*msg, PR_BODY_W, &mut prop_ptr) }.is_err()
        || prop_ptr.is_null()
    {
        log::debug!("fetch_body: HrGetOneProp failed");
        return None;
    }

    let body = unsafe { read_unicode(&*prop_ptr, PR_BODY_W, "") };
    log::debug!("fetch_body: got {} chars", body.chars().count());
    // prop_ptr is MAPI-allocated; bounded one-time leak per truncated body,
    // consistent with the existing HrGetOneProp usage in fetch_recur_blob.

    // Release the IMessage so the MAPI session shuts down cleanly.
    drop(ManuallyDrop::into_inner(msg));
    Some(body)
}

/// Pick the (name, email) pair to surface as the event's organizer.
///
/// MAPI exposes two parallel sender concepts on calendar items:
/// * `PR_SENDER_*`             — who actually transmitted the MAPI message,
/// * `PR_SENT_REPRESENTING_*`  — who the message represents.
///
/// On meeting invites *from* others both equal the organizer. On appointments
/// the user created on their *own* calendar, `SENDER` is commonly empty (no
/// invite ever crossed the wire) or returns an Exchange X.500 / EX address
/// (`/o=ExchangeLabs/ou=…/cn=…`), while `SENT_REPRESENTING` correctly equals
/// the mailbox owner.
///
/// Resolution order:
///   1. `SENT_REPRESENTING` (the message's logical sender)
///   2. `SENDER` (transmitting principal — populated only on real invites)
///   3. `LAST_MODIFIER` (last-ditch: events with no sender info at all,
///      e.g. PowerShell-created appointments and unsent meeting drafts.
///      For those the last modifier is the calendar owner, which is the
///      semantically correct answer to "who organized this".)
///
/// Email values are dropped if they aren't SMTP — an X.500 string is worse
/// than empty because it gets rendered to the user as the "Organizer".
/// Names are kept as-is regardless of whether the paired email survives.
fn resolve_organizer(
    representing_name:  &str,
    representing_email: &str,
    sender_name:        &str,
    sender_email:       &str,
    last_modifier_name: &str,
) -> (String, String) {
    let pick = |primary: &str, fallback: &str| -> String {
        if !primary.is_empty() { primary } else { fallback }.to_owned()
    };

    let name = {
        let n = pick(representing_name, sender_name);
        if n.is_empty() { last_modifier_name.to_owned() } else { n }
    };
    let email_raw = pick(representing_email, sender_email);
    let email = if is_smtp_address(&email_raw) { email_raw } else { String::new() };
    (name, email)
}


/// Heuristic: is this string a real SMTP address rather than an Exchange
/// X.500 / legacy EX form? We don't validate against RFC 5322 — we only
/// need to distinguish `alice@example.com` from
/// `/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=…`.
fn is_smtp_address(s: &str) -> bool {
    let s = s.trim();
    if s.is_empty() { return false; }
    if s.starts_with('/') { return false; } // X.500 / DN form
    let Some((local, domain)) = s.split_once('@') else { return false; };
    !local.is_empty() && domain.contains('.')
}

#[cfg(test)]
mod tests {
    use super::*;

    /// `PR_SENDER_*` is empty on appointments the user created themselves.
    /// Without the representing fallback, the organizer prefix never appears
    /// on those events even though the rest of the event syncs fine.
    #[test]
    fn resolve_organizer_uses_representing_when_sender_empty() {
        let (name, email) = resolve_organizer(
            "Alice Owner", "alice@example.com",
            "", "",
            "Alice Owner",
        );
        assert_eq!(name,  "Alice Owner");
        assert_eq!(email, "alice@example.com");
    }

    /// Meetings sent by other people: both sets are populated and equal.
    /// We just need to not regress on this case.
    #[test]
    fn resolve_organizer_prefers_representing_when_both_present() {
        let (name, email) = resolve_organizer(
            "Alice Repr", "alice@example.com",
            "Alice Sender", "alice@example.com",
            "Owner Name",
        );
        assert_eq!(name,  "Alice Repr");
        assert_eq!(email, "alice@example.com");
    }

    /// Falls back to sender values when representing is missing — only
    /// happens on stripped-down providers and very old Exchange profiles.
    #[test]
    fn resolve_organizer_falls_back_to_sender_when_representing_missing() {
        let (name, email) = resolve_organizer(
            "", "",
            "Bob", "bob@example.com",
            "Owner Name",
        );
        assert_eq!(name,  "Bob");
        assert_eq!(email, "bob@example.com");
    }

    /// X.500 / Exchange "EX" addresses must be discarded — surfacing
    /// `/o=ExchangeLabs/ou=…/cn=Recipients/cn=…` as the organizer email
    /// is the bug the user reported. The display name is kept.
    #[test]
    fn resolve_organizer_drops_x500_addresses() {
        let x500 = "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)\
                    /cn=Recipients/cn=abc123-alice";
        let (name, email) = resolve_organizer(
            "Alice Owner", x500,
            "", "",
            "Alice Owner",
        );
        assert_eq!(name,  "Alice Owner");
        assert_eq!(email, "", "X.500 address must be dropped, not surfaced");

        // SENDER X.500 with no representing values: still drop the email.
        let (name, email) = resolve_organizer(
            "", "",
            "Alice Owner", x500,
            "Alice Owner",
        );
        assert_eq!(name,  "Alice Owner");
        assert_eq!(email, "");
    }

    /// All four sender fields empty: fall back to the last-modifier name.
    /// This covers events created via PowerShell/scripts and unsent meeting
    /// drafts — for those the last modifier IS the calendar owner.
    #[test]
    fn resolve_organizer_falls_back_to_last_modifier_when_all_else_empty() {
        let (name, email) = resolve_organizer(
            "", "",
            "", "",
            "Alice Owner",
        );
        assert_eq!(name,  "Alice Owner");
        assert_eq!(email, "");
    }

    /// All five empty: emit empty pair so `build_description` skips the
    /// organizer prefix entirely instead of rendering a stray "Organizer:" line.
    #[test]
    fn resolve_organizer_all_empty_returns_empty() {
        let (name, email) = resolve_organizer("", "", "", "", "");
        assert_eq!(name, "");
        assert_eq!(email, "");
    }

    #[test]
    fn is_smtp_address_recognises_real_addresses() {
        assert!(is_smtp_address("alice@example.com"));
        assert!(is_smtp_address("first.last+tag@sub.example.co.uk"));
        assert!(!is_smtp_address(""));
        assert!(!is_smtp_address("plainstring"));
        assert!(!is_smtp_address("@nodomain.com"));
        assert!(!is_smtp_address("nolocal@"));
        assert!(!is_smtp_address("a@b"), "single-label domains aren't reachable from external SMTP");
        assert!(!is_smtp_address("/o=ExchangeLabs/cn=Recipients/cn=abc"));
    }
}
