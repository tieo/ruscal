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
    NamedProps, PR_BODY, PR_CONTAINER_CLASS, PR_DISPLAY_NAME, PR_END_DATE,
    PR_ENTRYID, PR_IPM_SUBTREE_ENTRYID, PR_SENDER_NAME, PR_SENDER_SMTP_ADDRESS,
    PR_SENSITIVITY, PR_START_DATE, PR_SUBJECT,
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
    _store:       &IMsgStore,
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
    // 0  PR_SUBJECT            6  PR_SENDER_SMTP_ADDRESS   12 named.clip_end
    // 1  PR_START_DATE         7  named.location           13 named.clean_global_id
    // 2  PR_END_DATE           8  named.all_day
    // 3  PR_BODY               9  named.busy_status
    // 4  PR_SENSITIVITY        10 named.response_status
    // 5  PR_SENDER_NAME        11 named.recurring
    let mut cols = build_tag_array(&[
        PR_SUBJECT, PR_START_DATE, PR_END_DATE, PR_BODY, PR_SENSITIVITY,
        PR_SENDER_NAME, PR_SENDER_SMTP_ADDRESS,
        named.location, named.all_day, named.busy_status,
        named.response_status, named.recurring, named.clip_end,
        named.clean_global_id,
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

        events.push(CalendarEvent {
            subject:         unsafe { read_str8(&p[0],    PR_SUBJECT,              "(no subject)") },
            start:           unsafe { read_filetime(&p[1], PR_START_DATE) },
            end:             unsafe { read_filetime(&p[2], PR_END_DATE) },
            body:            unsafe { read_str8(&p[3],    PR_BODY,                 "") },
            sensitivity:     Sensitivity::from(unsafe { read_long(&p[4], PR_SENSITIVITY, 0) }),
            organizer_name:  unsafe { read_str8(&p[5],    PR_SENDER_NAME,          "") },
            organizer_email: unsafe { read_unicode(&p[6],  PR_SENDER_SMTP_ADDRESS,  "") },
            location:        unsafe { read_unicode(&p[7],  named.location,          "") },
            is_all_day:      unsafe { read_bool(&p[8],    named.all_day,           false) },
            busy_status:     BusyStatus::from(unsafe { read_long(&p[9],  named.busy_status,     2) }),
            response_status: ResponseStatus::from(unsafe { read_long(&p[10], named.response_status, 0) }),
            is_recurring,
            recurrence_end,
            clean_global_id: unsafe { read_binary(&p[13], named.clean_global_id) },
        });
    }

    unsafe { FreeProws(rows) };
    Ok(events)
}
