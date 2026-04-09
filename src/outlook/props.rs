/// Property tag construction helpers and named-property resolution.
///
/// MAPI property tags are 32-bit values: `(property_id << 16) | property_type`.
///
/// *Standard* properties have fixed IDs below `0x8000`.
/// *Named* properties are looked up at runtime via [`resolve`], which maps
/// `(GUID, LID)` pairs to session-local IDs in the `0x8000–0xFFFE` range.
use windows::Win32::System::AddressBook::{
    IMsgStore, MAPINAMEID, MAPINAMEID_0, SPropTagArray, SPropValue,
};
use windows::core::GUID;

// ── Property types ────────────────────────────────────────────────────────────

#[allow(dead_code)]
pub const PT_BOOLEAN: u32 = 0x000B;
#[allow(dead_code)]
pub const PT_LONG:    u32 = 0x0003;
#[allow(dead_code)]
pub const PT_SYSTIME: u32 = 0x0040;
#[allow(dead_code)]
pub const PT_STRING8: u32 = 0x001E;
#[allow(dead_code)]
pub const PT_UNICODE: u32 = 0x001F;
#[allow(dead_code)]
pub const PT_BINARY:  u32 = 0x0102;

// ── Standard property tags ────────────────────────────────────────────────────

pub const PR_ENTRYID: u32             = 0x0FFF0102;
pub const PR_DEFAULT_STORE: u32       = 0x3400000B;
pub const PR_IPM_SUBTREE_ENTRYID: u32 = 0x35E00102;
pub const PR_DISPLAY_NAME: u32        = 0x3001001E;
pub const PR_CONTAINER_CLASS: u32     = 0x3613001E;
pub const PR_SUBJECT: u32             = 0x0037001E;
pub const PR_START_DATE: u32          = 0x00600040;
pub const PR_END_DATE: u32            = 0x00610040;
pub const PR_BODY: u32                = 0x1000001E;
pub const PR_SENSITIVITY: u32         = 0x00360003;
pub const PR_SENDER_NAME: u32         = 0x0C1A001E;
pub const PR_SENDER_SMTP_ADDRESS: u32 = 0x5D01001F;

// ── Named property GUIDs ──────────────────────────────────────────────────────

/// `PSETID_Appointment` — contains location, all-day, busy status, recurrence, etc.
///
/// `{00062002-0000-0000-C000-000000000046}`
pub const PSETID_APPOINTMENT: GUID = GUID {
    data1: 0x00062002, data2: 0x0000, data3: 0x0000,
    data4: [0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46],
};

/// `PSETID_Meeting` — contains the Global Object ID used for sync matching.
///
/// `{6ED8DA90-450B-101B-98DA-00AA003F1305}`
pub const PSETID_MEETING: GUID = GUID {
    data1: 0x6ED8DA90, data2: 0x450B, data3: 0x101B,
    data4: [0x98, 0xDA, 0x00, 0xAA, 0x00, 0x3F, 0x13, 0x05],
};

// ── Named property LIDs ───────────────────────────────────────────────────────

/// Free-text location / meeting URL (`PidLidLocation`, `PT_UNICODE`).
pub const LID_LOCATION: i32 = 0x8208;

/// All-day event flag (`PidLidAppointmentSubType`, `PT_BOOLEAN`).
pub const LID_ALL_DAY: i32 = 0x8215;

/// Busy/free status (`PidLidBusyStatus`, `PT_LONG`).
pub const LID_BUSY_STATUS: i32 = 0x8205;

/// Attendee response status (`PidLidResponseStatus`, `PT_LONG`).
pub const LID_RESPONSE_STATUS: i32 = 0x8218;

/// Whether this item is a recurring series master (`PidLidRecurring`, `PT_BOOLEAN`).
pub const LID_RECURRING: i32 = 0x8223;

/// Last occurrence date of a recurring series (`PidLidClipEnd`, `PT_SYSTIME`).
pub const LID_CLIP_END: i32 = 0x8236;

/// Stable cross-system sync ID (`PidLidCleanGlobalObjectId`, `PT_BINARY`).
///
/// Identical across organizer + attendee copies and all occurrences of a series.
/// Maps to `iCalUID` in Google Calendar.
pub const LID_CLEAN_GLOBAL_ID: i32 = 0x0023;

// ── Resolved named property tags ──────────────────────────────────────────────

/// Named property IDs resolved at session start via [`resolve`].
///
/// The IDs are stable within one store session but may differ across machines
/// or store providers, so they must always be looked up rather than hardcoded.
pub struct NamedProps {
    pub location:        u32,
    pub all_day:         u32,
    pub busy_status:     u32,
    pub response_status: u32,
    pub recurring:       u32,
    pub clip_end:        u32,
    pub clean_global_id: u32,
}

/// Resolve named property LIDs to session-local property tags.
///
/// Calls `GetIDsFromNames` on the store once. The returned tags are valid for
/// all objects in the same store for the duration of the MAPI session.
///
/// # Safety
/// Must be called from the thread that called `CoInitialize`.
pub unsafe fn resolve(store: &IMsgStore) -> Result<NamedProps, crate::error::MapiError> {
    let mut guid_appt    = PSETID_APPOINTMENT;
    let mut guid_meeting = PSETID_MEETING;

    let mut names = [
        MAPINAMEID { lpguid: &mut guid_appt,    ulKind: 0, Kind: MAPINAMEID_0 { lID: LID_LOCATION        } },
        MAPINAMEID { lpguid: &mut guid_appt,    ulKind: 0, Kind: MAPINAMEID_0 { lID: LID_ALL_DAY         } },
        MAPINAMEID { lpguid: &mut guid_appt,    ulKind: 0, Kind: MAPINAMEID_0 { lID: LID_BUSY_STATUS     } },
        MAPINAMEID { lpguid: &mut guid_appt,    ulKind: 0, Kind: MAPINAMEID_0 { lID: LID_RESPONSE_STATUS } },
        MAPINAMEID { lpguid: &mut guid_appt,    ulKind: 0, Kind: MAPINAMEID_0 { lID: LID_RECURRING       } },
        MAPINAMEID { lpguid: &mut guid_appt,    ulKind: 0, Kind: MAPINAMEID_0 { lID: LID_CLIP_END        } },
        MAPINAMEID { lpguid: &mut guid_meeting, ulKind: 0, Kind: MAPINAMEID_0 { lID: LID_CLEAN_GLOBAL_ID } },
    ];

    let mut ptrs: Vec<*mut MAPINAMEID> = names.iter_mut().map(|n| n as *mut _).collect();
    let mut tag_array: *mut SPropTagArray = core::ptr::null_mut();

    unsafe {
        store
            .GetIDsFromNames(ptrs.len() as u32, ptrs.as_mut_ptr(), 0, &mut tag_array)
            .map_err(|e| crate::error::MapiError(e.code().0 as u32))?;

        // Each returned tag has the dynamic ID in the high 16 bits and PT_UNSPECIFIED
        // (0x0000) in the low 16 bits. OR in the actual property type for each field.
        let ids = std::slice::from_raw_parts(
            &(*tag_array).aulPropTag as *const u32,
            (*tag_array).cValues as usize,
        );

        Ok(NamedProps {
            location:        (ids[0] & 0xFFFF_0000) | PT_UNICODE,
            all_day:         (ids[1] & 0xFFFF_0000) | PT_BOOLEAN,
            busy_status:     (ids[2] & 0xFFFF_0000) | PT_LONG,
            response_status: (ids[3] & 0xFFFF_0000) | PT_LONG,
            recurring:       (ids[4] & 0xFFFF_0000) | PT_BOOLEAN,
            clip_end:        (ids[5] & 0xFFFF_0000) | PT_SYSTIME,
            clean_global_id: (ids[6] & 0xFFFF_0000) | PT_BINARY,
        })
    }
}

// ── Property reading helpers ──────────────────────────────────────────────────

/// Build a MAPI property tag array from a slice.
///
/// [`SPropTagArray`] uses a `[u32; 1]` C flexible-array-member placeholder.
/// We build a `Vec<u32>` with the count prepended and cast its pointer —
/// the memory layout is identical to what MAPI expects.
pub fn build_tag_array(tags: &[u32]) -> Vec<u32> {
    let mut v = Vec::with_capacity(1 + tags.len());
    v.push(tags.len() as u32);
    v.extend_from_slice(tags);
    v
}

/// Read a `PT_STRING8` (ANSI) property, returning `fallback` if the tag doesn't match.
///
/// # Safety
/// `prop.Value.lpszA` must be a valid null-terminated string when the tag matches.
pub unsafe fn read_str8(prop: &SPropValue, expected_tag: u32, fallback: &str) -> String {
    if prop.ulPropTag != expected_tag {
        return fallback.to_owned();
    }
    unsafe {
        std::ffi::CStr::from_ptr(prop.Value.lpszA.0 as *const i8)
            .to_string_lossy()
            .into_owned()
    }
}

/// Read a `PT_UNICODE` (UTF-16) property, returning `fallback` if the tag doesn't match.
///
/// # Safety
/// `prop.Value.lpszW` must be a valid null-terminated UTF-16 string when the tag matches.
pub unsafe fn read_unicode(prop: &SPropValue, expected_tag: u32, fallback: &str) -> String {
    if prop.ulPropTag != expected_tag {
        return fallback.to_owned();
    }
    unsafe {
        let ptr = prop.Value.lpszW.0;
        if ptr.is_null() {
            return fallback.to_owned();
        }
        let len = (0..).take_while(|&i| *ptr.add(i) != 0).count();
        String::from_utf16_lossy(std::slice::from_raw_parts(ptr, len))
    }
}

/// Read a `PT_LONG` property, returning `fallback` if the tag doesn't match.
///
/// # Safety
/// `prop.Value.l` must be valid when the tag matches.
pub unsafe fn read_long(prop: &SPropValue, expected_tag: u32, fallback: u32) -> u32 {
    if prop.ulPropTag != expected_tag {
        return fallback;
    }
    unsafe { prop.Value.l as u32 }
}

/// Read a `PT_BOOLEAN` property, returning `fallback` if the tag doesn't match.
///
/// # Safety
/// `prop.Value.b` must be valid when the tag matches.
pub unsafe fn read_bool(prop: &SPropValue, expected_tag: u32, fallback: bool) -> bool {
    if prop.ulPropTag != expected_tag {
        return fallback;
    }
    unsafe { prop.Value.b != 0 }
}

/// Read a `PT_BINARY` property into a `Vec<u8>`, returning empty vec if the tag doesn't match.
///
/// # Safety
/// `prop.Value.bin` must contain a valid length and pointer when the tag matches.
pub unsafe fn read_binary(prop: &SPropValue, expected_tag: u32) -> Vec<u8> {
    if prop.ulPropTag != expected_tag {
        return Vec::new();
    }
    unsafe {
        let bin = &prop.Value.bin;
        std::slice::from_raw_parts(bin.lpb, bin.cb as usize).to_vec()
    }
}

/// Read a `PT_SYSTIME` property as a UTC [`chrono::DateTime`],
/// returning [`chrono::DateTime::UNIX_EPOCH`] if the tag doesn't match.
///
/// # Safety
/// `prop.Value.ft` must be valid when the tag matches.
pub unsafe fn read_filetime(
    prop: &SPropValue,
    expected_tag: u32,
) -> chrono::DateTime<chrono::Utc> {
    if prop.ulPropTag != expected_tag {
        return chrono::DateTime::UNIX_EPOCH;
    }
    filetime_to_datetime(unsafe { prop.Value.ft })
}

/// Convert a Windows `FILETIME` (100 ns intervals since 1601-01-01) to UTC [`chrono::DateTime`].
pub fn filetime_to_datetime(
    ft: windows::Win32::Foundation::FILETIME,
) -> chrono::DateTime<chrono::Utc> {
    let ticks = ((ft.dwHighDateTime as u64) << 32) | ft.dwLowDateTime as u64;
    let unix_secs = (ticks / 10_000_000).saturating_sub(11_644_473_600);
    chrono::DateTime::from_timestamp(unix_secs as i64, 0).unwrap_or_default()
}

/// Convert a UTC [`chrono::DateTime`] to a Windows `FILETIME`.
pub fn datetime_to_filetime(
    dt: chrono::DateTime<chrono::Utc>,
) -> windows::Win32::Foundation::FILETIME {
    let ticks = (dt.timestamp() as u64 + 11_644_473_600) * 10_000_000;
    windows::Win32::Foundation::FILETIME {
        dwLowDateTime:  (ticks & 0xFFFF_FFFF) as u32,
        dwHighDateTime: (ticks >> 32) as u32,
    }
}
