#![allow(dead_code)] // will be wired to UI once sync engine is implemented

/// Outlook calendar access via Extended MAPI.
///
/// # Threading
///
/// All MAPI work must run on a single thread that has called `CoInitialize`.
/// Use [`read_calendar_events`] which handles this internally.
pub mod calendar;
pub mod props;
pub mod session;

pub use calendar::OutlookCalendar;

use std::mem::ManuallyDrop;

use windows::Win32::System::Com::{CoInitialize, CoUninitialize};

use crate::error::MapiError;
use crate::event::CalendarEvent;

// ── Sync window defaults ──────────────────────────────────────────────────────

/// Default number of days into the past to include.
// TODO: make configurable in settings
pub const DEFAULT_PAST_DAYS: i64 = 14;

/// Default number of days into the future to include.
// TODO: make configurable in settings
pub const DEFAULT_FUTURE_DAYS: i64 = 183;

/// Resolve the path to the Outlook Extended MAPI provider DLL.
///
/// Reads `HKLM\SOFTWARE\Clients\Mail\Microsoft Outlook\DLLPathEx` first.
/// Falls back to the typical Click-to-Run install path if the registry key is
/// absent (non-English locales use a different four-digit subfolder, so the
/// fallback may not work on all systems — the registry is the authoritative source).
fn mapi_dll_path() -> String {
    read_mapi_dll_from_registry().unwrap_or_else(|| {
        "C:\\Program Files\\Microsoft Office\\root\\VFS\\ProgramFilesCommonX64\
         \\system\\msmapi\\1031\\msmapi32.dll"
            .to_owned()
    })
}

fn read_mapi_dll_from_registry() -> Option<String> {
    use windows::Win32::System::Registry::{
        RegCloseKey, RegOpenKeyExW, RegQueryValueExW,
        HKEY, HKEY_LOCAL_MACHINE, KEY_READ, REG_VALUE_TYPE,
    };
    use windows::core::PCWSTR;

    let key_path: Vec<u16> =
        "SOFTWARE\\Clients\\Mail\\Microsoft Outlook\0".encode_utf16().collect();
    let value_name: Vec<u16> = "DLLPathEx\0".encode_utf16().collect();

    unsafe {
        let mut key = HKEY::default();
        if RegOpenKeyExW(
            HKEY_LOCAL_MACHINE,
            PCWSTR::from_raw(key_path.as_ptr()),
            0,
            KEY_READ,
            &mut key,
        ).is_err() {
            return None;
        }

        let mut buf = vec![0u16; 512];
        let mut size = (buf.len() * 2) as u32;
        let mut _kind = REG_VALUE_TYPE::default();
        let ok = RegQueryValueExW(
            key,
            PCWSTR::from_raw(value_name.as_ptr()),
            None,
            Some(&mut _kind),
            Some(buf.as_mut_ptr() as *mut u8),
            Some(&mut size),
        );
        let _ = RegCloseKey(key);
        if ok.is_err() { return None; }

        // size is byte-count including null terminator; convert to char count.
        let chars = (size as usize / 2).saturating_sub(1);
        Some(String::from_utf16_lossy(&buf[..chars]))
    }
}

/// List all Outlook message stores available in the default profile.
///
/// Spawns a dedicated MAPI thread and returns the results. Use the returned
/// display names to populate a calendar picker; the selected index maps back
/// to the same list for later store identification.
pub fn list_calendar_sources() -> Result<Vec<OutlookCalendar>, MapiError> {
    std::thread::spawn(move || {
        unsafe { CoInitialize(None) }.ok().map_err(|e| MapiError(e.code().0 as u32))?;
        let dll = mapi_dll_path();
        let result = unsafe {
            let session = session::Session::new(&dll)?;
            calendar::list_calendar_stores(session.as_ptr())
        };
        unsafe { CoUninitialize() };
        result
    })
    .join()
    .map_err(|_| MapiError(0x80040106))?
}

/// Read all calendar events that fall within the given sync window.
///
/// Spawns a dedicated thread to satisfy COM/MAPI's single-thread-apartment
/// requirement, runs the MAPI query, and returns the results.
///
/// # Errors
/// Returns [`MapiError`] if MAPI initialisation, logon, or the calendar query fails.
pub fn read_calendar_events(
    window_start: chrono::DateTime<chrono::Utc>,
    window_end:   chrono::DateTime<chrono::Utc>,
) -> Result<Vec<CalendarEvent>, MapiError> {
    std::thread::spawn(move || {
        // SAFETY: CoInitialize is the first call on this thread.
        unsafe { read_on_mapi_thread(window_start, window_end) }
    })
    .join()
    .map_err(|_| MapiError(0x80040106))? // MAPI_E_CALL_FAILED if thread panicked
}

/// Inner implementation running on the dedicated MAPI thread.
///
/// # Safety
/// Must be called exactly once per thread, after `CoInitialize` and before
/// `CoUninitialize`. All COM objects must be released before `CoUninitialize`.
unsafe fn read_on_mapi_thread(
    window_start: chrono::DateTime<chrono::Utc>,
    window_end:   chrono::DateTime<chrono::Utc>,
) -> Result<Vec<CalendarEvent>, MapiError> {
    unsafe { CoInitialize(None) }.ok().map_err(|e| MapiError(e.code().0 as u32))?;

    let result = unsafe { read_calendar_inner(window_start, window_end) };

    unsafe { CoUninitialize() };
    result
}

/// Does the actual MAPI work. Separated so `CoUninitialize` always runs.
unsafe fn read_calendar_inner(
    window_start: chrono::DateTime<chrono::Utc>,
    window_end:   chrono::DateTime<chrono::Utc>,
) -> Result<Vec<CalendarEvent>, MapiError> {
    let dll = mapi_dll_path();
    let session = unsafe { session::Session::new(&dll) }?;

    let store  = unsafe { calendar::open_default_store(session.as_ptr()) }?;
    let named  = unsafe { props::resolve(&*store) }?;
    let folder = unsafe { calendar::open_calendar_folder(&*store) }?;

    let events = unsafe {
        calendar::read_events(&*store, &*folder, &named, window_start, window_end)
    }?;

    // Release COM objects before Session::drop calls MAPILogoff + MAPIUninitialize.
    drop(ManuallyDrop::into_inner(folder));
    drop(ManuallyDrop::into_inner(store));

    Ok(events)
}
