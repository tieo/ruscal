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

/// Path to the Outlook Extended MAPI provider DLL.
///
/// Read from `HKLM\SOFTWARE\Clients\Mail\Microsoft Outlook\DLLPathEx` at runtime.
/// Hardcoded for now; will be read from the registry in a future change.
// TODO: read from registry at startup
const MAPI_DLL_PATH: &str =
    "C:\\Program Files\\Microsoft Office\\root\\VFS\\ProgramFilesCommonX64\\system\\msmapi\\1031\\msmapi32.dll";

/// List all Outlook message stores available in the default profile.
///
/// Spawns a dedicated MAPI thread and returns the results. Use the returned
/// display names to populate a calendar picker; the selected index maps back
/// to the same list for later store identification.
pub fn list_calendar_sources() -> Result<Vec<OutlookCalendar>, MapiError> {
    std::thread::spawn(move || {
        unsafe { CoInitialize(None) }.ok().map_err(|e| MapiError(e.code().0 as u32))?;
        let result = unsafe {
            let session = session::Session::new(MAPI_DLL_PATH)?;
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
    let session = unsafe { session::Session::new(MAPI_DLL_PATH) }?;

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
