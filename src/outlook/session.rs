/// MAPI session lifecycle: initialise, logon, logoff, uninitialise.
///
/// We load `msmapi32.dll` directly rather than the `mapi32.dll` stub.
/// The stub's `MAPILogonEx` returns a Simple MAPI `LHANDLE` (an integer),
/// not an Extended MAPI `IMAPISession*` COM pointer. The real provider DLL
/// path is stored in:
/// `HKLM\SOFTWARE\Clients\Mail\Microsoft Outlook\DLLPathEx`
use std::ffi::c_void;

use crate::error::{check_hr, MapiError};
use windows::Win32::System::Mapi::MAPI_EXTENDED;

// ── MAPI function type aliases ────────────────────────────────────────────────

type MAPIInitializeFn   = unsafe extern "system" fn(*const c_void) -> i32;
type MAPIUninitializeFn = unsafe extern "system" fn();
type MAPILogonExFn      = unsafe extern "system" fn(
    uluiparam:       usize,
    lpszprofilename: *const u8,
    lpszpassword:    *const u8,
    flflags:         u32,
    lppses:          *mut *mut c_void,
) -> i32;
type MAPILogoffFn = unsafe extern "system" fn(
    lhsession:  *mut c_void,
    uluiparam:  usize,
    flflags:    u32,
    ulreserved: u32,
) -> i32;

// ── IMAPISession vtable ───────────────────────────────────────────────────────

/// Partial vtable for `IMAPISession`.
///
/// Only the entries we call are typed precisely. The rest are `usize` padding —
/// same size as a pointer — so the vtable offsets remain correct.
///
/// Vtable order per the MAPI SDK (`Mapix.h`):
/// 0 QueryInterface, 1 AddRef, 2 Release, 3 GetLastError,
/// 4 GetMsgStoresTable, 5 OpenMsgStore, …
#[repr(C)]
pub struct IMAPISessionVtbl {
    pub query_interface:      usize,
    pub add_ref:              usize,
    pub release:              usize,
    pub get_last_error:       usize,
    pub get_msg_stores_table: unsafe extern "system" fn(
        this:      *mut c_void,
        ul_flags:  u32,
        lpp_table: *mut *mut c_void,
    ) -> i32,
    pub open_msg_store: unsafe extern "system" fn(
        this:          *mut c_void,
        ul_ui_param:   usize,
        cb_entry_id:   u32,
        lp_entry_id:   *const c_void,
        lp_interface:  *const c_void,
        ul_flags:      u32,
        lpp_msg_store: *mut *mut c_void,
    ) -> i32,
}

/// A COM object whose first field is a pointer to its vtable.
/// This is the universal layout of every COM interface.
#[repr(C)]
pub struct IMAPISession {
    pub vtbl: *const IMAPISessionVtbl,
}

// ── Session ───────────────────────────────────────────────────────────────────

/// An active MAPI session, holding the loaded provider DLL and raw session pointer.
///
/// Drop this value to log off and unload the provider.
pub struct Session {
    lib:         libloading::Library,
    raw_session: *mut c_void,
}

impl Session {
    /// Load the Outlook MAPI provider, initialise the subsystem, and log on.
    ///
    /// # Errors
    /// Returns [`MapiError`] if any MAPI call fails.
    ///
    /// # Safety
    /// Must be called from a thread that has already called `CoInitialize`.
    /// The caller is responsible for keeping the returned `Session` alive
    /// for as long as any MAPI objects derived from it are in use.
    pub unsafe fn new(dll_path: &str) -> Result<Self, MapiError> {
        // SAFETY: libloading::Library::new is unsafe because loading a DLL
        // executes its DllMain, which may run arbitrary code.
        let lib = unsafe { libloading::Library::new(dll_path) }
            .map_err(|_| MapiError(0x80040101))?; // MAPI_E_DISK_ERROR as stand-in

        // SAFETY: lib.get is unsafe because the symbol may have any type;
        // we constrain it via the type alias.
        let initialize: libloading::Symbol<MAPIInitializeFn> =
            unsafe { lib.get(b"MAPIInitialize\0") }.map_err(|_| MapiError(0x80040108))?;
        check_hr(unsafe { initialize(core::ptr::null()) })?;

        let logon: libloading::Symbol<MAPILogonExFn> =
            unsafe { lib.get(b"MAPILogonEx\0") }.map_err(|_| MapiError(0x80040108))?;

        // MAPI_USE_DEFAULT (0x40) — use the default Outlook profile, no UI prompt.
        // MAPI_NO_MAIL    (0x8000) — don't start the spooler; we only read data.
        const MAPI_USE_DEFAULT: u32 = 0x00000040;
        const MAPI_NO_MAIL: u32     = 0x00008000;

        let mut raw_session: *mut c_void = core::ptr::null_mut();
        check_hr(unsafe {
            logon(
                0,
                core::ptr::null(),
                core::ptr::null(),
                MAPI_EXTENDED | MAPI_USE_DEFAULT | MAPI_NO_MAIL,
                &mut raw_session,
            )
        })?;

        Ok(Self { lib, raw_session })
    }

    /// The raw `IMAPISession` pointer.
    pub fn as_ptr(&self) -> *mut IMAPISession {
        self.raw_session as *mut IMAPISession
    }
}

impl Drop for Session {
    fn drop(&mut self) {
        // Log off and uninitialise. Errors here are non-recoverable, so we log
        // and continue rather than panicking in a destructor.
        unsafe {
            if let Ok(logoff) = self.lib.get::<MAPILogoffFn>(b"MAPILogoff\0") {
                let _ = logoff(self.raw_session, 0, 0, 0);
            }
            if let Ok(uninit) = self.lib.get::<MAPIUninitializeFn>(b"MAPIUninitialize\0") {
                uninit();
            }
        }
    }
}
