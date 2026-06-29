//! Single-instance guard, focus signalling, clipboard, file picker.
//!
//! Earlier versions of this module also contained self-install, self-update,
//! and Start-menu shortcut creation. All of those were removed in the move to
//! WinGet-only distribution: WinGet handles placement, updates, and shortcuts
//! itself, and the install/update behaviours we used to do here also
//! contributed to a `Trojan:Win32/Bearfoos.A!ml` false-positive from Defender
//! (GUI exe in `%LOCALAPPDATA%` writes Run-key + .lnk + downloads new exe).

// ── Clipboard ─────────────────────────────────────────────────────────────────

/// Copy `text` onto the system clipboard as `CF_UNICODETEXT`.
pub fn set_clipboard_text(text: &str) {
    use windows::Win32::Foundation::{HANDLE, HGLOBAL, HWND};
    use windows::Win32::System::DataExchange::{
        CloseClipboard, EmptyClipboard, OpenClipboard, SetClipboardData,
    };
    use windows::Win32::System::Memory::{GlobalAlloc, GlobalLock, GlobalUnlock, GMEM_MOVEABLE};
    use windows::Win32::System::Ole::CF_UNICODETEXT;

    let wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
    let bytes = wide.len() * 2;

    unsafe {
        if OpenClipboard(HWND::default()).is_err() { return; }
        let _ = EmptyClipboard();

        let Ok(hmem): windows::core::Result<HGLOBAL> = GlobalAlloc(GMEM_MOVEABLE, bytes) else {
            let _ = CloseClipboard();
            return;
        };
        let dst = GlobalLock(hmem) as *mut u16;
        if !dst.is_null() {
            std::ptr::copy_nonoverlapping(wide.as_ptr(), dst, wide.len());
            let _ = GlobalUnlock(hmem);
            let _ = SetClipboardData(u32::from(CF_UNICODETEXT.0), HANDLE(hmem.0));
        }
        let _ = CloseClipboard();
    }
}

// ── Native file-open dialog ───────────────────────────────────────────────────

/// Show a system Open-File dialog filtered to `*.exe`, return the chosen path.
pub fn pick_exe_path(title: &str) -> Option<std::path::PathBuf> {
    use windows::Win32::System::Com::{
        CoCreateInstance, CoInitializeEx, CoUninitialize,
        CLSCTX_INPROC_SERVER, COINIT_APARTMENTTHREADED,
    };
    use windows::Win32::UI::Shell::Common::COMDLG_FILTERSPEC;
    use windows::Win32::UI::Shell::{
        FileOpenDialog, IFileOpenDialog, IShellItem, SIGDN_FILESYSPATH,
    };
    use windows::core::PCWSTR;

    let title_w:  Vec<u16> = title.encode_utf16().chain(std::iter::once(0)).collect();
    let label_w:  Vec<u16> = "Executables (*.exe)\0".encode_utf16().collect();
    let spec_w:   Vec<u16> = "*.exe\0".encode_utf16().collect();

    let mut result: Option<std::path::PathBuf> = None;
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        let _ = (|| -> windows::core::Result<()> {
            let dlg: IFileOpenDialog =
                CoCreateInstance(&FileOpenDialog, None, CLSCTX_INPROC_SERVER)?;
            let filter = [COMDLG_FILTERSPEC {
                pszName: PCWSTR(label_w.as_ptr()),
                pszSpec: PCWSTR(spec_w.as_ptr()),
            }];
            dlg.SetFileTypes(&filter)?;
            dlg.SetTitle(PCWSTR(title_w.as_ptr()))?;
            if dlg.Show(None).is_ok() {
                let item: IShellItem = dlg.GetResult()?;
                let pwstr = item.GetDisplayName(SIGDN_FILESYSPATH)?;
                let s = pwstr.to_string()?;
                windows::Win32::System::Com::CoTaskMemFree(Some(pwstr.0 as *const _));
                result = Some(std::path::PathBuf::from(s));
            }
            Ok(())
        })();
        CoUninitialize();
    }
    result
}

// ── Single-instance guard ─────────────────────────────────────────────────────

use windows::Win32::Foundation::HANDLE;

/// RAII guard that holds the single-instance named mutex for the lifetime of
/// the process. The OS releases the mutex automatically on process exit.
pub struct SingleInstanceGuard(HANDLE);

// SAFETY: HANDLE is a plain integer; we never share it across threads.
unsafe impl Send for SingleInstanceGuard {}

impl Drop for SingleInstanceGuard {
    fn drop(&mut self) {
        unsafe {
            let _ = windows::Win32::Foundation::CloseHandle(self.0);
        }
    }
}

/// Attempt to acquire the single-instance mutex.
///
/// Returns `Some(guard)` if this is the only running instance.
/// Returns `None` if another instance already holds the mutex.
pub fn acquire_single_instance() -> Option<SingleInstanceGuard> {
    use windows::Win32::Foundation::{GetLastError, ERROR_ALREADY_EXISTS};
    use windows::Win32::System::Threading::CreateMutexW;
    use windows::core::PCWSTR;

    let name: Vec<u16> = "Local\\ruscal_single_instance\0"
        .encode_utf16()
        .collect();
    unsafe {
        let handle = CreateMutexW(None, true, PCWSTR::from_raw(name.as_ptr())).ok()?;
        if GetLastError() == ERROR_ALREADY_EXISTS {
            let _ = windows::Win32::Foundation::CloseHandle(handle);
            return None;
        }
        Some(SingleInstanceGuard(handle))
    }
}

// ── Cross-process focus signalling ───────────────────────────────────────────
//
// A second launch must bring the first instance's window to the front.
// The earlier approach — `FindWindowW` + `ShowWindow(SW_SHOW)` from the second
// process — worked visually but broke Slint's internal "window visible" state:
// Slint did not know the window had been un-hidden behind its back, so a later
// `p.hide()` (our minimize-to-tray handler) silently no-op'd and left the
// taskbar entry stuck on screen.
//
// Instead, a named Win32 event is used as a wake-up channel. The primary
// instance creates it, blocks a worker thread on `WaitForSingleObject`, and
// handles each wake by calling `p.show()` through `invoke_from_event_loop` so
// Slint's bookkeeping stays coherent. The secondary instance opens the event,
// calls `SetEvent`, and exits — never touches the other process's window.

const FOCUS_EVENT_NAME: &str = "Local\\ruscal_focus_request";

fn wide_event_name(name: &str) -> Vec<u16> {
    name.encode_utf16().chain(std::iter::once(0)).collect()
}

/// Create the focus-request event and spawn a thread that waits on it,
/// invoking `on_request` each time a second instance signals.
///
/// `on_request` runs on a background thread and should marshal any Slint
/// work onto the event loop via `slint::invoke_from_event_loop`. The returned
/// handle must be kept alive for the process lifetime; dropping it closes
/// the event and the waiting thread exits on the next signal.
pub fn listen_for_focus_requests<F>(on_request: F) -> Option<FocusListener>
where
    F: Fn() + Send + 'static,
{
    listen_for_focus_requests_named(FOCUS_EVENT_NAME, on_request)
}

/// Like [`listen_for_focus_requests`] but with a caller-supplied event name.
/// Exists so the integration test can run on a unique name and not collide
/// with the live primary instance whose listener would otherwise consume
/// the test's signals (auto-reset events deliver each `SetEvent` to exactly
/// one waiter, so a parallel listener is fatal).
fn listen_for_focus_requests_named<F>(name: &str, on_request: F) -> Option<FocusListener>
where
    F: Fn() + Send + 'static,
{
    use windows::Win32::Foundation::HANDLE;
    use windows::Win32::System::Threading::CreateEventW;
    use windows::core::PCWSTR;

    let wide = wide_event_name(name);
    let handle: HANDLE = unsafe {
        // Manual-reset=false → auto-reset after each wait returns. No security
        // attrs: the event lives in the Local\ namespace so only the current
        // logon session can open it.
        CreateEventW(None, false, false, PCWSTR::from_raw(wide.as_ptr())).ok()?
    };

    // Cross the thread boundary as an integer. `HANDLE` wraps `*mut c_void`
    // and Rust's disjoint-capture in closures would otherwise try to move the
    // raw pointer itself (not Send) even if the wrapping struct was marked
    // Send. `isize` is trivially Send and round-trips through `HANDLE(...)`.
    let handle_raw: isize = handle.0 as isize;
    std::thread::spawn(move || {
        use windows::Win32::System::Threading::WaitForSingleObject;
        use windows::Win32::Foundation::{HANDLE, WAIT_OBJECT_0};
        let waited = HANDLE(handle_raw as *mut _);
        loop {
            let status = unsafe { WaitForSingleObject(waited, u32::MAX) };
            if status != WAIT_OBJECT_0 { break; }
            on_request();
        }
    });

    Some(FocusListener(handle))
}

/// Opaque token representing the primary instance's focus-request listener.
/// Held for the process lifetime; closed on drop.
pub struct FocusListener(windows::Win32::Foundation::HANDLE);

// SAFETY: HANDLE wraps a pointer but its only use after Drop is CloseHandle,
// which is thread-safe. The listener holds a copy for WaitForSingleObject.
unsafe impl Send for FocusListener {}

impl Drop for FocusListener {
    fn drop(&mut self) {
        unsafe { let _ = windows::Win32::Foundation::CloseHandle(self.0); }
    }
}

/// Signal the primary instance to bring its window to the foreground.
/// Called only by a second instance that has already lost the single-instance
/// race — this function never touches windows directly, so Slint's state in
/// the primary instance stays consistent.
pub fn signal_focus_request() {
    signal_focus_request_named(FOCUS_EVENT_NAME);
}

fn signal_focus_request_named(name: &str) {
    use windows::Win32::System::Threading::{OpenEventW, SetEvent, EVENT_MODIFY_STATE};
    use windows::core::PCWSTR;

    let wide = wide_event_name(name);
    unsafe {
        let Ok(handle) = OpenEventW(EVENT_MODIFY_STATE, false, PCWSTR::from_raw(wide.as_ptr()))
        else { return };
        let _ = SetEvent(handle);
        let _ = windows::Win32::Foundation::CloseHandle(handle);
    }
}

/// Raise the ruscal window to the foreground. Must be called from the
/// primary instance's *own* event loop, after `p.show()` — so Slint knows
/// the window is visible and our Win32 calls don't desync its bookkeeping.
/// Uses `FindWindowW` in-process: the title is under our control and the
/// extra Win32 lookup is cheaper than threading the HWND through Slint's
/// raw-window-handle feature gate.
pub fn bring_own_window_forward() {
    use windows::Win32::UI::WindowsAndMessaging::{
        FindWindowW, SetForegroundWindow, ShowWindow, SW_RESTORE, IsIconic,
    };
    use windows::core::PCWSTR;

    let title: Vec<u16> = "ruscal\0".encode_utf16().collect();
    unsafe {
        let Ok(hwnd) = FindWindowW(PCWSTR::null(), PCWSTR::from_raw(title.as_ptr())) else {
            return;
        };
        if hwnd.is_invalid() { return; }
        if IsIconic(hwnd).as_bool() {
            let _ = ShowWindow(hwnd, SW_RESTORE);
        }
        let _ = SetForegroundWindow(hwnd);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    // ── Cross-instance focus signalling ────────────────────────────────────
    //
    // Regression guard: earlier versions reached across process boundaries
    // with raw `ShowWindow` calls, which desynced Slint's internal visibility
    // state and left a stuck taskbar entry after minimize-to-tray. The fix
    // is a named event: secondary signals, primary listens and surfaces the
    // window through its own event loop. Verify the end-to-end delivery.

    #[cfg(target_os = "windows")]
    #[test]
    fn signal_reaches_listener() {
        use std::sync::Arc;
        use std::sync::atomic::{AtomicU32, Ordering};

        // Use a per-test unique event name so this test never collides
        // with a live ruscal instance running on the same machine. The
        // auto-reset event delivers each SetEvent to exactly one waiter,
        // so a parallel listener (the real app) would otherwise eat our
        // signals and fail the test non-deterministically.
        let event_name = format!(
            "Local\\ruscal_focus_request_test_{}_{}",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos(),
        );

        let fired = Arc::new(AtomicU32::new(0));
        let fired_cb = Arc::clone(&fired);
        let _listener = super::listen_for_focus_requests_named(&event_name, move || {
            fired_cb.fetch_add(1, Ordering::SeqCst);
        }).expect("create focus-request event");

        // Signal once and wait for the wake-up. Then signal again and wait
        // for the next. Back-to-back SetEvents can debounce on an
        // already-signaled event — desirable in production (bursts collapse
        // to a single window-focus action), but we need to acknowledge each
        // signal before sending the next so the test stays deterministic.

        super::signal_focus_request_named(&event_name);
        wait_for_count(&fired, 1);
        super::signal_focus_request_named(&event_name);
        wait_for_count(&fired, 2);

        assert_eq!(fired.load(Ordering::SeqCst), 2);
    }

    #[cfg(target_os = "windows")]
    fn wait_for_count(counter: &std::sync::Arc<std::sync::atomic::AtomicU32>, target: u32) {
        use std::sync::atomic::Ordering;
        let deadline = std::time::Instant::now()
            + std::time::Duration::from_millis(500);
        while counter.load(Ordering::SeqCst) < target
            && std::time::Instant::now() < deadline
        {
            std::thread::sleep(std::time::Duration::from_millis(2));
        }
        assert!(
            counter.load(Ordering::SeqCst) >= target,
            "wake-up never arrived (expected >= {target}, got {})",
            counter.load(Ordering::SeqCst)
        );
    }
}
