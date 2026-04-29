//! Self-install, single-instance guard, and auto-update.

use std::path::PathBuf;

// ── CLI args ──────────────────────────────────────────────────────────────────

#[derive(Default)]
pub struct LaunchArgs {
    /// Launched after a fresh install from an outside path.
    pub just_installed: bool,
    /// Launched after an in-app update; value is the *previous* version string.
    pub just_updated: Option<String>,
}

pub fn parse_args() -> LaunchArgs {
    let mut args = LaunchArgs::default();
    for arg in std::env::args().skip(1) {
        if arg == "--just-installed" {
            args.just_installed = true;
        } else if let Some(ver) = arg.strip_prefix("--just-updated=") {
            args.just_updated = Some(ver.to_owned());
        }
    }
    args
}

// ── Installed path ────────────────────────────────────────────────────────────

/// `%LOCALAPPDATA%\ruscal\ruscal.exe`
pub fn installed_path() -> Option<PathBuf> {
    dirs::data_local_dir().map(|d| d.join("ruscal").join("ruscal.exe"))
}

/// True if this process is already running from the canonical install location.
pub fn is_installed_path() -> bool {
    let Ok(current) = std::env::current_exe() else { return false };
    let Some(target) = installed_path() else { return false };
    // Case-insensitive: Windows paths are case-insensitive.
    current
        .to_string_lossy()
        .eq_ignore_ascii_case(&target.to_string_lossy())
}

// ── Self-install / self-update ────────────────────────────────────────────────

/// Copy this executable to the installed path, terminate any existing instance
/// at that path, relaunch from there, then exit this process.
///
/// Never returns — always exits.
pub fn self_install(flag: Option<&str>) -> ! {
    let installed = installed_path().expect("cannot determine install path");

    if let Some(parent) = installed.parent() {
        let _ = std::fs::create_dir_all(parent);
    }

    // Kill any running instance at the install location so the file is not locked.
    terminate_at_path(&installed);

    let current = std::env::current_exe().expect("current_exe");
    std::fs::copy(&current, &installed).expect("failed to copy to install path");

    // Make ruscal searchable in the Start menu. Idempotent — does nothing
    // when a shortcut is already present from a previous install.
    ensure_start_menu_shortcut();

    // Clean up the update temp file if this process IS the update temp file.
    // We've already copied ourselves to the installed path, so we can delete ourselves.
    if current.file_name().map(|n| n == "ruscal_update.exe").unwrap_or(false) {
        let _ = std::fs::remove_file(&current);
    }

    let mut cmd = std::process::Command::new(&installed);
    if let Some(f) = flag {
        cmd.arg(f);
    }
    // CREATE_NO_WINDOW prevents a console flash when spawning the installed GUI exe.
    #[cfg(target_os = "windows")]
    {
        use std::os::windows::process::CommandExt;
        cmd.creation_flags(0x08000000); // CREATE_NO_WINDOW
    }
    cmd.spawn().expect("failed to launch installed exe");

    std::process::exit(0);
}

// ── Start-menu shortcut ───────────────────────────────────────────────────────

/// Path of the per-user Start-menu `.lnk` for ruscal.
///
/// Windows Search indexes the `Programs` folder; without a `.lnk` in there,
/// ruscal doesn't appear in Start-menu search regardless of how it was
/// installed. Using `%APPDATA%\...\Start Menu\Programs` means the shortcut
/// lives with the current user and doesn't require elevation.
pub fn start_menu_shortcut_path() -> Option<PathBuf> {
    dirs::data_dir().map(|appdata| {
        appdata
            .join("Microsoft")
            .join("Windows")
            .join("Start Menu")
            .join("Programs")
            .join("ruscal.lnk")
    })
}

/// Create the Start-menu shortcut if it doesn't exist yet, pointing at the
/// installed exe. Idempotent — a no-op when the shortcut is already present.
/// Called from `self_install` (so a fresh install is immediately searchable)
/// and from release-build startup (so existing installs self-heal the first
/// time they run a version that ships with this function).
pub fn ensure_start_menu_shortcut() {
    let Some(lnk) = start_menu_shortcut_path() else { return; };
    if lnk.exists() { return; }
    let Some(target) = installed_path() else { return; };
    if !target.exists() { return; }
    write_start_menu_shortcut(&lnk, &target);
}

/// Spawn PowerShell to materialise the `.lnk`. We use the shell's scripting
/// host because the only alternative is the `IShellLinkW` COM dance (several
/// hundred lines of `windows` crate boilerplate for a one-shot write). Paths
/// are quoted as PowerShell single-quoted strings (`'` doubled for escape).
fn write_start_menu_shortcut(lnk: &std::path::Path, target: &std::path::Path) {
    if let Some(parent) = lnk.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let q = |p: &std::path::Path| p.to_string_lossy().replace('\'', "''");
    let lnk_q    = q(lnk);
    let target_q = q(target);
    let working  = target.parent().unwrap_or(target);
    let working_q = q(working);

    let script = format!(
        "$s = New-Object -ComObject WScript.Shell; \
         $l = $s.CreateShortcut('{lnk_q}'); \
         $l.TargetPath = '{target_q}'; \
         $l.WorkingDirectory = '{working_q}'; \
         $l.IconLocation = '{target_q},0'; \
         $l.Description = 'ruscal — Outlook to Google Calendar sync'; \
         $l.Save()"
    );

    let mut cmd = std::process::Command::new("powershell");
    cmd.args(["-NoProfile", "-NonInteractive", "-Command", &script]);
    hide_console_window(&mut cmd);
    let _ = cmd.output();
}

/// Kill the process (if any) whose image path equals `path`.
fn terminate_at_path(path: &std::path::Path) {
    // Escape single-quotes for PowerShell single-quoted string.
    let path_str = path.to_string_lossy().replace('\'', "''");
    let mut cmd = std::process::Command::new("powershell");
    cmd.args([
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        &format!(
            "Get-Process | Where-Object {{ $_.Path -eq '{path_str}' }} | Stop-Process -Force"
        ),
    ]);
    hide_console_window(&mut cmd);
    let _ = cmd.output();
    // Wait for the process to fully release file locks before we copy over it.
    std::thread::sleep(std::time::Duration::from_millis(500));
}

/// Mark a `Command` so Windows does not allocate a console window for it.
///
/// ruscal is compiled with `windows_subsystem = "windows"` — it has no
/// attached console. When such a GUI process spawns a console-subsystem
/// child (powershell is one), Windows allocates a *new* console for the
/// child. That window pops visibly on the user's screen and lingers for
/// the child's lifetime. `CREATE_NO_WINDOW` suppresses it. Dialog windows
/// raised *from* the child (e.g. `OpenFileDialog`) still appear normally.
#[cfg(target_os = "windows")]
pub fn hide_console_window(cmd: &mut std::process::Command) {
    use std::os::windows::process::CommandExt;
    cmd.creation_flags(0x08000000); // CREATE_NO_WINDOW
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
// Instead, we use a named Win32 event as a wake-up channel. The primary
// instance creates it, blocks a worker thread on `WaitForSingleObject`, and
// handles each wake by calling `p.show()` through `invoke_from_event_loop` so
// Slint's bookkeeping stays coherent. The secondary instance opens the event,
// calls `SetEvent`, and exits — it never touches the other process's window.

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

// ── Update check ──────────────────────────────────────────────────────────────

/// Poll the GitHub Releases API.  Returns `Some(version)` if a newer release
/// exists, `None` if up-to-date or the check fails (network errors are silent).
pub fn check_for_update(current_version: &str) -> Option<String> {
    let resp = reqwest::blocking::Client::new()
        .get("https://api.github.com/repos/tieo/ruscal/releases/latest")
        .header("User-Agent", format!("ruscal/{current_version}"))
        .timeout(std::time::Duration::from_secs(8))
        .send()
        .ok()?;

    if !resp.status().is_success() {
        return None;
    }

    let json: serde_json::Value = resp.json().ok()?;
    let latest = json["tag_name"].as_str()?;

    decide_update(current_version, read_installed_version().as_deref(), latest)
}

/// Decide whether to offer an update — pure, no I/O. Separated from
/// [`check_for_update`] so the decision logic is unit-testable.
///
/// `installed` is the version the installed (production) binary at
/// `%LOCALAPPDATA%\ruscal\ruscal.exe` last reported on startup — that's the
/// copy the self-updater will actually replace, not the running exe. When the
/// developer runs a dev build from `cargo run`, comparing against the dev
/// version would misleadingly report "up to date" even though the cached
/// production copy is stale. We compare against the installed version instead.
///
/// Returns `Some(version_without_v)` if an update should be offered.
fn decide_update(current: &str, installed: Option<&str>, latest: &str) -> Option<String> {
    // Dev builds contain unpublished work — the running exe isn't the one
    // the self-updater touches. Fall back to the recorded installed version;
    // if there isn't one (user never launched the installed exe) we have no
    // meaningful comparison, so stay silent.
    let effective = if is_dev_build(current) {
        match installed {
            Some(v) if !is_dev_build(v) => v,
            _ => return None,
        }
    } else {
        current
    };

    let latest_clean    = latest.trim_start_matches('v');
    let effective_clean = effective.trim_start_matches('v');

    if semver_gt(latest_clean, effective_clean) {
        Some(latest_clean.to_owned())
    } else {
        None
    }
}

/// Path of the sidecar file recording the version of the installed exe.
fn installed_version_path() -> Option<PathBuf> {
    dirs::data_local_dir().map(|d| d.join("ruscal").join(".installed_version"))
}

/// Read the version string last recorded by the installed binary at startup.
/// Returns `None` if the sidecar doesn't exist or is unreadable.
pub fn read_installed_version() -> Option<String> {
    let content = std::fs::read_to_string(installed_version_path()?).ok()?;
    let trimmed = content.trim();
    if trimmed.is_empty() { None } else { Some(trimmed.to_owned()) }
}

/// Stamp the installed version sidecar. Call on startup when the running exe
/// IS the installed binary — so later dev runs can still tell whether the
/// cached production copy is out of date.
pub fn record_installed_version(version: &str) {
    let Some(path) = installed_version_path() else { return };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    let _ = std::fs::write(path, version);
}

/// Is this version string a dev build?
///
/// Shapes produced by `git describe --tags --always --dirty` (our build.rs):
/// ```text
/// v1.2.3                    — clean tagged release       (NOT dev)
/// v1.2.3-dirty              — tagged + local edits       (dev)
/// v1.2.3-5-gabc1234         — 5 commits past v1.2.3      (dev)
/// v1.2.3-5-gabc1234-dirty   — plus uncommitted edits     (dev)
/// dev                       — `git_version!` fallback    (dev)
/// ```
fn is_dev_build(version: &str) -> bool {
    if version == "dev" || version.ends_with("-dirty") {
        return true;
    }
    // Detect the `-N-gSHA` chunk that git describe appends once there are
    // commits past the nearest matching tag: between the tag and the SHA,
    // the middle hyphen-part is a pure digit count ("5") and the SHA chunk
    // starts with 'g'.
    let mut parts = version.split('-');
    let _tag   = parts.next();
    let middle = parts.next();
    let sha    = parts.next();
    matches!(
        (middle, sha),
        (Some(n), Some(s))
            if !n.is_empty()
            && n.chars().all(|c| c.is_ascii_digit())
            && s.starts_with('g')
    )
}

/// Download the release asset for `version` to a temp file beside the install
/// path.  Returns the path on success.
pub fn download_update(version: &str) -> Result<PathBuf, String> {
    let url = format!(
        "https://github.com/tieo/ruscal/releases/download/v{version}/ruscal.exe"
    );

    let temp = dirs::data_local_dir()
        .map(|d| d.join("ruscal").join("ruscal_update.exe"))
        .ok_or_else(|| "cannot locate local data directory".to_owned())?;

    if let Some(parent) = temp.parent() {
        let _ = std::fs::create_dir_all(parent);
    }

    let bytes = reqwest::blocking::get(&url)
        .map_err(|e| e.to_string())?
        .bytes()
        .map_err(|e| e.to_string())?;

    std::fs::write(&temp, &bytes).map_err(|e| e.to_string())?;
    Ok(temp)
}

/// Remove a leftover `ruscal_update.exe` from a prior self-install run.
///
/// The self-install path already deletes the temp when it's the running exe,
/// but if the user later launches ruscal a different way (explorer, autostart)
/// the stale temp can linger. Call this on startup — no-op if absent or if
/// this process happens to be running from that path.
pub fn cleanup_stale_update_exe() {
    let Some(temp) = dirs::data_local_dir()
        .map(|d| d.join("ruscal").join("ruscal_update.exe")) else { return };
    if !temp.exists() { return; }
    if let Ok(cur) = std::env::current_exe() {
        if cur == temp { return; }
    }
    let _ = std::fs::remove_file(&temp);
}

fn semver_gt(a: &str, b: &str) -> bool {
    fn parse(s: &str) -> (u32, u32, u32) {
        let mut it = s.splitn(3, '.').map(|p| p.parse::<u32>().unwrap_or(0));
        (it.next().unwrap_or(0), it.next().unwrap_or(0), it.next().unwrap_or(0))
    }
    parse(a) > parse(b)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn is_dev_build_recognises_every_git_describe_shape() {
        // Clean release tags — not dev.
        assert!(!is_dev_build("v1.0.8"));
        assert!(!is_dev_build("1.0.8"));

        // Dev shapes.
        assert!(is_dev_build("dev"),                      "fallback");
        assert!(is_dev_build("v1.0.8-dirty"),             "tag + dirty");
        assert!(is_dev_build("v1.0.8-1-g1342d9d"),        "commits past tag");
        assert!(is_dev_build("v1.0.8-5-gabc1234-dirty"),  "commits + dirty");

        // Adversarial shapes that must NOT be misread as dev builds.
        assert!(!is_dev_build("v1.0.8-beta"),        "pre-release tag, not commits-past");
        assert!(!is_dev_build("v1.0.8-rc.1"),        "pre-release tag with dot");
    }

    /// Dev builds with no recorded installed version stay silent: nothing
    /// actionable to report (the running exe isn't what would get updated,
    /// and there's no production copy on this machine to compare against).
    #[test]
    fn decide_update_silent_for_dev_build_without_installed_record() {
        assert_eq!(decide_update("v1.0.7-3-gabc",     None, "v1.0.8"), None);
        assert_eq!(decide_update("v1.0.8-1-g1342d9d", None, "v1.0.9"), None);
        assert_eq!(decide_update("v1.0.8-dirty",      None, "v1.0.9"), None);
        assert_eq!(decide_update("dev",               None, "v1.0.8"), None);
    }

    /// The regression that motivated this round: developer runs the dev
    /// build from `cargo run` and the production copy in %LOCALAPPDATA% is
    /// stale. Check-for-updates must surface the update so they know to
    /// refresh the cached exe, even though the running dev exe is silent.
    #[test]
    fn decide_update_compares_installed_when_running_dev_build() {
        // Cached production copy is v1.0.8, latest release is v1.0.9 → offer.
        assert_eq!(
            decide_update("v1.0.9-1-gabc", Some("v1.0.8"), "v1.0.9"),
            Some("1.0.9".into()),
        );
        // Cached == latest → silent even on dev build.
        assert_eq!(
            decide_update("v1.0.9-1-gabc", Some("v1.0.9"), "v1.0.9"),
            None,
        );
    }

    /// Pathological case: the sidecar somehow holds a dev string. Treat
    /// it as "no useful record" rather than comparing dev-vs-latest.
    #[test]
    fn decide_update_ignores_dev_installed_record() {
        assert_eq!(
            decide_update("v1.0.9-1-gabc", Some("v1.0.8-dirty"), "v1.0.9"),
            None,
        );
    }

    #[test]
    fn decide_update_offers_newer_releases_to_clean_builds() {
        assert_eq!(decide_update("v1.0.7", None, "v1.0.8"), Some("1.0.8".into()));
        assert_eq!(decide_update("v1.0.8", None, "v1.0.9"), Some("1.0.9".into()));
        assert_eq!(decide_update("v0.9.0", None, "v1.0.0"), Some("1.0.0".into()));
    }

    /// On a clean build, the running exe *is* the installed exe (self-install
    /// copied it), so the installed-version record is redundant — current
    /// wins. Asserts we don't accidentally let the sidecar override a clean
    /// current (e.g. sidecar lagging behind after a self-install).
    #[test]
    fn decide_update_clean_build_ignores_installed_record() {
        assert_eq!(
            decide_update("v1.0.9", Some("v1.0.7"), "v1.0.9"),
            None,
            "clean current on latest: installed record must not drag us back",
        );
    }

    #[test]
    fn decide_update_returns_none_when_up_to_date_or_ahead() {
        assert_eq!(decide_update("v1.0.8", None, "v1.0.8"), None); // exactly on latest
        assert_eq!(decide_update("v1.0.9", None, "v1.0.8"), None); // ahead of latest
    }

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
        // for the next. We can't rely on two back-to-back SetEvents
        // producing two wake-ups — the auto-reset event *debounces* them
        // (SetEvent on an already-signaled event is a no-op), which is
        // desirable in production (bursts collapse to a single window-focus
        // action). What matters is that each wait-acknowledged signal is
        // then followed by servicing of the next.

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
