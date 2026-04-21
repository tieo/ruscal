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

/// Kill the process (if any) whose image path equals `path`.
fn terminate_at_path(path: &std::path::Path) {
    // Escape single-quotes for PowerShell single-quoted string.
    let path_str = path.to_string_lossy().replace('\'', "''");
    let _ = std::process::Command::new("powershell")
        .args([
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            &format!(
                "Get-Process | Where-Object {{ $_.Path -eq '{path_str}' }} | Stop-Process -Force"
            ),
        ])
        .output();
    // Wait for the process to fully release file locks before we copy over it.
    std::thread::sleep(std::time::Duration::from_millis(500));
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

/// Bring the existing running instance's window to the foreground.
pub fn focus_existing_window() {
    use windows::Win32::UI::WindowsAndMessaging::{
        FindWindowW, SetForegroundWindow, ShowWindow, SW_SHOW,
    };
    use windows::core::PCWSTR;

    let title: Vec<u16> = "ruscal\0".encode_utf16().collect();
    unsafe {
        let Ok(hwnd) = FindWindowW(PCWSTR::null(), PCWSTR::from_raw(title.as_ptr())) else {
            return;
        };
        if hwnd.is_invalid() { return; }
        let _ = ShowWindow(hwnd, SW_SHOW);
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
}
