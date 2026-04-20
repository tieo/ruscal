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
    let latest = json["tag_name"].as_str()?.trim_start_matches('v').to_owned();

    // Strip any git-describe suffix (e.g. "1.0.3-2-gabcdef" → "1.0.3")
    let current_clean = current_version.trim_start_matches('v')
        .split('-').next().unwrap_or(current_version);

    if semver_gt(&latest, current_clean) {
        Some(latest)
    } else {
        None
    }
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
