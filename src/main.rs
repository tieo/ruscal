#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod caldav;
mod error;
mod event;
mod google;
mod outlook;
mod state;
mod sync;
mod updater;

/// Embedded version from `git describe`. Shapes:
/// * `v1.2.3`               → clean build on a tagged release commit
/// * `v1.2.3-5-gabc1234`    → 5 commits past the last tag (dev build)
/// * `v1.2.3-5-gabc1234-dirty` → plus uncommitted changes (local hack)
/// * `dev`                  → not a git checkout at build time
const APP_VERSION: &str = git_version::git_version!(
    args = ["--tags", "--match", "v*", "--always", "--dirty"],
    prefix = "",
    fallback = "dev",
);

slint::include_modules!();

/// Wizard/page indices shared with `ui/app.slint`'s `current-page` property.
/// Slint's `int` maps to `i32` in the generated bindings, so these constants
/// match that type exactly and are usable directly with
/// `set_current_page` / `get_current_page` — no casts at the call sites.
mod page {
    pub const PAIRS:        i32 = 0;
    pub const SOURCE_PICK:  i32 = 1;
    pub const DEST_PICK:    i32 = 2;
    pub const GOOGLE_PICK:  i32 = 3;
    pub const SETTINGS:     i32 = 4;

    /// The settings cog toggles between the settings page and the pairs list.
    /// Extracted so the rule is unit-testable without a Slint window.
    pub fn toggle_settings(current: i32) -> i32 {
        if current == SETTINGS { PAIRS } else { SETTINGS }
    }
}

use std::rc::Rc;
use std::sync::{Arc, Mutex};
use slint::{Model, VecModel};
use tray_icon::{Icon, MouseButton, TrayIconBuilder, TrayIconEvent};
use tray_icon::menu::{Menu, MenuEvent, MenuItem, PredefinedMenuItem};
use windows::Win32::Foundation::RECT;
use windows::Win32::UI::HiDpi::{
    SetProcessDpiAwarenessContext, DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2,
};
use windows::Win32::UI::WindowsAndMessaging::{
    SystemParametersInfoW, SPI_GETWORKAREA,
    SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS,
};

// ── Persistence ───────────────────────────────────────────────────────────────

#[derive(serde::Serialize, serde::Deserialize, Clone, Default)]
struct SavedPair {
    source_account: String,
    dest_account:   String,
    dest_id:        String,
    #[serde(default)]
    google_email:   String,
}

fn default_interval_minutes() -> u64 { 15 }

// ── Autostart (Windows registry) ──────────────────────────────────────────────

const AUTOSTART_VALUE: &str = "ruscal";
const AUTOSTART_KEY: &str =
    "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run";

fn get_autostart() -> bool {
    use windows::Win32::System::Registry::{
        RegCloseKey, RegOpenKeyExW, RegQueryValueExW,
        HKEY, HKEY_CURRENT_USER, KEY_READ, REG_VALUE_TYPE,
    };
    use windows::core::PCWSTR;
    let key_path: Vec<u16> = format!("{AUTOSTART_KEY}\0").encode_utf16().collect();
    let val_name: Vec<u16> = format!("{AUTOSTART_VALUE}\0").encode_utf16().collect();
    unsafe {
        let mut key = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, PCWSTR::from_raw(key_path.as_ptr()),
                         0, KEY_READ, &mut key).is_err() {
            return false;
        }
        let mut size = 0u32;
        let mut _kind = REG_VALUE_TYPE::default();
        let found = RegQueryValueExW(key, PCWSTR::from_raw(val_name.as_ptr()),
                                     None, Some(&mut _kind), None, Some(&mut size)).is_ok();
        let _ = RegCloseKey(key);
        found
    }
}

/// Pick the path that should be written to the autostart registry entry.
///
/// The registry value must always point at the installed release binary
/// (under `%LOCALAPPDATA%\ruscal\ruscal.exe`). Writing `current_exe()`
/// verbatim is wrong during development — the debug target is a
/// console-subsystem binary, which Windows launches at logon by popping
/// a terminal window. Closing that terminal then kills ruscal (the
/// console sends `CTRL_CLOSE_EVENT`). Returns `None` when there is no
/// installed binary to register, in which case the caller must skip the
/// write rather than leave a broken entry behind.
fn pick_autostart_path(installed: Option<&std::path::Path>, installed_exists: bool)
    -> Option<&std::path::Path>
{
    match installed {
        Some(p) if installed_exists => Some(p),
        _ => None,
    }
}

fn set_autostart(enable: bool) {
    use windows::Win32::System::Registry::{
        RegCloseKey, RegOpenKeyExW, RegSetValueExW, RegDeleteValueW,
        HKEY, HKEY_CURRENT_USER, KEY_WRITE, REG_SZ,
    };
    use windows::core::PCWSTR;
    let key_path: Vec<u16> = format!("{AUTOSTART_KEY}\0").encode_utf16().collect();
    let val_name: Vec<u16> = format!("{AUTOSTART_VALUE}\0").encode_utf16().collect();
    unsafe {
        let mut key = HKEY::default();
        if RegOpenKeyExW(HKEY_CURRENT_USER, PCWSTR::from_raw(key_path.as_ptr()),
                         0, KEY_WRITE, &mut key).is_err() {
            return;
        }
        if enable {
            let installed = updater::installed_path();
            let exists    = installed.as_deref().map(|p| p.exists()).unwrap_or(false);
            if let Some(exe) = pick_autostart_path(installed.as_deref(), exists) {
                let path_str: Vec<u16> = format!("{}\0", exe.display()).encode_utf16().collect();
                let bytes =
                    std::slice::from_raw_parts(path_str.as_ptr() as *const u8, path_str.len() * 2);
                let _ = RegSetValueExW(key, PCWSTR::from_raw(val_name.as_ptr()),
                                       0, REG_SZ, Some(bytes));
            } else {
                log::warn!(
                    "set_autostart: no installed binary on disk — refusing to register \
                     a startup entry (would otherwise point at a dev target)"
                );
            }
        } else {
            let _ = RegDeleteValueW(key, PCWSTR::from_raw(val_name.as_ptr()));
        }
        let _ = RegCloseKey(key);
    }
}

/// Overwrite the autostart registry value with the installed binary path
/// whenever the entry exists. Self-heals users whose registry still points
/// at `target/debug/ruscal.exe` from an earlier version that used
/// `current_exe()` at first-run. Idempotent: a no-op once the value is
/// already correct (Windows treats identical writes as a single update).
fn repair_autostart() {
    if get_autostart() {
        set_autostart(true);
    }
}

#[derive(serde::Serialize, serde::Deserialize, Default)]
struct AppConfig {
    pairs:          Vec<SavedPair>,
    last_synced_at: Option<chrono::DateTime<chrono::Utc>>,
    #[serde(default = "default_interval_minutes")]
    sync_interval_minutes: u64,
    #[serde(default)]
    browser_path: Option<String>,
    /// Whether ruscal should launch hidden to the tray on startup.
    ///
    /// `None` means "never configured" — we treat that as the default-on state
    /// so fresh installs start minimized without requiring a config write on
    /// first launch.
    #[serde(default)]
    start_minimized: Option<bool>,
}

/// Effective value of the start-minimized setting with its default (`true`) applied.
fn start_minimized_effective(cfg: &AppConfig) -> bool {
    cfg.start_minimized.unwrap_or(true)
}

fn config_path() -> std::path::PathBuf {
    dirs::data_local_dir()
        .unwrap_or_else(|| std::path::PathBuf::from("."))
        .join("ruscal")
        .join("config.json")
}

fn load_config() -> AppConfig {
    let path = config_path();
    let mut cfg: AppConfig = std::fs::read_to_string(path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default();
    if cfg.sync_interval_minutes == 0 {
        cfg.sync_interval_minutes = default_interval_minutes();
    }
    cfg
}

/// Save pairs to disk. Preserves the existing `last_synced_at` and `sync_interval_minutes`
/// unless overridden.
fn save_config(pairs: &VecModel<SyncPair>, new_sync_time: Option<chrono::DateTime<chrono::Utc>>) {
    let existing = load_config();
    let last_synced_at = new_sync_time.or(existing.last_synced_at);

    let saved: Vec<SavedPair> = (0..pairs.row_count())
        .filter_map(|i| pairs.row_data(i))
        .filter(|p| !p.dest_id.is_empty())
        .map(|p| SavedPair {
            source_account: p.source_account.to_string(),
            dest_account:   p.dest_name.to_string(),    // calendar display name
            google_email:   p.dest_account.to_string(), // google account email
            dest_id:        p.dest_id.to_string(),
        })
        .collect();

    write_config(AppConfig {
        pairs: saved,
        last_synced_at,
        sync_interval_minutes: existing.sync_interval_minutes,
        browser_path: existing.browser_path,
        start_minimized: existing.start_minimized,
    });
}

/// Save from a Vec<SavedPair> (used from sync completion on the UI thread).
fn save_config_vec(pairs: &[SavedPair], new_sync_time: Option<chrono::DateTime<chrono::Utc>>) {
    let existing = load_config();
    let last_synced_at = new_sync_time.or(existing.last_synced_at);
    write_config(AppConfig {
        pairs: pairs.to_vec(),
        last_synced_at,
        sync_interval_minutes: existing.sync_interval_minutes,
        browser_path: existing.browser_path,
        start_minimized: existing.start_minimized,
    });
}

/// Save a new sync interval, preserving everything else.
fn save_interval(minutes: u64) {
    let mut config = load_config();
    config.sync_interval_minutes = minutes;
    write_config(config);
}

/// Save the start-minimized toggle, preserving everything else.
fn save_start_minimized(enabled: bool) {
    let mut config = load_config();
    config.start_minimized = Some(enabled);
    write_config(config);
}


fn write_config(config: AppConfig) {
    let path = config_path();
    if let Some(dir) = path.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    if let Ok(json) = serde_json::to_string_pretty(&config) {
        let _ = std::fs::write(path, json);
    }
}

// ── Status text helpers ───────────────────────────────────────────────────────

/// Format "Synced Xm ago" or "Synced at HH:MM" for older syncs.
fn format_last_sync(dt: chrono::DateTime<chrono::Utc>) -> String {
    let secs = chrono::Utc::now()
        .signed_duration_since(dt)
        .num_seconds()
        .max(0);

    if secs < 60 {
        format!("Synced {}s ago", secs)
    } else if secs < 3600 {
        format!("Synced {}m ago", secs / 60)
    } else if secs < 86400 {
        let h = secs / 3600;
        let m = (secs % 3600) / 60;
        if m == 0 { format!("Synced {}h ago", h) } else { format!("Synced {}h {}m ago", h, m) }
    } else {
        let local = dt.with_timezone(&chrono::Local);
        format!("Synced {}", local.format("%-d %b at %-I:%M %p"))
    }
}

/// "next in Xm" countdown based on when the last sync happened + interval.
/// Falls back to `app_start` if never synced.
fn format_next_sync(
    last_synced: Option<chrono::DateTime<chrono::Utc>>,
    app_start: chrono::DateTime<chrono::Utc>,
    interval_secs: u64,
) -> String {
    let base = last_synced.unwrap_or(app_start);
    let next_at = base + chrono::Duration::seconds(interval_secs as i64);
    let secs_until = (next_at - chrono::Utc::now()).num_seconds().max(0);

    if secs_until < 60 {
        "next in <1m".into()
    } else {
        format!("next in {}m", secs_until / 60)
    }
}

/// Full status line for the header.
fn status_text(
    last: Option<chrono::DateTime<chrono::Utc>>,
    app_start: chrono::DateTime<chrono::Utc>,
    has_pairs: bool,
    interval_secs: u64,
) -> String {
    if !has_pairs {
        return "No calendars configured".into();
    }
    let next = format_next_sync(last, app_start, interval_secs);
    match last {
        Some(dt) => format!("{} · {}", format_last_sync(dt), next),
        None     => format!("Never synced · {}", next),
    }
}

/// Does this raw sync error mean the user needs to re-authenticate with Google?
///
/// Covers every shape the sync path can produce for "Google no longer accepts
/// our credentials":
/// * `auth revoked` — [`GoogleError::AuthRevoked`] (refresh returned `invalid_grant`)
/// * `invalid_grant` — raw OAuth error, in case it reaches us without going through the variant
/// * `No stored token` — there is no token file for the configured account yet
/// * `401` / `403` — Google API rejected the access token we sent
///
/// One detector, one UI treatment: the "Re-authenticate" button.
fn needs_reauth(e: &str) -> bool {
    e.contains("auth revoked")
        || e.contains("invalid_grant")
        || e.contains("No stored token")
        || e.contains("401")
        || e.contains("403")
}

/// Translate raw sync errors into user-friendly messages.
fn friendly_error(e: &str) -> String {
    if e.contains("MAPI_E_LOGON_FAILED") || e.contains("0x80040111") {
        "Outlook not running — open Outlook and try again".into()
    } else if e.contains("MAPI_E_NOT_INITIALIZED") || e.contains("0x8004011F") {
        "Outlook MAPI not initialized — restart Outlook".into()
    } else if e.contains("MAPI") {
        format!("Outlook error: {e}")
    } else if needs_reauth(e) {
        "Google access needs re-authentication — click Re-authenticate below".into()
    } else {
        format!("Sync error: {e}")
    }
}

#[cfg(test)]
mod tests {
    use super::{friendly_error, needs_reauth};

    /// Every shape of "Google no longer accepts our creds" must be detected.
    /// One detector, so any auth failure shows the same UI treatment.
    #[test]
    fn needs_reauth_covers_every_auth_failure_shape() {
        let reauth_cases = [
            "auth revoked: refresh token rejected",
            r#"auth: refresh failed: {"error":"invalid_grant"}"#,
            "auth: No stored token for foo@bar.com",
            "API: HTTP 401 Unauthorized",
            "API: HTTP 403 Forbidden",
        ];
        for raw in reauth_cases {
            assert!(needs_reauth(raw), "needs_reauth missed: {raw}");
        }

        let non_reauth_cases = [
            "Outlook error: MAPI_E_LOGON_FAILED",
            "some random network blip",
            "HTTP 500 Internal Server Error",
        ];
        for raw in non_reauth_cases {
            assert!(!needs_reauth(raw), "needs_reauth false-positive: {raw}");
        }
    }

    /// All reauth cases must map to a single friendly message that mentions
    /// the Re-authenticate button, so there is exactly one hardcoded string.
    #[test]
    fn friendly_error_gives_one_message_for_all_reauth_cases() {
        let expected = "Google access needs re-authentication — click Re-authenticate below";
        let cases = [
            "auth revoked: refresh token rejected by Google",
            r#"auth: refresh failed: {"error":"invalid_grant"}"#,
            "auth: No stored token for foo@bar.com — please re-authenticate",
            "API: HTTP 401 Unauthorized",
            "API: HTTP 403 Forbidden",
        ];
        for raw in cases {
            assert_eq!(friendly_error(raw), expected, "mismatch for: {raw}");
        }
    }

    #[test]
    fn friendly_error_passes_through_generic() {
        let raw = "some unknown failure";
        let out = friendly_error(raw);
        assert!(out.starts_with("Sync error"), "got: {out}");
    }

    // ── Autostart path picker ────────────────────────────────────────────────
    //
    // Regression guard: before this test, `set_autostart(true)` wrote
    // `current_exe()` into `HKCU\...\Run`. On a dev machine that resolved to
    // `target\debug\ruscal.exe`, a console-subsystem binary that pops a
    // terminal window at logon and closes the app when that terminal is
    // closed. The pure helper must *never* hand out that path — the registry
    // entry has to point at the installed release binary or nothing at all.

    use super::pick_autostart_path;
    use std::path::PathBuf;

    #[test]
    fn autostart_picks_installed_path_when_it_exists() {
        let installed = PathBuf::from(r"C:\Users\x\AppData\Local\ruscal\ruscal.exe");
        let picked = pick_autostart_path(Some(installed.as_path()), true);
        assert_eq!(picked, Some(installed.as_path()));
    }

    #[test]
    fn autostart_refuses_when_installed_binary_missing() {
        // A dev machine that never ran a release has no installed exe.
        // Writing anything would either register a debug build (terminal at
        // logon) or a bogus non-existent path — both worse than no entry.
        let installed = PathBuf::from(r"C:\Users\x\AppData\Local\ruscal\ruscal.exe");
        assert_eq!(pick_autostart_path(Some(installed.as_path()), false), None);
    }

    #[test]
    fn autostart_refuses_when_install_path_undetermined() {
        assert_eq!(pick_autostart_path(None, false), None);
        assert_eq!(pick_autostart_path(None, true), None);
    }

    // ── Settings cog toggle ────────────────────────────────────────────────
    //
    // The cog in the header is a single button and the user expects it to
    // both open *and* close the settings page. Clicking it twice used to
    // leave you stranded on the settings screen with no obvious way back.
    use super::page;

    #[test]
    fn settings_toggle_from_settings_returns_to_pairs() {
        assert_eq!(page::toggle_settings(page::SETTINGS), page::PAIRS);
    }

    #[test]
    fn settings_toggle_from_any_other_page_opens_settings() {
        for from in [page::PAIRS, page::SOURCE_PICK, page::DEST_PICK, page::GOOGLE_PICK] {
            assert_eq!(page::toggle_settings(from), page::SETTINGS, "from page {from}");
        }
    }
}

// ── Entry point ───────────────────────────────────────────────────────────────

fn main() {
    env_logger::init();

    let launch_args = updater::parse_args();
    updater::cleanup_stale_update_exe();

    // Self-install runs release-only: a `cargo run` dev loop must not stomp
    // the installed production binary at %LOCALAPPDATA%\ruscal.
    #[cfg(not(debug_assertions))]
    if !updater::is_installed_path() {
        let flag = launch_args.just_updated.as_deref()
            .map(|v| format!("--just-updated={v}"))
            .unwrap_or_else(|| "--just-installed".to_owned());
        updater::self_install(Some(&flag)); // diverges — exits this process
    }

    // Single-instance: if ruscal is already running, bring the existing
    // window to the front and quit this process. Standard Windows app
    // behavior — a second click on the shortcut must not respawn the app
    // (that would close the running window, drop the tray icon, and show
    // a fresh startup flash). Applies to dev + release: `cargo run` also
    // defers to an autostarted release instead of starting side-by-side.
    let _instance_guard = match updater::acquire_single_instance() {
        Some(g) => g,
        None => {
            // Don't touch the other process's window directly — we go behind
            // Slint's back if we call ShowWindow/SetForegroundWindow on a
            // foreign HWND, and the primary's `p.hide()` would then silently
            // no-op because its visibility state stayed stale. Instead, nudge
            // the primary via a named event and let it surface *itself*
            // through Slint's own API.
            updater::signal_focus_request();
            return;
        }
    };

    #[cfg(not(debug_assertions))]
    {
        // Record the version of the installed (production) binary so a later
        // `cargo run` from a dev build can still tell the user whether the
        // cached copy in %LOCALAPPDATA% is out of date relative to GitHub.
        updater::record_installed_version(APP_VERSION);

        // Self-heal a stale autostart entry from earlier versions that wrote
        // `target/debug/ruscal.exe` on first-run — that's a console-subsystem
        // binary and pops a terminal at logon.
        repair_autostart();

        // Existing installs from before the Start-menu integration shipped
        // don't have a `.lnk` in Programs; create one so ruscal appears in
        // Windows Search without waiting for another self-install cycle.
        updater::ensure_start_menu_shortcut();
    }

    unsafe {
        let _ = SetProcessDpiAwarenessContext(
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        );
    }

    let app_start = chrono::Utc::now();

    // ── Tray icon ─────────────────────────────────────────────────────────────

    let tray_menu = Menu::new();
    let sync_item = MenuItem::new("Sync Now",    true, None);
    let sep       = PredefinedMenuItem::separator();
    let quit_item = MenuItem::new("Quit ruscal", true, None);
    tray_menu.append_items(&[&sync_item, &sep, &quit_item]).unwrap();

    let _tray = TrayIconBuilder::new()
        .with_icon(build_icon())
        .with_tooltip("ruscal")
        .with_menu(Box::new(tray_menu))
        .with_menu_on_left_click(false)
        .build()
        .expect("failed to create tray icon");

    // ── Slint panel ───────────────────────────────────────────────────────────

    let panel = TrayPanel::new().expect("failed to create panel");

    let outlook_color = slint::Color::from_rgb_u8(0, 114, 198);
    let gcal_color    = slint::Color::from_rgb_u8(66, 133, 244);

    // Load persisted config.
    let first_run = !config_path().exists();
    let config = load_config();
    let last_synced: Arc<Mutex<Option<chrono::DateTime<chrono::Utc>>>> =
        Arc::new(Mutex::new(config.last_synced_at));
    let browser_path: Arc<Mutex<Option<String>>> =
        Arc::new(Mutex::new(config.browser_path.clone()));

    // Enable "Start with Windows" by default on first run.
    if first_run { set_autostart(true); }
    let interval_secs: Rc<std::cell::Cell<u64>> =
        Rc::new(std::cell::Cell::new(config.sync_interval_minutes.max(1) * 60));

    let make_placeholder = move || SyncPair {
        source_name:    "".into(),
        source_account: "".into(),
        source_color:   outlook_color,
        dest_name:      "".into(),
        dest_account:   "".into(),
        dest_color:     gcal_color,
        dest_id:        "".into(),
    };

    let mut initial_pairs: Vec<SyncPair> = config.pairs.iter().map(|p| SyncPair {
        source_name:    "Outlook Calendar".into(),
        source_account: p.source_account.as_str().into(),
        source_color:   outlook_color,
        dest_name:      p.dest_account.as_str().into(),  // calendar name
        dest_account:   p.google_email.as_str().into(),  // google account
        dest_color:     gcal_color,
        dest_id:        p.dest_id.as_str().into(),
    }).collect();
    initial_pairs.push(make_placeholder());

    let pairs_model = Rc::new(VecModel::from(initial_pairs));
    panel.set_pairs(pairs_model.clone().into());
    panel.set_sync_status(SyncStatus::Idle);
    panel.set_sync_interval_minutes(config.sync_interval_minutes.max(1) as i32);
    panel.set_start_with_windows(get_autostart());
    panel.set_start_minimized(start_minimized_effective(&config));
    panel.set_browser_path(config.browser_path.clone().unwrap_or_default().into());

    let stored_accounts = google::list_stored_accounts();
    panel.set_google_accounts(Rc::new(VecModel::from(
        stored_accounts.into_iter().map(slint::SharedString::from).collect::<Vec<_>>()
    )).into());

    let has_pairs = !config.pairs.is_empty();
    panel.set_last_sync_text(
        status_text(*last_synced.lock().unwrap(), app_start, has_pairs, interval_secs.get()).into()
    );

    // Override status line for first-run / post-update messages.
    if launch_args.just_installed {
        panel.set_last_sync_text(
            "Installed · starts with Windows".into()
        );
    } else if launch_args.just_updated.is_some() {
        panel.set_last_sync_text(
            format!("Updated to {}", APP_VERSION).into()
        );
    }

    // ── Sync callback ─────────────────────────────────────────────────────────

    panel.on_sync_now({
        let weak          = panel.as_weak();
        let pairs         = pairs_model.clone();
        let last_synced   = last_synced.clone();
        let interval_secs = interval_secs.clone();
        move || {
            let Some(p) = weak.upgrade() else { return };
            if p.get_sync_status() == SyncStatus::Syncing { return; }

            // Collect (dest_url, google_email) for each configured pair.
            // dest_account in Slint holds the google email.
            let sync_targets: Vec<(String, String, String)> = (0..pairs.row_count())
                .filter_map(|i| pairs.row_data(i))
                .filter(|pair| !pair.source_name.is_empty() && !pair.dest_id.is_empty())
                .map(|pair| (
                    state::pair_id(&pair.source_account, &pair.dest_id),
                    pair.dest_id.to_string(),
                    pair.dest_account.to_string(),
                ))
                .collect();

            if sync_targets.is_empty() {
                p.set_sync_status(SyncStatus::Idle);
                p.set_last_sync_text("No calendars configured".into());
                return;
            }

            // Snapshot pairs for saving in the completion callback.
            let saved_pairs: Vec<SavedPair> = (0..pairs.row_count())
                .filter_map(|i| pairs.row_data(i))
                .filter(|pair| !pair.dest_id.is_empty())
                .map(|pair| SavedPair {
                    source_account: pair.source_account.to_string(),
                    dest_account:   pair.dest_name.to_string(),
                    google_email:   pair.dest_account.to_string(),
                    dest_id:        pair.dest_id.to_string(),
                })
                .collect();

            p.set_sync_status(SyncStatus::Syncing);
            p.set_last_sync_text("Syncing…".into());

            let secs  = interval_secs.get();
            let weak2 = weak.clone();
            let last_synced = last_synced.clone();
            std::thread::spawn(move || {
                let result: Result<(usize, Vec<String>), String> = (|| {
                    let mut synced  = 0usize;
                    let mut skipped_titles: Vec<String> = Vec::new();
                    for (pair_id, dest_url, google_email) in &sync_targets {
                        let token = google::get_access_token_for(google_email)
                            .map_err(|e| e.to_string())?;
                        let report = sync::run_sync(pair_id, dest_url, &token)?;
                        synced  += report.synced;
                        skipped_titles.extend(report.skipped_titles);
                    }
                    Ok((synced, skipped_titles))
                })();

                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    match result {
                        Ok((_synced, skipped_titles)) => {
                            let now = chrono::Utc::now();
                            *last_synced.lock().unwrap() = Some(now);
                            save_config_vec(&saved_pairs, Some(now));
                            p.set_sync_status(SyncStatus::Success);
                            p.set_skipped_count(skipped_titles.len() as i32);
                            p.set_skipped_detail(skipped_titles.join("\n").into());
                            p.set_last_sync_text(
                                format!("{} · {}",
                                    format_last_sync(now),
                                    format_next_sync(Some(now), app_start, secs),
                                ).into()
                            );
                        }
                        Err(ref e) => {
                            log::error!("Sync failed: {e}");
                            let raw = e.as_str();
                            p.set_sync_status(SyncStatus::Error);
                            p.set_sync_error_detail(friendly_error(raw).into());
                            p.set_auth_needed(needs_reauth(raw));
                            p.set_last_sync_text(
                                status_text(*last_synced.lock().unwrap(), app_start, true, secs).into()
                            );
                        }
                    }
                }).ok();
            });
        }
    });

    // ── Delete pair ───────────────────────────────────────────────────────────

    panel.on_delete_pair({
        let weak          = panel.as_weak();
        let pairs         = pairs_model.clone();
        let last_synced   = last_synced.clone();
        let interval_secs = interval_secs.clone();
        move |i| {
            let i = i as usize;
            // Don't delete the placeholder row.
            if i >= pairs.row_count() { return; }
            if pairs.row_data(i).map(|p| p.dest_id.is_empty()).unwrap_or(true) { return; }

            pairs.remove(i);

            // Ensure there's always a placeholder at the end.
            let last_is_placeholder = pairs.row_count() > 0
                && pairs.row_data(pairs.row_count() - 1)
                    .map(|p| p.dest_id.is_empty())
                    .unwrap_or(false);
            if !last_is_placeholder {
                pairs.push(make_placeholder());
            }

            save_config(&pairs, None);

            // Update status if no pairs remain.
            let has_pairs = (0..pairs.row_count())
                .filter_map(|j| pairs.row_data(j))
                .any(|p| !p.dest_id.is_empty());
            if let Some(p) = weak.upgrade() {
                p.set_last_sync_text(
                    status_text(*last_synced.lock().unwrap(), app_start, has_pairs, interval_secs.get()).into()
                );
            }
        }
    });

    // ── Clear source / dest (single-side clear) ───────────────────────────────

    panel.on_clear_source({
        let pairs         = pairs_model.clone();
        let last_synced   = last_synced.clone();
        let interval_secs = interval_secs.clone();
        let weak          = panel.as_weak();
        move |i| {
            let i = i as usize;
            let Some(mut pair) = pairs.row_data(i) else { return };
            pair.source_name    = "".into();
            pair.source_account = "".into();
            pairs.set_row_data(i, pair);
            save_config(&pairs, None);
            if let Some(p) = weak.upgrade() {
                let has_pairs = (0..pairs.row_count())
                    .filter_map(|j| pairs.row_data(j))
                    .any(|p| !p.dest_id.is_empty());
                p.set_last_sync_text(
                    status_text(*last_synced.lock().unwrap(), app_start, has_pairs, interval_secs.get()).into()
                );
            }
        }
    });

    panel.on_clear_dest({
        let pairs         = pairs_model.clone();
        let last_synced   = last_synced.clone();
        let interval_secs = interval_secs.clone();
        let weak          = panel.as_weak();
        move |i| {
            let i = i as usize;
            let Some(mut pair) = pairs.row_data(i) else { return };
            pair.dest_name    = "".into();
            pair.dest_account = "".into();
            pair.dest_id      = "".into();
            pairs.set_row_data(i, pair);
            // Clean up: if both sides are now empty and it's not the placeholder, remove it.
            let both_empty = pairs.row_data(i)
                .map(|p| p.source_name.is_empty() && p.dest_id.is_empty())
                .unwrap_or(false);
            let is_placeholder_slot = i + 1 == pairs.row_count();
            if both_empty && !is_placeholder_slot {
                pairs.remove(i);
            }
            save_config(&pairs, None);
            if let Some(p) = weak.upgrade() {
                let has_pairs = (0..pairs.row_count())
                    .filter_map(|j| pairs.row_data(j))
                    .any(|p| !p.dest_id.is_empty());
                p.set_last_sync_text(
                    status_text(*last_synced.lock().unwrap(), app_start, has_pairs, interval_secs.get()).into()
                );
            }
        }
    });

    // ── Other callbacks ───────────────────────────────────────────────────────

    panel.on_minimize({
        let weak = panel.as_weak();
        move || { if let Some(p) = weak.upgrade() { p.hide().ok(); } }
    });

    panel.on_quit(|| { slint::quit_event_loop().ok(); });

    panel.on_open_settings({
        let weak = panel.as_weak();
        move || {
            if let Some(p) = weak.upgrade() {
                // Toggle: if already on the Settings page, return to the pairs
                // list so the cog button acts like a show/hide switch.
                p.set_current_page(page::toggle_settings(p.get_current_page()));
            }
        }
    });

    panel.on_configure_source({
        let weak = panel.as_weak();
        move |i| {
            let Some(p) = weak.upgrade() else { return };
            p.set_outlook_calendars(Rc::new(VecModel::from(vec![])).into());
            p.set_config_pair_index(i);
            p.set_current_page(page::SOURCE_PICK);

            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = outlook::list_calendar_sources();
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    if p.get_current_page() != page::SOURCE_PICK { return; }
                    match result {
                        Ok(calendars) => {
                            let names: Rc<VecModel<slint::SharedString>> = Rc::new(
                                VecModel::from(
                                    calendars.into_iter()
                                        .map(|c| slint::SharedString::from(c.display_name))
                                        .collect::<Vec<_>>()
                                )
                            );
                            p.set_outlook_calendars(names.into());
                        }
                        Err(e) => {
                            log::error!("Failed to list Outlook calendars: {e:?}");
                            p.set_current_page(page::PAIRS);
                            p.set_last_sync_text(
                                friendly_error(&format!("Outlook: {e}")).into()
                            );
                        }
                    }
                }).ok();
            });
        }
    });

    panel.on_source_selected({
        let weak  = panel.as_weak();
        let pairs = pairs_model.clone();
        move |cal_index| {
            let Some(p) = weak.upgrade() else { return };
            let pair_index = p.get_config_pair_index() as usize;
            let calendars  = p.get_outlook_calendars();
            if let Some(display_name) = calendars.row_data(cal_index as usize) {
                let mut pair = pairs.row_data(pair_index).unwrap_or_default();
                pair.source_name    = "Outlook Calendar".into();
                pair.source_account = display_name;
                pair.source_color   = outlook_color;
                pairs.set_row_data(pair_index, pair);
                save_config(&pairs, None);
            }
            p.set_current_page(page::PAIRS);
        }
    });

    panel.on_configure_dest({
        let weak = panel.as_weak();
        move |i| {
            let Some(p) = weak.upgrade() else { return };
            p.set_config_pair_index(i);
            p.set_current_page(page::DEST_PICK);
        }
    });

    // Helper: navigate to page 3 and kick off a background calendar load.
    // `email_hint` = Some(email) to reuse existing tokens, None to force new OAuth.
    let start_google_picker = {
        let weak = panel.as_weak();
        let browser_path = browser_path.clone();
        move |email_hint: Option<String>| {
            let browser = browser_path.lock().unwrap().clone();
            let Some(p) = weak.upgrade() else { return };
            p.set_google_calendars(Rc::new(VecModel::from(vec![])).into());
            p.set_google_calendar_ids(Rc::new(VecModel::from(vec![])).into());
            p.set_google_account_email("".into());
            let status = if email_hint.is_none() {
                "Waiting for Google authorization in your browser…"
            } else {
                "Loading calendars…"
            };
            p.set_google_status(status.into());
            p.set_current_page(page::GOOGLE_PICK);

            // Hide the window while the browser OAuth flow is open.
            if email_hint.is_none() {
                p.hide().ok();
            }

            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = google::list_google_calendars(email_hint.as_deref(), browser.as_deref());
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    // Bring the window back regardless of outcome.
                    snap_to_tray(&p);
                    p.show().ok();
                    if p.get_current_page() != page::GOOGLE_PICK { return; }
                    match result {
                        Ok((calendars, email)) => {
                            let names: Rc<VecModel<slint::SharedString>> = Rc::new(VecModel::from(
                                calendars.iter().map(|c| slint::SharedString::from(&c.summary)).collect::<Vec<_>>()
                            ));
                            let ids: Rc<VecModel<slint::SharedString>> = Rc::new(VecModel::from(
                                calendars.iter().map(|c| slint::SharedString::from(&c.id)).collect::<Vec<_>>()
                            ));
                            p.set_google_calendars(names.into());
                            p.set_google_calendar_ids(ids.into());
                            p.set_google_account_email(email.into());
                            p.set_google_status("".into());
                            // Refresh accounts list in case a new account was just added.
                            let accounts = google::list_stored_accounts();
                            p.set_google_accounts(Rc::new(VecModel::from(
                                accounts.into_iter().map(slint::SharedString::from).collect::<Vec<_>>()
                            )).into());
                        }
                        Err(e) => {
                            log::error!("Failed to list Google Calendars: {e}");
                            p.set_google_status(format!("Error: {e}").into());
                        }
                    }
                }).ok();
            });
        }
    };

    panel.on_dest_provider_selected({
        let weak  = panel.as_weak();
        let pairs = pairs_model.clone();
        let start = start_google_picker.clone();
        move |_provider_index| {
            let Some(p) = weak.upgrade() else { return };
            if p.get_current_page() == page::GOOGLE_PICK { return; }
            // Use the existing account for this pair if it has one.
            let pair_index = p.get_config_pair_index() as usize;
            let email_hint = pairs.row_data(pair_index)
                .filter(|pair| !pair.dest_account.is_empty())
                .map(|pair| pair.dest_account.to_string());
            start(email_hint);
        }
    });

    // "+ Add account" chip on page 3 footer — forces new OAuth, stays on page 3.
    panel.on_settings_google_sign_out({
        let start = start_google_picker.clone();
        move || { start(None); }
    });

    // Account chip clicked on page 3 footer — switch to an existing account.
    panel.on_switch_google_account({
        let start = start_google_picker.clone();
        move |email| { start(Some(email.to_string())); }
    });

    // "Add Google account" from Settings — OAuth without navigating away from settings.
    panel.on_add_google_account({
        let weak = panel.as_weak();
        let browser_path = browser_path.clone();
        move || {
            let Some(p) = weak.upgrade() else { return };
            p.hide().ok();
            let browser = browser_path.lock().unwrap().clone();
            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = google::authorize_new_account(browser.as_deref());
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    snap_to_tray(&p);
                    p.show().ok();
                    if let Ok((_token, _email)) = result {
                        let accounts = google::list_stored_accounts();
                        p.set_google_accounts(Rc::new(VecModel::from(
                            accounts.into_iter().map(slint::SharedString::from).collect::<Vec<_>>()
                        )).into());
                    }
                }).ok();
            });
        }
    });

    // "Re-authenticate" button inside the sync-error popup — only reached when
    // the sync failed with a revoked/missing Google refresh token. Runs the
    // same OAuth flow as "Add Google account" and clears the auth-needed flag
    // on success so the banner doesn't linger.
    panel.on_re_authenticate({
        let weak = panel.as_weak();
        let browser_path = browser_path.clone();
        move || {
            let Some(p) = weak.upgrade() else { return };
            p.hide().ok();
            let browser = browser_path.lock().unwrap().clone();
            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = google::authorize_new_account(browser.as_deref());
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    snap_to_tray(&p);
                    p.show().ok();
                    if let Ok((_token, _email)) = result {
                        p.set_auth_needed(false);
                        let accounts = google::list_stored_accounts();
                        p.set_google_accounts(Rc::new(VecModel::from(
                            accounts.into_iter().map(slint::SharedString::from).collect::<Vec<_>>()
                        )).into());
                    }
                }).ok();
            });
        }
    });

    // "Remove" button next to an account in Settings.
    panel.on_remove_google_account({
        let weak = panel.as_weak();
        move |email| {
            google::sign_out_account(&email);
            let accounts = google::list_stored_accounts();
            if let Some(p) = weak.upgrade() {
                p.set_google_accounts(Rc::new(VecModel::from(
                    accounts.into_iter().map(slint::SharedString::from).collect::<Vec<_>>()
                )).into());
            }
        }
    });

    panel.on_dest_selected({
        let weak  = panel.as_weak();
        let pairs = pairs_model.clone();
        move |cal_index| {
            let Some(p) = weak.upgrade() else { return };
            let pair_index   = p.get_config_pair_index() as usize;
            let google_email = p.get_google_account_email().to_string();

            let cal_ids   = p.get_google_calendar_ids();
            let calendars = p.get_google_calendars();
            let Some(summary) = calendars.row_data(cal_index as usize) else { return };
            let dest_id = cal_ids.row_data(cal_index as usize).unwrap_or_default();

            let mut pair = pairs.row_data(pair_index).unwrap_or_default();
            pair.dest_name    = summary;           // calendar display name
            pair.dest_account = google_email.into(); // google account email
            pair.dest_color   = gcal_color;
            pair.dest_id      = dest_id;
            pairs.set_row_data(pair_index, pair.clone());

            if pair_index + 1 >= pairs.row_count()
                && !pair.source_name.is_empty()
                && !pair.dest_id.is_empty()
            {
                pairs.push(make_placeholder());
            }

            save_config(&pairs, None);
            p.set_current_page(page::PAIRS);
            p.invoke_sync_now();
        }
    });

    // ── Update callbacks ──────────────────────────────────────────────────────

    panel.on_update_now({
        let weak = panel.as_weak();
        move || {
            let Some(p) = weak.upgrade() else { return };
            let version = p.get_update_version().to_string();
            if version.is_empty() { return; }

            p.set_update_version("".into());
            p.set_last_sync_text(format!("Downloading v{version}…").into());

            let weak2 = weak.clone();
            std::thread::spawn(move || {
                match updater::download_update(&version) {
                    Ok(temp_path) => {
                        // Spawn the downloaded exe — it self-installs (terminates us, copies itself).
                        let flag = format!("--just-updated={}", APP_VERSION);
                        if std::process::Command::new(&temp_path).arg(&flag).spawn().is_ok() {
                            slint::invoke_from_event_loop(|| {
                                slint::quit_event_loop().ok();
                            }).ok();
                        } else {
                            slint::invoke_from_event_loop(move || {
                                if let Some(p) = weak2.upgrade() {
                                    p.set_last_sync_text("Update failed: could not launch installer".into());
                                    p.set_update_version(version.into());
                                }
                            }).ok();
                        }
                    }
                    Err(e) => {
                        log::error!("Update download failed: {e}");
                        slint::invoke_from_event_loop(move || {
                            if let Some(p) = weak2.upgrade() {
                                p.set_last_sync_text(format!("Update failed: {e}").into());
                                p.set_update_version(version.into());
                            }
                        }).ok();
                    }
                }
            });
        }
    });

    panel.on_browse_for_browser({
        let weak = panel.as_weak();
        let browser_path = browser_path.clone();
        move || {
            // Open a file picker via PowerShell to select an exe.
            let mut cmd = std::process::Command::new("powershell");
            cmd.args(["-NoProfile", "-NonInteractive", "-Command",
                "[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; \
                 $d = New-Object System.Windows.Forms.OpenFileDialog; \
                 $d.Filter = 'Executables (*.exe)|*.exe'; \
                 $d.Title = 'Select browser'; \
                 if ($d.ShowDialog() -eq 'OK') { $d.FileName }"]);
            updater::hide_console_window(&mut cmd);
            let result = cmd.output();
            if let Ok(out) = result {
                let path = String::from_utf8_lossy(&out.stdout).trim().to_string();
                if !path.is_empty() {
                    *browser_path.lock().unwrap() = Some(path.clone());
                    let mut cfg = load_config();
                    cfg.browser_path = Some(path.clone());
                    let _ = std::fs::write(config_path(), serde_json::to_string_pretty(&cfg).unwrap_or_default());
                    if let Some(p) = weak.upgrade() {
                        p.set_browser_path(path.into());
                    }
                }
            }
        }
    });

    panel.on_open_config_dir(|| {
        if let Some(dir) = dirs::data_local_dir() {
            let _ = std::process::Command::new("explorer")
                .arg(dir.join("ruscal"))
                .spawn();
        }
    });

    panel.on_dismiss_notification({
        let weak = panel.as_weak();
        move || {
            if let Some(p) = weak.upgrade() {
                p.set_update_version("".into());
            }
        }
    });

    panel.on_check_for_update({
        let weak = panel.as_weak();
        move || {
            let Some(p) = weak.upgrade() else { return };
            p.set_checking_for_update(true);
            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = updater::check_for_update(APP_VERSION);
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    p.set_checking_for_update(false);
                    if let Some(version) = result {
                        p.set_update_version(version.into());
                    }
                }).ok();
            });
        }
    });

    panel.on_copy_error({
        let weak = panel.as_weak();
        move || {
            let Some(p) = weak.upgrade() else { return };
            let text = p.get_sync_error_detail().to_string();
            if text.is_empty() { return; }
            let escaped = text.replace('\'', "''");
            let mut cmd = std::process::Command::new("powershell");
            cmd.args(["-NoProfile", "-NonInteractive", "-Command",
                      &format!("Set-Clipboard -Value '{escaped}'")]);
            updater::hide_console_window(&mut cmd);
            let _ = cmd.output();
            p.set_error_copy_done(true);
            let weak2 = weak.clone();
            slint::Timer::single_shot(std::time::Duration::from_millis(1500), move || {
                if let Some(p) = weak2.upgrade() { p.set_error_copy_done(false); }
            });
        }
    });

    panel.set_app_version(APP_VERSION.into());

    // The window must be shown at least once — Slint's event loop exits if no
    // window is ever realised. When start-minimized is on, we defer a
    // `hide()` via a single-shot timer so the hide runs *inside* the event
    // loop; calling it before the loop starts would race with the window's
    // realisation and cause the process to exit immediately.
    //
    // The window is always surfaced on notable first-impression events —
    // first run (setup wizard), fresh install, or just-updated — so the user
    // gets visible confirmation regardless of the setting.
    panel.show().ok();

    // Listen for focus-requests from subsequent `ruscal.exe` launches. The
    // listener keeps the HANDLE alive for the process lifetime; dropping
    // `_focus_listener` would close the event and stop servicing wake-ups.
    let _focus_listener = {
        let weak = panel.as_weak();
        updater::listen_for_focus_requests(move || {
            let weak = weak.clone();
            slint::invoke_from_event_loop(move || {
                if let Some(p) = weak.upgrade() {
                    p.show().ok();
                    updater::bring_own_window_forward();
                }
            }).ok();
        })
    };

    let force_show =
        first_run
        || launch_args.just_installed
        || launch_args.just_updated.is_some();

    if !force_show && start_minimized_effective(&config) {
        let weak = panel.as_weak();
        slint::Timer::single_shot(std::time::Duration::from_millis(0), move || {
            if let Some(p) = weak.upgrade() { p.hide().ok(); }
        });
    }


    // ── Background update check ────────────────────────────────────────────────

    {
        let weak = panel.as_weak();
        std::thread::spawn(move || {
            if let Some(version) = updater::check_for_update(APP_VERSION) {
                slint::invoke_from_event_loop(move || {
                    if let Some(p) = weak.upgrade() {
                        p.set_update_version(version.into());
                    }
                }).ok();
            }
        });
    }

    // ── Sync on launch ────────────────────────────────────────────────────────

    if !config.pairs.is_empty() {
        let weak_launch = panel.as_weak();
        slint::Timer::single_shot(std::time::Duration::from_millis(500), move || {
            if let Some(p) = weak_launch.upgrade() { p.invoke_sync_now(); }
        });
    }

    // ── Snap panel to bottom-right of work area ───────────────────────────────

    let weak_for_pos = panel.as_weak();
    let _pos_timer = slint::Timer::default();
    _pos_timer.start(
        slint::TimerMode::SingleShot,
        std::time::Duration::from_millis(16),
        move || {
            if let Some(p) = weak_for_pos.upgrade() { snap_to_tray(&p); }
        },
    );

    // ── Status text update every 30 seconds ───────────────────────────────────

    let status_weak    = panel.as_weak();
    let status_synced  = last_synced.clone();
    let status_isecs   = interval_secs.clone();
    let _status_timer  = slint::Timer::default();
    _status_timer.start(
        slint::TimerMode::Repeated,
        std::time::Duration::from_secs(30),
        move || {
            let Some(p) = status_weak.upgrade() else { return };
            if p.get_sync_status() == SyncStatus::Syncing { return; }
            let has_pairs = (0..p.get_pairs().row_count())
                .filter_map(|i| p.get_pairs().row_data(i))
                .any(|pair: SyncPair| !pair.dest_id.is_empty());
            p.set_last_sync_text(
                status_text(*status_synced.lock().unwrap(), app_start, has_pairs, status_isecs.get()).into()
            );
        },
    );

    // ── Auto-sync timer (restartable when interval changes) ───────────────────

    let auto_sync_timer = Rc::new(slint::Timer::default());
    {
        let auto_sync_weak = panel.as_weak();
        let t = auto_sync_timer.clone();
        t.start(
            slint::TimerMode::Repeated,
            std::time::Duration::from_secs(interval_secs.get()),
            move || {
                if let Some(p) = auto_sync_weak.upgrade() {
                    p.invoke_sync_now();
                }
            },
        );
    }

    // ── Settings: interval changed ────────────────────────────────────────────

    panel.on_settings_interval_changed({
        let weak          = panel.as_weak();
        let interval_secs = interval_secs.clone();
        let timer         = auto_sync_timer.clone();
        move |minutes| {
            let minutes = (minutes as u64).max(1);
            let secs    = minutes * 60;
            interval_secs.set(secs);
            save_interval(minutes);

            // Restart auto-sync timer with new interval.
            let weak2 = weak.clone();
            timer.start(
                slint::TimerMode::Repeated,
                std::time::Duration::from_secs(secs),
                move || {
                    if let Some(p) = weak2.upgrade() { p.invoke_sync_now(); }
                },
            );
        }
    });

    panel.on_settings_startup_changed(|enable| {
        set_autostart(enable);
    });

    panel.on_settings_start_minimized_changed(|enable| {
        save_start_minimized(enable);
    });


    // ── Tray + menu event polling ─────────────────────────────────────────────

    let sync_id    = sync_item.id().clone();
    let quit_id    = quit_item.id().clone();
    let panel_weak = panel.as_weak();

    let _tray_timer = slint::Timer::default();
    _tray_timer.start(
        slint::TimerMode::Repeated,
        std::time::Duration::from_millis(50),
        move || {
            while let Ok(ev) = TrayIconEvent::receiver().try_recv() {
                if let TrayIconEvent::Click { button: MouseButton::Left, .. } = ev {
                    if let Some(p) = panel_weak.upgrade() {
                        snap_to_tray(&p);
                        p.show().ok();
                    }
                }
            }

            while let Ok(ev) = MenuEvent::receiver().try_recv() {
                if ev.id() == &sync_id {
                    if let Some(p) = panel_weak.upgrade() {
                        p.invoke_sync_now();
                    }
                } else if ev.id() == &quit_id {
                    slint::quit_event_loop().ok();
                }
            }
        },
    );

    slint::run_event_loop_until_quit().expect("event loop failed");
}

fn snap_to_tray(panel: &TrayPanel) {
    let work   = work_area_phys();
    let size   = panel.window().size();
    let scale  = panel.window().scale_factor();
    let margin = (6.0 * scale).round() as i32;

    let x = work.right  - size.width  as i32 - margin;
    let y = work.bottom - size.height as i32 - margin;

    panel.window().set_position(slint::PhysicalPosition::new(x, y));
}

fn work_area_phys() -> RECT {
    let mut rect = RECT::default();
    unsafe {
        let _ = SystemParametersInfoW(
            SPI_GETWORKAREA, 0,
            Some(&mut rect as *mut RECT as *mut core::ffi::c_void),
            SYSTEM_PARAMETERS_INFO_UPDATE_FLAGS(0),
        );
    }
    rect
}

fn build_icon() -> Icon {
    const SIZE: u32 = 32;
    let svg_data = include_bytes!("../assets/icon.svg");
    let options  = resvg::usvg::Options::default();
    let tree     = resvg::usvg::Tree::from_data(svg_data, &options).expect("valid SVG");
    let mut pixmap = resvg::tiny_skia::Pixmap::new(SIZE, SIZE).expect("pixmap");
    let scale = SIZE as f32 / tree.size().width().max(tree.size().height());
    resvg::render(&tree, resvg::tiny_skia::Transform::from_scale(scale, scale), &mut pixmap.as_mut());
    let mut rgba = pixmap.take();
    for px in rgba.chunks_exact_mut(4) {
        let a = px[3];
        if a > 0 && a < 255 {
            let inv = 255.0 / a as f32;
            px[0] = (px[0] as f32 * inv).min(255.0) as u8;
            px[1] = (px[1] as f32 * inv).min(255.0) as u8;
            px[2] = (px[2] as f32 * inv).min(255.0) as u8;
        }
    }
    Icon::from_rgba(rgba, SIZE, SIZE).expect("valid icon")
}
