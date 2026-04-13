#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod caldav;
mod error;
mod event;
mod google;
mod outlook;
mod sync;
mod updater;

const APP_VERSION: &str = git_version::git_version!(
    args = ["--tags", "--match", "v*", "--always"],
    prefix = "",
    fallback = "dev",
);

slint::include_modules!();

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
            if let Ok(exe) = std::env::current_exe() {
                let path_str: Vec<u16> = format!("{}\0", exe.display()).encode_utf16().collect();
                let bytes =
                    std::slice::from_raw_parts(path_str.as_ptr() as *const u8, path_str.len() * 2);
                let _ = RegSetValueExW(key, PCWSTR::from_raw(val_name.as_ptr()),
                                       0, REG_SZ, Some(bytes));
            }
        } else {
            let _ = RegDeleteValueW(key, PCWSTR::from_raw(val_name.as_ptr()));
        }
        let _ = RegCloseKey(key);
    }
}

#[derive(serde::Serialize, serde::Deserialize, Default)]
struct AppConfig {
    pairs:          Vec<SavedPair>,
    last_synced_at: Option<chrono::DateTime<chrono::Utc>>,
    #[serde(default = "default_interval_minutes")]
    sync_interval_minutes: u64,
}

fn config_path() -> std::path::PathBuf {
    dirs::data_local_dir()
        .unwrap_or_else(|| std::path::PathBuf::from("."))
        .join("ruscal")
        .join("config.json")
}

fn load_config() -> AppConfig {
    let path = config_path();
    std::fs::read_to_string(path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
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
    });
}

/// Save a new sync interval, preserving everything else.
fn save_interval(minutes: u64) {
    let mut config = load_config();
    config.sync_interval_minutes = minutes;
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

/// Translate raw sync errors into user-friendly messages.
fn friendly_error(e: &str) -> String {
    if e.contains("MAPI_E_LOGON_FAILED") || e.contains("0x80040111") {
        "Outlook not running — open Outlook and try again".into()
    } else if e.contains("MAPI_E_NOT_INITIALIZED") || e.contains("0x8004011F") {
        "Outlook MAPI not initialized — restart Outlook".into()
    } else if e.contains("MAPI") {
        format!("Outlook error: {e}")
    } else if e.contains("401") || e.contains("403") {
        "Google auth expired — re-open the app to re-authenticate".into()
    } else {
        format!("Sync error: {e}")
    }
}

// ── Entry point ───────────────────────────────────────────────────────────────

fn main() {
    env_logger::init();

    let launch_args = updater::parse_args();

    // In debug builds, skip self-install and single-instance so `cargo run`
    // works as a plain dev loop without touching the installed production binary.
    #[cfg(not(debug_assertions))]
    {
        // Self-install: if not running from %LOCALAPPDATA%\ruscal\ruscal.exe,
        // copy there, terminate any old instance, relaunch, exit.
        if !updater::is_installed_path() {
            let flag = launch_args.just_updated.as_deref()
                .map(|v| format!("--just-updated={v}"))
                .unwrap_or_else(|| "--just-installed".to_owned());
            updater::self_install(Some(&flag)); // diverges — exits this process
        }

        // Single-instance guard: if ruscal is already running, bring it to front.
        let _instance_guard = match updater::acquire_single_instance() {
            Some(g) => g,
            None => {
                updater::focus_existing_window();
                return;
            }
        };
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
            "Installed — ruscal will start with Windows automatically".into()
        );
    } else if let Some(ref prev_ver) = launch_args.just_updated {
        panel.set_last_sync_text(
            format!("Updated from v{prev_ver} to v{}", APP_VERSION).into()
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
            let sync_targets: Vec<(String, String)> = (0..pairs.row_count())
                .filter_map(|i| pairs.row_data(i))
                .filter(|pair| !pair.source_name.is_empty() && !pair.dest_id.is_empty())
                .map(|pair| (pair.dest_id.to_string(), pair.dest_account.to_string()))
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
                let result: Result<usize, String> = (|| {
                    let mut total = 0usize;
                    for (dest_url, google_email) in &sync_targets {
                        let token = google::get_access_token_for(google_email)
                            .map_err(|e| e.to_string())?;
                        total += sync::run_sync(dest_url, &token)?;
                    }
                    Ok(total)
                })();

                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    match result {
                        Ok(count) => {
                            let now = chrono::Utc::now();
                            *last_synced.lock().unwrap() = Some(now);
                            save_config_vec(&saved_pairs, Some(now));
                            p.set_sync_status(SyncStatus::Success);
                            p.set_last_sync_text(
                                format!("{} · {}",
                                    format_last_sync(now),
                                    format_next_sync(Some(now), app_start, secs),
                                ).into()
                            );
                            let _ = count; // surfaced in status line via last_sync
                        }
                        Err(ref e) => {
                            log::error!("Sync failed: {e}");
                            p.set_sync_status(SyncStatus::Error);
                            p.set_sync_error_detail(e.as_str().into());
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
                p.set_current_page(4);
            }
        }
    });

    panel.on_configure_source({
        let weak = panel.as_weak();
        move |i| {
            let Some(p) = weak.upgrade() else { return };
            p.set_outlook_calendars(Rc::new(VecModel::from(vec![])).into());
            p.set_config_pair_index(i);
            p.set_current_page(1);

            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = outlook::list_calendar_sources();
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    if p.get_current_page() != 1 { return; }
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
                            p.set_current_page(0);
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
            p.set_current_page(0);
        }
    });

    panel.on_configure_dest({
        let weak = panel.as_weak();
        move |i| {
            let Some(p) = weak.upgrade() else { return };
            p.set_config_pair_index(i);
            p.set_current_page(2);
        }
    });

    // Helper: navigate to page 3 and kick off a background calendar load.
    // `email_hint` = Some(email) to reuse existing tokens, None to force new OAuth.
    let start_google_picker = {
        let weak = panel.as_weak();
        move |email_hint: Option<String>| {
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
            p.set_current_page(3);

            // Hide the window while the browser OAuth flow is open.
            if email_hint.is_none() {
                p.hide().ok();
            }

            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = google::list_google_calendars(email_hint.as_deref());
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    // Bring the window back regardless of outcome.
                    snap_to_tray(&p);
                    p.show().ok();
                    if p.get_current_page() != 3 { return; }
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
            if p.get_current_page() == 3 { return; }
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
        move || {
            let Some(p) = weak.upgrade() else { return };
            p.hide().ok();
            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = google::authorize_new_account();
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
            p.set_current_page(0);
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
            let _ = std::process::Command::new("powershell")
                .args(["-NoProfile", "-NonInteractive", "-Command",
                       &format!("Set-Clipboard -Value '{escaped}'")])
                .output();
            p.set_error_copy_done(true);
            let weak2 = weak.clone();
            slint::Timer::single_shot(std::time::Duration::from_millis(1500), move || {
                if let Some(p) = weak2.upgrade() { p.set_error_copy_done(false); }
            });
        }
    });

    panel.set_app_version(APP_VERSION.into());
    panel.show().ok();


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

    // ── Sync on launch (release only — dev uses mock state) ──────────────────

    #[cfg(not(debug_assertions))]
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
