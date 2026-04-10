mod caldav;
mod error;
mod event;
mod google;
mod outlook;
mod sync;

slint::include_modules!();

use std::rc::Rc;
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

/// Serializable mirror of a configured sync pair.
/// Slint-generated types cannot be annotated with serde directly.
#[derive(serde::Serialize, serde::Deserialize, Clone, Default)]
struct SavedPair {
    source_account: String, // Outlook calendar display name
    dest_account:   String, // Google Calendar summary
    dest_id:        String, // Google Calendar CalDAV URL
}

#[derive(serde::Serialize, serde::Deserialize, Default)]
struct AppConfig {
    pairs:          Vec<SavedPair>,
    /// UTC timestamp of the last successful sync, ISO-8601.
    last_synced_at: Option<chrono::DateTime<chrono::Utc>>,
}

fn config_path() -> std::path::PathBuf {
    dirs::config_dir()
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

fn save_config(pairs: &VecModel<SyncPair>, last_synced_at: Option<chrono::DateTime<chrono::Utc>>) {
    let saved: Vec<SavedPair> = (0..pairs.row_count())
        .filter_map(|i| pairs.row_data(i))
        .filter(|p| !p.dest_id.is_empty()) // only fully configured pairs
        .map(|p| SavedPair {
            source_account: p.source_account.to_string(),
            dest_account:   p.dest_account.to_string(),
            dest_id:        p.dest_id.to_string(),
        })
        .collect();

    let config = AppConfig { pairs: saved, last_synced_at };
    let path = config_path();
    if let Some(dir) = path.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    if let Ok(json) = serde_json::to_string_pretty(&config) {
        let _ = std::fs::write(path, json);
    }
}

// ── UI helpers ────────────────────────────────────────────────────────────────

/// Format a UTC timestamp as a human-readable "X ago" string.
fn format_time_ago(dt: chrono::DateTime<chrono::Utc>) -> String {
    let secs = chrono::Utc::now()
        .signed_duration_since(dt)
        .num_seconds()
        .max(0);
    if secs < 60 {
        "Synced · just now".into()
    } else if secs < 3600 {
        format!("Synced · {}m ago", secs / 60)
    } else if secs < 86400 {
        format!("Synced · {}h ago", secs / 3600)
    } else {
        format!("Synced · {}d ago", secs / 86400)
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

    unsafe {
        let _ = SetProcessDpiAwarenessContext(
            DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2
        );
    }

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

    // Load persisted config and build the initial pairs list.
    let config = load_config();

    let mut initial_pairs: Vec<SyncPair> = config.pairs.iter().map(|p| SyncPair {
        source_name:    "Outlook Calendar".into(),
        source_account: p.source_account.as_str().into(),
        source_color:   outlook_color,
        dest_name:      "Google Calendar".into(),
        dest_account:   p.dest_account.as_str().into(),
        dest_color:     gcal_color,
        dest_id:        p.dest_id.as_str().into(),
    }).collect();

    // Always keep a blank placeholder at the end for adding new pairs.
    initial_pairs.push(SyncPair {
        source_name:    "".into(),
        source_account: "".into(),
        source_color:   outlook_color,
        dest_name:      "".into(),
        dest_account:   "".into(),
        dest_color:     gcal_color,
        dest_id:        "".into(),
    });

    let pairs_model = Rc::new(VecModel::from(initial_pairs));
    panel.set_pairs(pairs_model.clone().into());
    panel.set_sync_status(SyncStatus::Idle);

    // Show the last successful sync time if we have it.
    let initial_status = match config.last_synced_at {
        Some(dt) if !config.pairs.is_empty() => format_time_ago(dt),
        _ if config.pairs.is_empty()         => "No calendars configured".into(),
        _                                    => "Never synced".into(),
    };
    panel.set_last_sync_text(initial_status.into());

    // ── Sync callback (used by both the UI button and the tray menu) ──────────

    panel.on_sync_now({
        let weak  = panel.as_weak();
        let pairs = pairs_model.clone();
        move || {
            let Some(p) = weak.upgrade() else { return };

            // Don't start a second sync while one is already running.
            if p.get_sync_status() == SyncStatus::Syncing { return; }

            // Collect URLs of fully-configured pairs.
            let dest_urls: Vec<String> = (0..pairs.row_count())
                .filter_map(|i| pairs.row_data(i))
                .filter(|pair| !pair.source_name.is_empty() && !pair.dest_id.is_empty())
                .map(|pair| pair.dest_id.to_string())
                .collect();

            if dest_urls.is_empty() {
                p.set_sync_status(SyncStatus::Idle);
                p.set_last_sync_text("No calendars configured".into());
                return;
            }

            // Snapshot pairs for saving in the completion callback (thread-safe).
            let saved_pairs: Vec<SavedPair> = (0..pairs.row_count())
                .filter_map(|i| pairs.row_data(i))
                .filter(|pair| !pair.dest_id.is_empty())
                .map(|pair| SavedPair {
                    source_account: pair.source_account.to_string(),
                    dest_account:   pair.dest_account.to_string(),
                    dest_id:        pair.dest_id.to_string(),
                })
                .collect();

            p.set_sync_status(SyncStatus::Syncing);
            p.set_last_sync_text("Syncing…".into());

            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result: Result<usize, String> = (|| {
                    let token = google::get_access_token().map_err(|e| e.to_string())?;
                    let mut total = 0usize;
                    for dest_url in &dest_urls {
                        total += sync::run_sync(dest_url, &token)?;
                    }
                    Ok(total)
                })();

                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    match result {
                        Ok(count) => {
                            let now = chrono::Utc::now();
                            save_config_vec(&saved_pairs, Some(now));
                            p.set_sync_status(SyncStatus::Success);
                            p.set_last_sync_text(
                                format!("Synced {count} events · just now").into()
                            );
                        }
                        Err(ref e) => {
                            log::error!("Sync failed: {e}");
                            p.set_sync_status(SyncStatus::Error);
                            p.set_last_sync_text(friendly_error(e).into());
                        }
                    }
                }).ok();
            });
        }
    });

    panel.on_minimize({
        let weak = panel.as_weak();
        move || { if let Some(p) = weak.upgrade() { p.hide().ok(); } }
    });

    panel.on_quit(|| { slint::quit_event_loop().ok(); });

    panel.on_open_settings({
        let weak = panel.as_weak();
        move || { toast(&weak, "Settings coming soon"); }
    });

    panel.on_configure_source({
        let weak = panel.as_weak();
        move |i| {
            let Some(p) = weak.upgrade() else { return };
            // Navigate immediately — gives instant feedback, prevents double-click.
            p.set_outlook_calendars(Rc::new(VecModel::from(vec![])).into());
            p.set_config_pair_index(i);
            p.set_current_page(1);

            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = outlook::list_calendar_sources();
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    // User may have navigated back — don't overwrite their state.
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
            let pair_index  = p.get_config_pair_index() as usize;
            let calendars   = p.get_outlook_calendars();
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
            p.set_current_page(2); // destination provider picker
        }
    });

    panel.on_dest_provider_selected({
        let weak = panel.as_weak();
        move |_provider_index| {
            // Only Google Calendar (index 0) for now.
            let Some(p) = weak.upgrade() else { return };
            // Guard against double-click while already loading.
            if p.get_current_page() == 3 { return; }

            p.set_google_calendars(Rc::new(VecModel::from(vec![])).into());
            p.set_google_calendar_ids(Rc::new(VecModel::from(vec![])).into());
            p.set_google_status("Waiting for Google authorization in your browser…".into());
            p.set_current_page(3);

            let weak2 = weak.clone();
            std::thread::spawn(move || {
                let result = google::list_google_calendars();
                slint::invoke_from_event_loop(move || {
                    let Some(p) = weak2.upgrade() else { return };
                    if p.get_current_page() != 3 { return; }
                    match result {
                        Ok(calendars) => {
                            let names: Rc<VecModel<slint::SharedString>> = Rc::new(
                                VecModel::from(
                                    calendars.iter()
                                        .map(|c: &google::GoogleCalendar| {
                                            slint::SharedString::from(&c.summary)
                                        })
                                        .collect::<Vec<_>>()
                                )
                            );
                            let ids: Rc<VecModel<slint::SharedString>> = Rc::new(
                                VecModel::from(
                                    calendars.iter()
                                        .map(|c: &google::GoogleCalendar| {
                                            slint::SharedString::from(&c.id)
                                        })
                                        .collect::<Vec<_>>()
                                )
                            );
                            p.set_google_calendars(names.into());
                            p.set_google_calendar_ids(ids.into());
                            p.set_google_status("".into());
                        }
                        Err(e) => {
                            log::error!("Failed to list Google Calendars: {e}");
                            p.set_google_status(format!("Error: {e}").into());
                        }
                    }
                }).ok();
            });
        }
    });

    panel.on_dest_selected({
        let weak  = panel.as_weak();
        let pairs = pairs_model.clone();
        move |cal_index| {
            let Some(p) = weak.upgrade() else { return };
            let pair_index = p.get_config_pair_index() as usize;

            let cal_ids   = p.get_google_calendar_ids();
            let calendars = p.get_google_calendars();
            let Some(summary) = calendars.row_data(cal_index as usize) else { return };
            let dest_id = cal_ids.row_data(cal_index as usize).unwrap_or_default();

            let mut pair = pairs.row_data(pair_index).unwrap_or_default();
            pair.dest_name    = "Google Calendar".into();
            pair.dest_account = summary;
            pair.dest_color   = gcal_color;
            pair.dest_id      = dest_id;
            pairs.set_row_data(pair_index, pair.clone());

            // Once both source and dest are configured and this was the last
            // pair, append a fresh empty placeholder for the next one.
            if pair_index + 1 >= pairs.row_count()
                && !pair.source_name.is_empty()
                && !pair.dest_name.is_empty()
            {
                pairs.push(SyncPair {
                    source_name:    "".into(),
                    source_account: "".into(),
                    source_color:   outlook_color,
                    dest_name:      "".into(),
                    dest_account:   "".into(),
                    dest_color:     gcal_color,
                    dest_id:        "".into(),
                });
            }

            // Persist the updated pair configuration.
            save_config(&pairs, None);
            p.set_current_page(0);
        }
    });

    panel.show().ok();

    // ── Snap panel to bottom-right corner of work area ────────────────────────

    let weak_for_pos = panel.as_weak();
    let _pos_timer = slint::Timer::default();
    _pos_timer.start(
        slint::TimerMode::SingleShot,
        std::time::Duration::from_millis(16),
        move || {
            if let Some(p) = weak_for_pos.upgrade() { snap_to_tray(&p); }
        },
    );

    // ── Automatic background sync every 15 minutes ────────────────────────────

    let auto_sync_weak = panel.as_weak();
    let _auto_sync_timer = slint::Timer::default();
    _auto_sync_timer.start(
        slint::TimerMode::Repeated,
        std::time::Duration::from_secs(15 * 60),
        move || {
            if let Some(p) = auto_sync_weak.upgrade() {
                // invoke_sync_now() calls the on_sync_now callback, which already
                // guards against double-starts (checks SyncStatus::Syncing).
                p.invoke_sync_now();
            }
        },
    );

    // ── Poll tray + menu events ───────────────────────────────────────────────

    let sync_id    = sync_item.id().clone();
    let quit_id    = quit_item.id().clone();
    let panel_weak = panel.as_weak();

    let _timer = slint::Timer::default();
    _timer.start(
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
                    // Invoke the same sync callback as the UI button.
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

/// Save config from a pre-collected Vec<SavedPair> — used from sync threads.
fn save_config_vec(pairs: &[SavedPair], last_synced_at: Option<chrono::DateTime<chrono::Utc>>) {
    let config = AppConfig {
        pairs: pairs.to_vec(),
        last_synced_at,
    };
    let path = config_path();
    if let Some(dir) = path.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    if let Ok(json) = serde_json::to_string_pretty(&config) {
        let _ = std::fs::write(path, json);
    }
}

fn toast(weak: &slint::Weak<TrayPanel>, msg: &'static str) {
    if let Some(p) = weak.upgrade() {
        let prev = p.get_last_sync_text().to_string();
        p.set_last_sync_text(msg.into());
        let weak2 = weak.clone();
        slint::Timer::single_shot(std::time::Duration::from_secs(2), move || {
            if let Some(p2) = weak2.upgrade() {
                p2.set_last_sync_text(prev.into());
            }
        });
    }
}

fn snap_to_tray(panel: &TrayPanel) {
    let work  = work_area_phys();
    let size  = panel.window().size();
    let scale = panel.window().scale_factor();
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
