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

    // Mock pair to preview the layout.
    let outlook_color = slint::Color::from_rgb_u8(0, 114, 198);
    let gcal_color    = slint::Color::from_rgb_u8(66, 133, 244);

    let pairs_model = Rc::new(VecModel::from(vec![
        // Empty placeholder — always one at the end.
        SyncPair {
            source_name:    "".into(),
            source_account: "".into(),
            source_color:   outlook_color,
            dest_name:      "".into(),
            dest_account:   "".into(),
            dest_color:     gcal_color,
            dest_id:        "".into(),
        },
    ]));

    panel.set_pairs(pairs_model.clone().into());
    panel.set_sync_status(SyncStatus::Success);
    panel.set_last_sync_text("Synced · just now".into());

    panel.on_sync_now({
        let weak  = panel.as_weak();
        let pairs = pairs_model.clone();
        move || {
            let p = weak.unwrap();

            // Collect dest URLs of fully-configured pairs.
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
                            p.set_sync_status(SyncStatus::Success);
                            p.set_last_sync_text(
                                format!("Synced {count} events · just now").into()
                            );
                        }
                        Err(e) => {
                            log::error!("Sync failed: {e}");
                            p.set_sync_status(SyncStatus::Error);
                            p.set_last_sync_text(format!("Sync error: {e}").into());
                        }
                    }
                }).ok();
            });
        }
    });

    panel.on_minimize({
        let weak = panel.as_weak();
        move || { weak.unwrap().hide().unwrap(); }
    });

    panel.on_quit(|| slint::quit_event_loop().unwrap());

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
                            p.set_last_sync_text("Could not read Outlook calendars".into());
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
                // New placeholder row is added only once the pair is fully
                // configured (both source and dest). That happens in on_dest_selected.
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
                                        .map(|c: &google::GoogleCalendar| slint::SharedString::from(&c.summary))
                                        .collect::<Vec<_>>()
                                )
                            );
                            let ids: Rc<VecModel<slint::SharedString>> = Rc::new(
                                VecModel::from(
                                    calendars.iter()
                                        .map(|c: &google::GoogleCalendar| slint::SharedString::from(&c.id))
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

            p.set_current_page(0);
        }
    });

    panel.show().expect("failed to show panel");

    let weak_for_pos = panel.as_weak();
    let _pos_timer = slint::Timer::default();
    _pos_timer.start(
        slint::TimerMode::SingleShot,
        std::time::Duration::from_millis(16),
        move || {
            if let Some(p) = weak_for_pos.upgrade() { snap_to_tray(&p); }
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
                        p.show().unwrap();
                    }
                }
            }

            while let Ok(ev) = MenuEvent::receiver().try_recv() {
                if ev.id() == &sync_id {
                    if let Some(p) = panel_weak.upgrade() {
                        p.set_sync_status(SyncStatus::Syncing);
                        p.set_last_sync_text("Syncing…".into());
                    }
                } else if ev.id() == &quit_id {
                    slint::quit_event_loop().unwrap();
                }
            }
        },
    );

    slint::run_event_loop_until_quit().expect("event loop failed");
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
