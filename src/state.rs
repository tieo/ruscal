/// Per-pair persistent sync state: hash cache.
///
/// Stored as a single JSON file at `%LOCALAPPDATA%\ruscal\state.json` with
/// one entry per sync pair (source+destination). The pair-id key is
/// `{source_account}__{dest_id}`, so renaming either side invalidates only
/// that pair — not all of them.
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::PathBuf;

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct PairState {
    /// `{uid}` → hash of the last successfully PUT iCalendar body. Used to
    /// skip identical re-PUTs (avoids Google's 409 on recurring masters).
    #[serde(default)]
    pub hash_cache: HashMap<String, u64>,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct AppState {
    #[serde(default)]
    pub pairs: HashMap<String, PairState>,
}

fn state_path() -> Option<PathBuf> {
    dirs::data_local_dir().map(|d| d.join("ruscal").join("state.json"))
}

pub fn pair_id(source_account: &str, dest_id: &str) -> String {
    format!("{source_account}__{dest_id}")
}

pub fn load() -> AppState {
    cleanup_legacy_caches();
    let Some(path) = state_path() else { return AppState::default() };
    std::fs::read_to_string(&path).ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

/// Remove pre-unified `sync_cache_<hex>.json` files from older versions.
/// Idempotent: if the directory is absent or no files match, it's a no-op.
fn cleanup_legacy_caches() {
    let Some(dir) = dirs::data_local_dir().map(|d| d.join("ruscal")) else { return };
    let Ok(entries) = std::fs::read_dir(&dir) else { return };
    for entry in entries.flatten() {
        let name = entry.file_name();
        let Some(name) = name.to_str() else { continue };
        if name.starts_with("sync_cache_") && name.ends_with(".json") {
            let _ = std::fs::remove_file(entry.path());
        }
    }
}

pub fn save(app: &AppState) {
    let Some(path) = state_path() else { return };
    if let Some(parent) = path.parent() {
        let _ = std::fs::create_dir_all(parent);
    }
    if let Ok(s) = serde_json::to_string_pretty(app) {
        let _ = std::fs::write(path, s);
    }
}

/// Convenience: load just one pair's state.
pub fn load_pair(pair_id: &str) -> PairState {
    load().pairs.remove(pair_id).unwrap_or_default()
}

/// Convenience: save one pair's state, preserving other pairs.
pub fn save_pair(pair_id: &str, state: PairState) {
    let mut app = load();
    app.pairs.insert(pair_id.to_owned(), state);
    save(&app);
}
