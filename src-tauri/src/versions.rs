//! Earlier versions of documents: before a document is saved over, the file as
//! it was is copied to a folder of DuckWriter's own on this computer
//! (%LOCALAPPDATA%\com.duckrice.duckwriter\Versions), the newest 20 kept.

use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};

use tauri::Manager;

const KEEP: usize = 20;

fn versions_root(app: &tauri::AppHandle) -> Result<PathBuf, String> {
    // Tests keep theirs next to their recovery folder.
    Ok(match std::env::var_os("DUCKWRITER_RECOVERY_DIR") {
        Some(dir) => PathBuf::from(dir).with_file_name("Versions"),
        None => app.path().app_local_data_dir().map_err(|e| e.to_string())?.join("Versions"),
    })
}

/// A folder per document, named after its path (FNV-1a), with the path inside.
fn folder_for(app: &tauri::AppHandle, document: &str) -> Result<PathBuf, String> {
    let mut hash: u64 = 0xcbf29ce484222325;
    for byte in document.to_lowercase().bytes() {
        hash ^= byte as u64;
        hash = hash.wrapping_mul(0x100000001b3);
    }
    Ok(versions_root(app)?.join(format!("{hash:016x}")))
}

fn now_ms() -> u128 {
    SystemTime::now().duration_since(UNIX_EPOCH).map(|d| d.as_millis()).unwrap_or(0)
}

fn versions_in(dir: &Path) -> Vec<(u128, PathBuf)> {
    let mut out: Vec<(u128, PathBuf)> = fs::read_dir(dir)
        .map(|entries| {
            entries
                .filter_map(|e| e.ok())
                .filter_map(|e| {
                    let name = e.file_name().to_string_lossy().into_owned();
                    let ms = name.strip_suffix(".docx")?.parse::<u128>().ok()?;
                    Some((ms, e.path()))
                })
                .collect()
        })
        .unwrap_or_default();
    out.sort_by(|a, b| b.0.cmp(&a.0));
    out
}

/// Keeps the document's file as it is now, before it is saved over; unless the
/// newest kept version is younger than `min_age_secs` (auto-save keeps fewer).
#[tauri::command]
pub fn versions_keep(app: tauri::AppHandle, path: String, min_age_secs: u64) -> Result<(), String> {
    let source = PathBuf::from(&path);
    if !source.is_file() {
        return Ok(());
    }
    let dir = folder_for(&app, &path)?;
    let kept = versions_in(&dir);
    if let Some((newest, _)) = kept.first() {
        if now_ms().saturating_sub(*newest) < (min_age_secs as u128) * 1000 {
            return Ok(());
        }
    }
    fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    fs::write(dir.join("document.txt"), &path).map_err(|e| e.to_string())?;
    fs::copy(&source, dir.join(format!("{}.docx", now_ms()))).map_err(|e| e.to_string())?;
    for (_, old) in versions_in(&dir).into_iter().skip(KEEP) {
        let _ = fs::remove_file(old);
    }
    Ok(())
}

#[derive(serde::Serialize)]
pub struct Version {
    path: String,
    /// When it was kept (milliseconds since 1970).
    time: u64,
    size: u64,
}

/// The kept versions of a document, newest first.
#[tauri::command]
pub fn versions_list(app: tauri::AppHandle, path: String) -> Result<Vec<Version>, String> {
    let dir = folder_for(&app, &path)?;
    Ok(versions_in(&dir)
        .into_iter()
        .map(|(ms, p)| Version { size: fs::metadata(&p).map(|m| m.len()).unwrap_or(0), path: p.to_string_lossy().into_owned(), time: ms as u64 })
        .collect())
}
