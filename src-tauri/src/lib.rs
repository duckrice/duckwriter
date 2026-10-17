use std::{
    fs,
    path::{Path, PathBuf},
};

use percent_encoding::percent_decode_str;
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::Manager;

/// The documents DuckWriter was started with ("Open with" or a double-click), as absolute paths.
#[tauri::command]
fn startup_files() -> Vec<String> {
    let cwd = std::env::current_dir().unwrap_or_default();
    std::env::args_os()
        .skip(1)
        .map(PathBuf::from)
        .map(|path| if path.is_absolute() { path } else { cwd.join(path) })
        .filter(|path| path.is_file())
        .map(|path| path.to_string_lossy().into_owned())
        .collect()
}

/// The folder for recovery copies of unsaved work, on this computer only
/// (%LOCALAPPDATA%\com.duckrice.duckwriter\Recovery). Created if missing.
#[tauri::command]
fn recovery_dir(app: tauri::AppHandle) -> Result<String, String> {
    // Tests use a folder of their own, so they never touch real recovery copies.
    let dir = match std::env::var_os("DUCKWRITER_RECOVERY_DIR") {
        Some(dir) => PathBuf::from(dir),
        None => app.path().app_local_data_dir().map_err(|err| err.to_string())?.join("Recovery"),
    };
    fs::create_dir_all(&dir).map_err(|err| format!("Could not create {}: {err}", dir.display()))?;
    Ok(dir.to_string_lossy().into_owned())
}

/// The files in the recovery folder.
#[tauri::command]
fn recovery_files(app: tauri::AppHandle) -> Result<Vec<String>, String> {
    let dir = PathBuf::from(recovery_dir(app)?);
    let entries = fs::read_dir(&dir).map_err(|err| err.to_string())?;
    Ok(entries
        .filter_map(Result::ok)
        .filter(|entry| entry.path().is_file())
        .map(|entry| entry.path().to_string_lossy().into_owned())
        .collect())
}

/// Deletes a recovery copy. Only files in the recovery folder can be deleted this way.
#[tauri::command]
fn delete_recovery(app: tauri::AppHandle, path: String) -> Result<(), String> {
    let dir = PathBuf::from(recovery_dir(app)?);
    let file = PathBuf::from(&path);
    if file.parent() != Some(dir.as_path()) {
        return Err(format!("Not a recovery file: {path}"));
    }
    match fs::remove_file(&file) {
        Ok(()) => Ok(()),
        Err(err) if err.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(err) => Err(err.to_string()),
    }
}

#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    fs::read(&path)
        .map(Response::new)
        .map_err(|err| format!("Could not read {path}: {err}"))
}

/// Saves the request body to the path given in the `x-path` header (percent-encoded,
/// since header values must be ASCII). The bytes travel as a raw body so large
/// documents aren't serialised as JSON.
#[tauri::command]
fn write_file(request: Request<'_>) -> Result<(), String> {
    let InvokeBody::Raw(bytes) = request.body() else {
        return Err("write_file expects a binary body".into());
    };
    let encoded = request
        .headers()
        .get("x-path")
        .and_then(|value| value.to_str().ok())
        .ok_or("write_file is missing the x-path header")?;
    let path = PathBuf::from(
        percent_decode_str(encoded)
            .decode_utf8()
            .map_err(|err| err.to_string())?
            .into_owned(),
    );
    write_atomically(&path, bytes).map_err(|err| format!("Could not save {}: {err}", path.display()))
}

/// Writes next to the target and renames over it, so a failed save can't leave
/// a half-written document behind.
fn write_atomically(path: &Path, bytes: &[u8]) -> std::io::Result<()> {
    let mut name = path.file_name().unwrap_or_default().to_os_string();
    name.push(format!(".{}.tmp", std::process::id()));
    let tmp = path.with_file_name(name);
    fs::write(&tmp, bytes)?;
    fs::rename(&tmp, path).inspect_err(|_| {
        let _ = fs::remove_file(&tmp);
    })
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .setup(|app| {
            let config = &app.config().app.windows[0];
            let builder = tauri::WebviewWindowBuilder::from_config(app.handle(), config)?;
            // Tests open the window off-screen and in the background, so they
            // can run without taking over the screen. A window can't be created
            // off-screen, so it's created hidden, moved, then shown without focus.
            if std::env::var_os("DUCKWRITER_TEST_OFFSCREEN").is_some() {
                let window = builder.visible(false).focused(false).skip_taskbar(true).build()?;
                window.set_position(tauri::PhysicalPosition::new(-20000, -20000))?;
                window.show()?;
            } else {
                builder.build()?;
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![startup_files, read_file, write_file, recovery_dir, recovery_files, delete_recovery])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
