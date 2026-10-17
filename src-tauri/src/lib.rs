mod spelling;

use std::{
    fs,
    path::{Path, PathBuf},
};

use percent_encoding::percent_decode_str;
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::Manager;

/// Opens a link in the default browser (or mail program). Only web and mail
/// addresses: a document's link can't start a program.
#[tauri::command]
fn open_url(url: String) -> Result<(), String> {
    let lower = url.to_ascii_lowercase();
    if !(lower.starts_with("http://") || lower.starts_with("https://") || lower.starts_with("mailto:")) {
        return Err("Only web and e-mail addresses can be opened.".into());
    }
    #[cfg(windows)]
    unsafe {
        use windows::core::{w, HSTRING};
        use windows::Win32::UI::Shell::ShellExecuteW;
        use windows::Win32::UI::WindowsAndMessaging::SW_SHOWNORMAL;
        let result = ShellExecuteW(None, w!("open"), &HSTRING::from(url.as_str()), None, None, SW_SHOWNORMAL);
        if result.0 as isize <= 32 {
            return Err("The link couldn't be opened.".into());
        }
    }
    Ok(())
}

/// Builds a window; for tests, off-screen and in the background. A window can't
/// be created off-screen, so it's created hidden, moved, then shown without focus.
fn build_window(builder: tauri::WebviewWindowBuilder<'_, tauri::Wry, tauri::AppHandle>) -> tauri::Result<()> {
    if std::env::var_os("DUCKWRITER_TEST_OFFSCREEN").is_some() {
        let window = builder.visible(false).focused(false).skip_taskbar(true).build()?;
        window.set_position(tauri::PhysicalPosition::new(-20000, -20000))?;
        window.show()?;
    } else {
        builder.build()?;
    }
    Ok(())
}

/// Opens another document (or a new one, with no path) in a window of its own, as Word does.
/// Async: on Windows a window made from a command on the main thread can hang.
#[tauri::command]
async fn open_window(app: tauri::AppHandle, path: Option<String>) -> Result<(), String> {
    static COUNT: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
    let n = COUNT.fetch_add(1, std::sync::atomic::Ordering::Relaxed) + 1;
    let mut config = app.config().app.windows[0].clone();
    config.label = format!("doc-{}-{n}", std::process::id());
    let open = serde_json::to_string(&path.unwrap_or_default()).map_err(|e| e.to_string())?;
    let builder = tauri::WebviewWindowBuilder::from_config(&app, &config)
        .map_err(|e| e.to_string())?
        .initialization_script(format!("window.__DUCKWRITER_OPEN = {open};"));
    build_window(builder).map_err(|e| e.to_string())
}

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

/// Saves what the window shows as a PDF, with WebView2's own PDF printer: no
/// dialog, no margins (the pages lay themselves out), page size in inches.
#[tauri::command]
async fn save_pdf(window: tauri::WebviewWindow, path: String, width: f64, height: f64) -> Result<(), String> {
    let (tx, rx) = std::sync::mpsc::channel::<Result<(), String>>();
    window
        .with_webview(move |webview| {
            #[cfg(windows)]
            unsafe {
                use webview2_com::Microsoft::Web::WebView2::Win32::{ICoreWebView2Environment6, ICoreWebView2_7};
                use windows::core::{Interface, HSTRING};
                let done = tx.clone();
                let started = (|| -> windows::core::Result<()> {
                    let core = webview.controller().CoreWebView2()?;
                    let core7: ICoreWebView2_7 = core.cast()?;
                    let environment: ICoreWebView2Environment6 = webview.environment().cast()?;
                    let settings = environment.CreatePrintSettings()?;
                    settings.SetPageWidth(width)?;
                    settings.SetPageHeight(height)?;
                    settings.SetMarginTop(0.0)?;
                    settings.SetMarginBottom(0.0)?;
                    settings.SetMarginLeft(0.0)?;
                    settings.SetMarginRight(0.0)?;
                    settings.SetScaleFactor(1.0)?;
                    settings.SetShouldPrintBackgrounds(true)?;
                    settings.SetShouldPrintHeaderAndFooter(false)?;
                    let handler = webview2_com::PrintToPdfCompletedHandler::create(Box::new(move |result, ok| {
                        let _ = done.send(if result.is_ok() && ok { Ok(()) } else { Err("The PDF could not be written.".into()) });
                        Ok(())
                    }));
                    core7.PrintToPdf(&HSTRING::from(path.as_str()), &settings, &handler)
                })();
                if let Err(err) = started {
                    let _ = tx.send(Err(err.to_string()));
                }
            }
            #[cfg(not(windows))]
            let _ = tx.send(Err("Saving as PDF needs Windows.".into()));
        })
        .map_err(|err| err.to_string())?;
    tauri::async_runtime::spawn_blocking(move || rx.recv().map_err(|err| err.to_string())?)
        .await
        .map_err(|err| err.to_string())?
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
            build_window(tauri::WebviewWindowBuilder::from_config(app.handle(), config)?)?;
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![startup_files, read_file, write_file, recovery_dir, recovery_files, delete_recovery, save_pdf, spelling::spell_languages, spelling::spell_check, spelling::spell_suggest, open_url, open_window])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
