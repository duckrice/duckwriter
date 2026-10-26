//! Office's own fonts. Word's newer default fonts (Aptos and others) are "cloud fonts": Office
//! downloads them into its font cache instead of installing them in Windows, so other programs
//! can't see them. DuckWriter lists them here and shows documents in them, so lines and pages
//! break where Word breaks them. Nothing is copied; without Office the list is simply empty.

use serde::Serialize;
use std::path::Path;

#[derive(Serialize)]
pub struct OfficeFont {
    family: String,
    weight: u16,
    italic: bool,
    path: String,
}

/// usWeightClass and the italic bit of a font file's OS/2 table.
fn style_of(bytes: &[u8]) -> Option<(u16, bool)> {
    let u16_at = |o: usize| bytes.get(o..o + 2).map(|b| u16::from_be_bytes([b[0], b[1]]));
    let u32_at = |o: usize| bytes.get(o..o + 4).map(|b| u32::from_be_bytes([b[0], b[1], b[2], b[3]]));
    let tables = u16_at(4)? as usize;
    for i in 0..tables {
        let record = 12 + i * 16;
        if bytes.get(record..record + 4)? == b"OS/2" {
            let os2 = u32_at(record + 8)? as usize;
            return Some((u16_at(os2 + 4)?, u16_at(os2 + 62)? & 1 == 1));
        }
    }
    None
}

#[tauri::command]
pub async fn office_fonts() -> Vec<OfficeFont> {
    tauri::async_runtime::spawn_blocking(|| {
        let Some(base) = std::env::var_os("LOCALAPPDATA") else { return vec![] };
        let root = Path::new(&base).join(r"Microsoft\FontCache\4\CloudFonts");
        let mut fonts = vec![];
        for family in std::fs::read_dir(root).into_iter().flatten().flatten() {
            let name = family.file_name().to_string_lossy().to_string();
            for file in std::fs::read_dir(family.path()).into_iter().flatten().flatten() {
                let path = file.path();
                if !path.extension().is_some_and(|e| e.eq_ignore_ascii_case("ttf") || e.eq_ignore_ascii_case("otf")) {
                    continue;
                }
                // shortcut: reads the whole file for its OS/2 table (a few MB in all, once per start); read just the header if it gets slow.
                let Ok(bytes) = std::fs::read(&path) else { continue };
                if let Some((weight, italic)) = style_of(&bytes) {
                    fonts.push(OfficeFont { family: name.clone(), weight, italic, path: path.to_string_lossy().to_string() });
                }
            }
        }
        fonts
    })
    .await
    .unwrap_or_default()
}
