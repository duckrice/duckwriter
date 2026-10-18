// Searching the Zotero library on this computer, through Zotero's local API
// (Zotero 7: Settings → Advanced → "Allow other applications on this computer
// to communicate with Zotero"). Done here rather than in the page, which
// Zotero's server wouldn't answer (it sends no CORS headers).

use std::io::{Read, Write};
use std::net::{SocketAddr, TcpStream};
use std::time::Duration;

const ADDRESS: &str = "127.0.0.1:23119";

/// The items matching `query`, as CSL-JSON (Zotero's answer, unchanged).
#[tauri::command]
pub async fn zotero_search(query: String) -> Result<String, String> {
    tauri::async_runtime::spawn_blocking(move || search(&query))
        .await
        .map_err(|e| e.to_string())?
}

fn encode(s: &str) -> String {
    s.bytes()
        .map(|b| match b {
            b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => (b as char).to_string(),
            _ => format!("%{b:02X}"),
        })
        .collect()
}

fn search(query: &str) -> Result<String, String> {
    let address: SocketAddr = ADDRESS.parse().map_err(|e| format!("{e}"))?;
    let mut stream = TcpStream::connect_timeout(&address, Duration::from_secs(2))
        .map_err(|_| "Zotero isn't running, or doesn't let other programs in (Zotero's settings, Advanced).".to_string())?;
    stream.set_read_timeout(Some(Duration::from_secs(15))).ok();
    stream.set_write_timeout(Some(Duration::from_secs(5))).ok();
    let request = format!(
        "GET /api/users/0/items/top?q={}&qmode=titleCreatorYear&limit=30&format=csljson HTTP/1.0\r\nHost: {ADDRESS}\r\nZotero-API-Version: 3\r\nConnection: close\r\n\r\n",
        encode(query)
    );
    stream.write_all(request.as_bytes()).map_err(|e| e.to_string())?;
    let mut response = Vec::new();
    stream.read_to_end(&mut response).map_err(|e| e.to_string())?;
    let text = String::from_utf8_lossy(&response);
    let (head, body) = text.split_once("\r\n\r\n").ok_or("Zotero gave an answer DuckWriter can't read.")?;
    let status = head.split_whitespace().nth(1).unwrap_or("");
    match status {
        "200" => Ok(body.to_string()),
        "403" | "404" => Err("Zotero's local API is off: in Zotero, Settings → Advanced, allow other applications on this computer to communicate with Zotero.".into()),
        _ => Err(format!("Zotero answered {status}.")),
    }
}

#[cfg(test)]
mod tests {
    use super::encode;

    #[test]
    fn encodes_queries() {
        assert_eq!(encode("Smith 2020"), "Smith%202020");
        assert_eq!(encode("ø&x"), "%C3%B8%26x");
    }
}
