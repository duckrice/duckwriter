//! Writing together. Two ways to pass messages between the people editing a
//! document, both carrying data the app has already encrypted:
//!
//! - On the local network: the sharing person's DuckWriter runs a small
//!   WebSocket hub that everyone (the sharer too) connects to.
//! - Through a shared folder (OneDrive, Dropbox, a network drive): each person
//!   appends to a log file of their own in a ".duckwriter-live" folder and reads
//!   the others'.
//!
//! The hub's rules are the same as those of the relay server (server/relay.mjs):
//! one "host" per room; guests talk to the host; the host talks to one guest or
//! to all of them.

use std::collections::HashMap;
use std::io::{Read, Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};

use futures_util::{SinkExt, StreamExt};
use tokio::net::TcpListener;
use tokio::sync::{mpsc, oneshot};
use tokio_tungstenite::tungstenite::handshake::server::{Request, Response};
use tokio_tungstenite::tungstenite::Message;

type Outbox = mpsc::UnboundedSender<String>;

#[derive(Default)]
struct Room {
    host: Option<Outbox>,
    guests: HashMap<String, Outbox>,
    next: u32,
}

struct Server {
    port: u16,
    stop: oneshot::Sender<()>,
}

static SERVER: Mutex<Option<Server>> = Mutex::new(None);

#[derive(serde::Serialize)]
pub struct LanInfo {
    port: u16,
    /// This computer's address on the local network, for the others to connect to.
    address: String,
}

/// The address others on the network reach this computer at. (Connecting a UDP
/// socket sends nothing; it only picks the network interface.)
fn local_address() -> String {
    std::net::UdpSocket::bind("0.0.0.0:0")
        .and_then(|s| s.connect("8.8.8.8:80").map(|_| s))
        .and_then(|s| s.local_addr())
        .map(|a| a.ip().to_string())
        .unwrap_or_else(|_| "127.0.0.1".into())
}

/// Starts the hub for room `room` (the first free port from `port` up).
#[tauri::command]
pub async fn lan_host_start(room: String, port: u16) -> Result<LanInfo, String> {
    if let Some(server) = SERVER.lock().unwrap().as_ref() {
        return Ok(LanInfo { port: server.port, address: local_address() });
    }
    // Tests share with this computer only, so Windows doesn't ask about the firewall.
    let testing = std::env::var_os("DUCKWRITER_TEST_OFFSCREEN").is_some();
    let bind = if testing { "127.0.0.1" } else { "0.0.0.0" };
    let mut listener = None;
    for p in port..port.saturating_add(20) {
        if let Ok(l) = TcpListener::bind((bind, p)).await {
            listener = Some(l);
            break;
        }
    }
    let listener = listener.ok_or("No free network port to share on.")?;
    let port = listener.local_addr().map_err(|e| e.to_string())?.port();
    let (stop, mut stopped) = oneshot::channel();
    let state = Arc::new(Mutex::new(Room::default()));
    tauri::async_runtime::spawn(async move {
        loop {
            tokio::select! {
                _ = &mut stopped => break,
                accepted = listener.accept() => {
                    let Ok((stream, peer)) = accepted else { continue };
                    let state = state.clone();
                    let room = room.clone();
                    tauri::async_runtime::spawn(async move {
                        let _ = serve(stream, peer.ip().is_loopback(), state, &room).await;
                    });
                }
            }
        }
    });
    *SERVER.lock().unwrap() = Some(Server { port, stop });
    Ok(LanInfo { port, address: if testing { "127.0.0.1".into() } else { local_address() } })
}

#[tauri::command]
pub fn lan_host_stop() {
    if let Some(server) = SERVER.lock().unwrap().take() {
        let _ = server.stop.send(());
    }
}

/// One connection: "/r/<room>?role=host" (only from this computer) or "?role=guest".
async fn serve(stream: tokio::net::TcpStream, local: bool, state: Arc<Mutex<Room>>, room: &str) -> Result<(), String> {
    let mut path = String::new();
    let ws = tokio_tungstenite::accept_hdr_async(stream, |req: &Request, resp: Response| {
        path = req.uri().to_string();
        Ok(resp)
    })
    .await
    .map_err(|e| e.to_string())?;
    let (mut sink, mut source) = ws.split();
    let (tx, mut rx) = mpsc::unbounded_channel::<String>();
    let right_room = path.starts_with(&format!("/r/{room}?")) || path == format!("/r/{room}");
    let wants_host = path.contains("role=host");

    // Who this is, or why not.
    let me: Result<String, &str> = {
        let mut r = state.lock().unwrap();
        if !right_room {
            Err("This isn't the shared document's address.")
        } else if wants_host {
            if !local || r.host.is_some() {
                Err("Someone is already sharing here.")
            } else {
                r.host = Some(tx.clone());
                Ok("host".into())
            }
        } else if r.host.is_none() {
            Err("No one is sharing a document here.")
        } else {
            r.next += 1;
            let id = format!("p{}", r.next);
            r.guests.insert(id.clone(), tx.clone());
            if let Some(host) = &r.host {
                let _ = host.send(serde_json::json!({ "hub": "joined", "peer": id }).to_string());
            }
            Ok(id)
        }
    };
    let me = match me {
        Ok(me) => me,
        Err(message) => {
            let _ = sink.send(Message::Text(serde_json::json!({ "hub": "error", "message": message }).to_string().into())).await;
            return Ok(());
        }
    };
    let _ = tx.send(serde_json::json!({ "hub": "welcome", "you": me }).to_string());

    let writer = tauri::async_runtime::spawn(async move {
        while let Some(text) = rx.recv().await {
            if sink.send(Message::Text(text.into())).await.is_err() {
                break;
            }
        }
    });
    while let Some(Ok(message)) = source.next().await {
        let Message::Text(text) = message else { continue };
        let Ok(value) = serde_json::from_str::<serde_json::Value>(text.as_str()) else { continue };
        let data = value.get("data").cloned().unwrap_or_default();
        let to = value.get("to").and_then(|t| t.as_str()).unwrap_or("host").to_string();
        let r = state.lock().unwrap();
        let envelope = serde_json::json!({ "from": me, "data": data }).to_string();
        if me == "host" {
            for (id, guest) in &r.guests {
                if to == "all" || &to == id {
                    let _ = guest.send(envelope.clone());
                }
            }
        } else if let Some(host) = &r.host {
            let _ = host.send(envelope);
        }
    }
    // Gone: the others are told.
    {
        let mut r = state.lock().unwrap();
        if me == "host" {
            r.host = None;
            for guest in r.guests.values() {
                let _ = guest.send(serde_json::json!({ "hub": "left", "peer": "host" }).to_string());
            }
        } else {
            r.guests.remove(&me);
            if let Some(host) = &r.host {
                let _ = host.send(serde_json::json!({ "hub": "left", "peer": me }).to_string());
            }
        }
    }
    writer.abort();
    Ok(())
}

// ---- a shared folder ---------------------------------------------------------------------

/// Only files inside a ".duckwriter-live" folder may be used this way.
fn live_path(path: &str) -> Result<PathBuf, String> {
    let p = PathBuf::from(path);
    if !p.is_absolute() || !p.components().any(|c| c.as_os_str().to_string_lossy().ends_with(".duckwriter-live")) {
        return Err("Not a DuckWriter sharing folder.".into());
    }
    if p.components().any(|c| matches!(c, std::path::Component::ParentDir)) {
        return Err("Not a DuckWriter sharing folder.".into());
    }
    Ok(p)
}

/// Adds lines to one's own log in the sharing folder.
#[tauri::command]
pub fn folder_append(path: String, lines: String) -> Result<(), String> {
    let p = live_path(&path)?;
    if let Some(dir) = p.parent() {
        std::fs::create_dir_all(dir).map_err(|e| e.to_string())?;
    }
    let mut file = std::fs::OpenOptions::new().create(true).append(true).open(&p).map_err(|e| e.to_string())?;
    file.write_all(lines.as_bytes()).map_err(|e| e.to_string())?;
    file.flush().map_err(|e| e.to_string())
}

/// The whole lines added to a log since `offset`, and where to read from next time.
#[tauri::command]
pub fn folder_read(path: String, offset: u64) -> Result<(String, u64), String> {
    let p = live_path(&path)?;
    let Ok(mut file) = std::fs::File::open(&p) else { return Ok((String::new(), offset)) };
    let len = file.metadata().map_err(|e| e.to_string())?.len();
    if len <= offset {
        return Ok((String::new(), offset));
    }
    file.seek(SeekFrom::Start(offset)).map_err(|e| e.to_string())?;
    let mut bytes = Vec::new();
    file.read_to_end(&mut bytes).map_err(|e| e.to_string())?;
    // Only up to the last complete line: a sync program may still be writing the rest.
    let Some(end) = bytes.iter().rposition(|b| *b == b'\n') else { return Ok((String::new(), offset)) };
    bytes.truncate(end + 1);
    Ok((String::from_utf8_lossy(&bytes).into_owned(), offset + bytes.len() as u64))
}

/// The logs in a room folder.
#[tauri::command]
pub fn folder_list(dir: String) -> Result<Vec<String>, String> {
    let p = live_path(&dir)?;
    let Ok(entries) = std::fs::read_dir(Path::new(&p)) else { return Ok(Vec::new()) };
    Ok(entries.filter_map(|e| e.ok()).map(|e| e.file_name().to_string_lossy().into_owned()).filter(|n| n.ends_with(".log")).collect())
}
