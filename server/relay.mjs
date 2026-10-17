// DuckWriter's relay: lets people write together over the internet without
// being on the same network. It passes messages between the person sharing a
// document (the host) and those who join, and nothing else: the messages are
// encrypted in DuckWriter with a key that never reaches this server, so it
// can't read the document. It keeps nothing on disk.
//
//   node relay.mjs                 (port 8787, or the PORT environment variable)
//
// Connect to  ws(s)://<server>/r/<room>?role=host   (the person sharing)
//         or  ws(s)://<server>/r/<room>?role=guest  (everyone else)
// The same rules as the hub in DuckWriter itself (src-tauri/src/collab.rs).

import { createServer } from "node:http";
import { WebSocketServer } from "ws";

const PORT = Number(process.env.PORT ?? 8787);
/** The largest message: a document with its pictures, encrypted. */
const MAX_MESSAGE = 64 * 1024 * 1024;
/** Rooms with no one in them for this long are forgotten. */
const IDLE = 60 * 60 * 1000;

/** @type {Map<string, { host: import("ws").WebSocket | null, guests: Map<string, import("ws").WebSocket>, next: number, seen: number }>} */
const rooms = new Map();

const http = createServer((req, res) => {
  res.writeHead(200, { "content-type": "text/plain; charset=utf-8" });
  res.end(`DuckWriter relay. ${rooms.size} ${rooms.size === 1 ? "room" : "rooms"} open.\n`);
});
const wss = new WebSocketServer({ server: http, maxPayload: MAX_MESSAGE });

const send = (ws, message) => {
  if (ws && ws.readyState === ws.OPEN) ws.send(JSON.stringify(message));
};

wss.on("connection", (ws, req) => {
  const url = new URL(req.url ?? "/", "http://relay");
  const match = /^\/r\/([A-Za-z0-9_-]{6,64})$/.exec(url.pathname);
  if (!match) {
    send(ws, { hub: "error", message: "This isn't a DuckWriter sharing address." });
    return ws.close();
  }
  const id = match[1];
  const role = url.searchParams.get("role") === "host" ? "host" : "guest";
  let room = rooms.get(id);
  let me;
  if (role === "host") {
    if (room?.host) {
      send(ws, { hub: "error", message: "Someone is already sharing here." });
      return ws.close();
    }
    room ??= { host: null, guests: new Map(), next: 0, seen: Date.now() };
    room.host = ws;
    rooms.set(id, room);
    me = "host";
  } else {
    if (!room?.host) {
      send(ws, { hub: "error", message: "No one is sharing a document here (any more)." });
      return ws.close();
    }
    me = `p${++room.next}`;
    room.guests.set(me, ws);
    send(room.host, { hub: "joined", peer: me });
  }
  send(ws, { hub: "welcome", you: me });

  ws.on("message", (raw) => {
    let message;
    try {
      message = JSON.parse(raw.toString());
    } catch {
      return;
    }
    room.seen = Date.now();
    const envelope = { from: me, data: message.data };
    if (me === "host") {
      for (const [gid, guest] of room.guests) if (message.to === "all" || message.to === gid) send(guest, envelope);
    } else send(room.host, envelope);
  });

  ws.on("close", () => {
    if (me === "host") {
      room.host = null;
      for (const guest of room.guests.values()) send(guest, { hub: "left", peer: "host" });
    } else {
      room.guests.delete(me);
      send(room.host, { hub: "left", peer: me });
    }
    if (!room.host && !room.guests.size) rooms.delete(id);
  });
});

// Keeps connections alive through proxies, and forgets forgotten rooms.
setInterval(() => {
  for (const ws of wss.clients) if (ws.readyState === ws.OPEN) ws.ping();
  for (const [id, room] of rooms) if (!room.host && Date.now() - room.seen > IDLE) rooms.delete(id);
}, 30_000).unref();

http.listen(PORT, () => console.log(`DuckWriter relay on port ${PORT}`));
