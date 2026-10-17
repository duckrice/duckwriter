// How messages get between the people writing together. A hub has one host
// (the person sharing) and guests; guests' messages go to the host, the host's
// to one guest or to all. Two kinds:
//
// - SocketHub: a WebSocket hub, either in the host's own DuckWriter (on the
//   local network) or a relay server on the internet (server/relay.mjs).
// - FolderHub: a shared folder (OneDrive, Dropbox, a network drive): everyone
//   appends to their own log file there and reads the others'.

import { invoke } from "@tauri-apps/api/core";

export interface HubEvents {
  data(from: string, data: string): void;
  joined(peer: string): void;
  left(peer: string): void;
  /** The connection is gone for good. */
  closed(reason: string): void;
}

export interface Hub {
  /** Connects; resolves to this participant's id ("host", or a guest's). */
  connect(events: HubEvents): Promise<string>;
  send(to: string, data: string): void;
  close(): void;
}

export class SocketHub implements Hub {
  private ws: WebSocket | null = null;
  private closing = false;

  /** @param server e.g. "ws://192.168.1.23:47800" or "wss://relay.example.com" */
  constructor(
    private server: string,
    private room: string,
    private role: "host" | "guest",
  ) {}

  connect(events: HubEvents): Promise<string> {
    return new Promise((resolve, reject) => {
      const url = `${this.server.replace(/\/+$/, "")}/r/${encodeURIComponent(this.room)}?role=${this.role}`;
      let ws: WebSocket;
      try {
        ws = new WebSocket(url);
      } catch {
        return reject(new Error("That isn't a valid address."));
      }
      this.ws = ws;
      let welcomed = false;
      const timer = setTimeout(() => {
        if (!welcomed) {
          ws.close();
          reject(new Error("No answer from there. Is the address right, and is the other computer sharing?"));
        }
      }, 10_000);
      ws.onmessage = (event) => {
        let message: { hub?: string; you?: string; peer?: string; message?: string; from?: string; data?: string };
        try {
          message = JSON.parse(String(event.data));
        } catch {
          return;
        }
        if (message.hub === "welcome") {
          welcomed = true;
          clearTimeout(timer);
          resolve(message.you!);
        } else if (message.hub === "error") {
          clearTimeout(timer);
          this.closing = true;
          reject(new Error(message.message ?? "Couldn't connect."));
        } else if (message.hub === "joined") events.joined(message.peer!);
        else if (message.hub === "left") events.left(message.peer!);
        else if (message.from !== undefined && typeof message.data === "string") events.data(message.from, message.data);
      };
      ws.onerror = () => {
        if (!welcomed) {
          clearTimeout(timer);
          reject(new Error("Couldn't connect. Is the address right, and is the other computer sharing?"));
        }
      };
      ws.onclose = () => {
        if (welcomed && !this.closing) events.closed("The connection was lost.");
      };
    });
  }

  send(to: string, data: string) {
    if (this.ws?.readyState === WebSocket.OPEN) this.ws.send(JSON.stringify({ to, data }));
  }

  close() {
    this.closing = true;
    this.ws?.close();
  }
}

/** How often the folder is looked at, and how long a quiet guest counts as gone. */
const POLL = 400;
const GONE_AFTER = 120_000;
const PING_EVERY = 20_000;

export class FolderHub implements Hub {
  private me = "";
  private timer: number | undefined;
  private pinger: number | undefined;
  private queue: string[] = [];
  private flushing = false;
  private offsets = new Map<string, number>();
  private lastSeen = new Map<string, number>();
  private events!: HubEvents;
  private closed = false;

  /** @param dir the room's folder: <shared folder>\.duckwriter-live\<room> */
  constructor(
    private dir: string,
    private role: "host" | "guest",
  ) {}

  private file = (id: string) => `${this.dir}\\${id}.log`;

  async connect(events: HubEvents): Promise<string> {
    this.events = events;
    this.me = this.role === "host" ? "host" : `g${Math.random().toString(36).slice(2, 10)}`;
    if (this.role === "guest") {
      const logs = await invoke<string[]>("folder_list", { dir: this.dir });
      if (!logs.includes("host.log")) throw new Error("No one is sharing a document in that folder. Is it the right folder, and has it finished syncing?");
      // Only what the host writes from now on matters, but the whole log is read: the welcome may be in it.
      this.offsets.set("host", 0);
    }
    await invoke("folder_append", { path: this.file(this.me), lines: `${JSON.stringify({ hub: "hello", at: Date.now() })}\n` });
    this.timer = window.setInterval(() => void this.poll(), POLL);
    this.pinger = window.setInterval(() => this.write({ hub: "ping", at: Date.now() }), PING_EVERY);
    return this.me;
  }

  private async poll() {
    if (this.closed) return;
    try {
      if (this.role === "host") {
        for (const name of await invoke<string[]>("folder_list", { dir: this.dir })) {
          const id = name.replace(/\.log$/, "");
          if (id === "host") continue;
          if (!this.offsets.has(id)) {
            this.offsets.set(id, 0);
            this.lastSeen.set(id, Date.now());
          }
          await this.read(id);
        }
        // A guest that has gone quiet (closed without saying goodbye) has left.
        for (const [id, seen] of this.lastSeen) {
          if (Date.now() - seen > GONE_AFTER) {
            this.lastSeen.delete(id);
            this.events.left(id);
          }
        }
      } else await this.read("host");
    } catch {
      // The folder may be briefly unavailable while syncing.
    }
  }

  private async read(id: string) {
    const [text, next] = await invoke<[string, number]>("folder_read", { path: this.file(id), offset: this.offsets.get(id) ?? 0 });
    this.offsets.set(id, next);
    if (!text) return;
    if (this.lastSeen.has(id)) this.lastSeen.set(id, Date.now());
    for (const line of text.split("\n")) {
      if (!line.trim()) continue;
      let m: { hub?: string; to?: string; data?: string };
      try {
        m = JSON.parse(line);
      } catch {
        continue;
      }
      if (this.role === "host") {
        if (m.hub === "hello") this.events.joined(id);
        else if (m.hub === "bye") {
          this.lastSeen.delete(id);
          this.events.left(id);
        } else if (typeof m.data === "string") this.events.data(id, m.data);
      } else if (m.hub === "bye") {
        this.closed = true;
        this.events.closed("The person sharing stopped.");
      } else if (typeof m.data === "string" && (m.to === "all" || m.to === this.me)) this.events.data("host", m.data);
    }
  }

  private write(message: object) {
    this.queue.push(JSON.stringify(message));
    void this.flush();
  }

  /** Writes what is waiting, a batch at a time, in order. */
  private async flush() {
    if (this.flushing) return;
    this.flushing = true;
    try {
      while (this.queue.length) {
        const lines = this.queue.splice(0).join("\n") + "\n";
        await invoke("folder_append", { path: this.file(this.me), lines });
      }
    } finally {
      this.flushing = false;
    }
  }

  send(to: string, data: string) {
    this.write(this.role === "host" ? { to, data } : { data });
  }

  close() {
    if (this.closed) return;
    this.write({ hub: "bye" });
    this.closed = true;
    clearInterval(this.timer);
    clearInterval(this.pinger);
  }
}
