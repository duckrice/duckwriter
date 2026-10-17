// Write together (Alt+Shift+W): share the open document, or join one someone
// shares, in one of three ways:
//
//   - This network: the sharer's DuckWriter is the meeting point; nothing goes
//     over the internet.
//   - A relay server: for the internet. A small program anyone can run
//     (server/relay.mjs); everything passing it is encrypted.
//   - A shared folder: OneDrive, Dropbox, Google Drive or a network drive. No
//     server at all; changes arrive as fast as the folder syncs.

import { invoke } from "@tauri-apps/api/core";
import { open as pickFolderDialog } from "@tauri-apps/plugin-dialog";
import { authorName } from "../author";
import { newSecret } from "../collab/crypto";
import { FolderHub, SocketHub, type Hub } from "../collab/hub";
import { decodeInvite, encodeInvite, type Method } from "../collab/invite";
import { CollabSession, type CollabApp } from "../collab/session";
import { h } from "../dom";
import { isDesktop } from "../platform";
import { askName } from "./comments";

export interface ShareHost extends CollabApp {
  collab: CollabSession | null;
  /** The open document's folder, if it has been saved. */
  readonly folder: string | null;
}

const RELAY_KEY = "duckwriter.relay";
const LAN_PORT = 47800;

function remembered(key: string): string {
  try {
    return localStorage.getItem(key) ?? "";
  } catch {
    return "";
  }
}

/** Starts sharing the open document; resolves to the invitation to give the others. */
export async function startSharing(app: ShareHost, method: Method, place: string): Promise<string> {
  const name = await askName("Your name is shown to the others, next to your cursor.");
  if (!name) throw new Error("A name is needed, so the others know who is writing.");
  const room = newSecret(9);
  const secret = newSecret(16);
  let hub: Hub;
  let address: string;
  if (method === "lan") {
    if (!isDesktop) throw new Error("Sharing on the network needs the DuckWriter app.");
    const info = await invoke<{ port: number; address: string }>("lan_host_start", { room, port: LAN_PORT });
    hub = new SocketHub(`ws://127.0.0.1:${info.port}`, room, "host");
    address = `ws://${info.address}:${info.port}`;
  } else if (method === "relay") {
    address = normalizeServer(place);
    if (!address) throw new Error("Give the relay server's address, like wss://relay.example.com.");
    localStorage.setItem(RELAY_KEY, address);
    hub = new SocketHub(address, room, "host");
  } else {
    if (!place) throw new Error("Choose the shared folder (in OneDrive, Dropbox, a network drive…).");
    hub = new FolderHub(`${place.replace(/[\\/]+$/, "")}\\.duckwriter-live\\${room}`, "host");
    address = place.split(/[\\/]/).filter(Boolean).pop() ?? "";
  }
  const invite = encodeInvite({ method, address, room, secret });
  try {
    app.collab = await CollabSession.host(app, hub, room, secret, invite, name);
  } catch (error) {
    if (method === "lan") void invoke("lan_host_stop");
    throw error;
  }
  return invite;
}

/** Joins a shared document from an invitation (and, for a folder, where that folder is here). */
export async function joinSharing(app: ShareHost, inviteText: string, folder: string): Promise<void> {
  const invite = decodeInvite(inviteText);
  const name = await askName("Your name is shown to the others, next to your cursor.");
  if (!name) throw new Error("A name is needed, so the others know who is writing.");
  let hub: Hub;
  if (invite.method === "folder") {
    if (!folder) throw new Error(`Choose the shared folder${invite.address ? ` ("${invite.address}")` : ""} as it is on this computer.`);
    hub = new FolderHub(`${folder.replace(/[\\/]+$/, "")}\\.duckwriter-live\\${invite.room}`, "guest");
  } else hub = new SocketHub(invite.address, invite.room, "guest");
  app.collab = await CollabSession.join(app, hub, invite.room, invite.secret, inviteText.trim(), name);
}

/** Stops sharing, or leaves. */
export function stopSharing(app: ShareHost) {
  const session = app.collab;
  if (!session) return;
  session.stop();
  if (session.role === "host" && decodeInviteSafe(session.invite)?.method === "lan") void invoke("lan_host_stop");
  app.collab = null;
  app.sharingChanged();
}

function decodeInviteSafe(text: string) {
  try {
    return decodeInvite(text);
  } catch {
    return null;
  }
}

function normalizeServer(text: string): string {
  const t = text.trim();
  if (!t) return "";
  if (/^wss?:\/\//i.test(t)) return t;
  if (/^https?:\/\//i.test(t)) return t.replace(/^http/i, "ws");
  return `wss://${t}`;
}

async function chooseFolder(input: HTMLInputElement) {
  const picked = await pickFolderDialog({ directory: true, defaultPath: input.value || undefined, title: "The shared folder" });
  if (typeof picked === "string") input.value = picked;
}

// ---- the dialog --------------------------------------------------------------------------

/** Shows sharing: who is here and the invitation, or how to share or join. */
export function showSharing(app: ShareHost) {
  const session = app.collab?.active ? app.collab : null;
  const dialog = h("dialog", { class: "dialog wide share-dialog" }) as HTMLDialogElement;
  dialog.addEventListener("close", () => {
    dialog.remove();
    app.body.focus();
  });
  document.body.append(dialog);
  if (session) sessionPanel(app, session, dialog);
  else choosePanel(app, dialog);
  dialog.showModal();
}

function inviteBox(invite: string): HTMLElement {
  const code = h("textarea", { class: "tb-input dialog-input wide invite-code", rows: 3, readonly: true, "aria-label": "Invitation" }) as HTMLTextAreaElement;
  code.value = invite;
  const copy = h("button", { type: "button", class: "btn" }, "Copy");
  copy.addEventListener("click", async () => {
    try {
      await navigator.clipboard.writeText(invite);
      copy.textContent = "Copied";
    } catch {
      code.select();
    }
  });
  return h("div", { class: "invite" }, code, copy);
}

function howToJoin(method: Method): string {
  if (method === "lan") return "Send this invitation to the others (chat, e-mail). On the same network, they choose Write together → Join and paste it. Your computer may ask once whether DuckWriter may use the network: allow it on private networks.";
  if (method === "relay") return "Send this invitation to the others. Wherever they are, they choose Write together → Join and paste it.";
  return "Send this invitation to the others. They need the same shared folder on their computer; they choose Write together → Join, paste it and pick the folder.";
}

function sessionPanel(app: ShareHost, session: CollabSession, dialog: HTMLDialogElement) {
  const method = decodeInviteSafe(session.invite)?.method ?? "lan";
  const people = h(
    "ul",
    { class: "people" },
    ...[...session.people.values()].map((p) => {
      const swatch = h("span", { class: "person-color" });
      swatch.style.background = p.color;
      return h("li", {}, swatch, `${p.name || "Someone"}${p.id === session.me ? " (you)" : p.id === "host" ? " (sharing)" : ""}`);
    }),
  );
  const stop = h("button", { type: "button", class: "btn" }, session.role === "host" ? "Stop sharing" : "Leave");
  stop.addEventListener("click", () => {
    stopSharing(app);
    dialog.close();
  });
  dialog.append(
    h("h2", {}, session.role === "host" ? "Sharing this document" : "Writing together"),
    h("p", { class: "dialog-hint" }, `${session.people.size} ${session.people.size === 1 ? "person" : "people"} here:`),
    people,
    ...(session.role === "host" ? [h("p", { class: "dialog-hint" }, howToJoin(method)), inviteBox(session.invite)] : [h("p", { class: "dialog-hint" }, "The person sharing keeps the document; save a copy (F12) to keep one yourself.")]),
    h("div", { class: "dialog-footer" }, stop, h("span", { class: "dialog-spacer" }), h("button", { type: "button", class: "btn primary", onclick: () => dialog.close() }, "Close")),
  );
}

function choosePanel(app: ShareHost, dialog: HTMLDialogElement) {
  const error = h("p", { class: "dialog-error", role: "alert" });
  const option = (value: Method, title: string, hint: string, checked = false) => {
    const radio = h("input", { type: "radio", name: "share-method", value, checked }) as HTMLInputElement;
    return h("label", { class: "share-option" }, radio, h("span", {}, h("strong", {}, title), h("span", { class: "dialog-hint" }, hint)));
  };
  const relay = h("input", { class: "tb-input dialog-input wide", placeholder: "wss://relay.example.com", value: remembered(RELAY_KEY), "aria-label": "Relay server" }) as HTMLInputElement;
  const shareFolder = h("input", { class: "tb-input dialog-input wide", placeholder: "C:\\Users\\…\\OneDrive\\Shared", value: app.folder ?? "", "aria-label": "Shared folder" }) as HTMLInputElement;
  const pickShare = h("button", { type: "button", class: "btn" }, "Choose…");
  pickShare.addEventListener("click", () => void chooseFolder(shareFolder));
  const result = h("div", { class: "share-result" });
  const start = h("button", { type: "button", class: "btn primary" }, "Start sharing");
  start.addEventListener("click", async () => {
    const method = (dialog.querySelector('input[name="share-method"]:checked') as HTMLInputElement).value as Method;
    error.textContent = "";
    start.disabled = true;
    try {
      const invite = await startSharing(app, method, method === "relay" ? relay.value : shareFolder.value);
      result.replaceChildren(h("p", { class: "dialog-hint" }, howToJoin(method)), inviteBox(invite));
      start.hidden = true;
    } catch (e) {
      error.textContent = e instanceof Error ? e.message : String(e);
      start.disabled = false;
    }
  });

  const inviteInput = h("textarea", { class: "tb-input dialog-input wide", rows: 3, placeholder: "Paste the invitation (DW1.…)", "aria-label": "Invitation" }) as HTMLTextAreaElement;
  const joinFolder = h("input", { class: "tb-input dialog-input wide", placeholder: "The shared folder, as it is on this computer", "aria-label": "Shared folder" }) as HTMLInputElement;
  const pickJoin = h("button", { type: "button", class: "btn" }, "Choose…");
  pickJoin.addEventListener("click", () => void chooseFolder(joinFolder));
  const folderRow = h("div", { class: "source-actions", hidden: true }, joinFolder, pickJoin);
  inviteInput.addEventListener("input", () => {
    folderRow.hidden = decodeInviteSafe(inviteInput.value)?.method !== "folder";
  });
  const join = h("button", { type: "button", class: "btn primary" }, "Join");
  join.addEventListener("click", async () => {
    error.textContent = "";
    join.disabled = true;
    join.textContent = "Joining…";
    try {
      await joinSharing(app, inviteInput.value, joinFolder.value);
      dialog.close();
    } catch (e) {
      error.textContent = e instanceof Error ? e.message : String(e);
      join.disabled = false;
      join.textContent = "Join";
    }
  });

  dialog.append(
    h("h2", {}, "Write together"),
    h("h3", {}, "Share this document"),
    option("lan", "On this network", "Wi-Fi at home, the office or the university. Your DuckWriter is the meeting point; nothing goes over the internet.", true),
    option("relay", "Over the internet, through a relay server", "A small program you (or someone you trust) run on a server. It only passes encrypted messages along; it can't read the document."),
    h("div", { class: "share-place" }, relay),
    option("folder", "Through a shared folder", "OneDrive, Dropbox, Google Drive or a network drive that everyone has. No server; changes arrive as fast as the folder syncs (usually seconds)."),
    h("div", { class: "share-place source-actions" }, shareFolder, pickShare),
    h("div", { class: "dialog-footer" }, start),
    result,
    h("h3", {}, "Join someone"),
    inviteInput,
    folderRow,
    h("div", { class: "dialog-footer" }, join),
    error,
    h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Close")),
  );
  void authorName;
}
