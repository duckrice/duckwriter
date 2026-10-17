// An invitation to write together: how to connect, which room, and the secret
// the messages are encrypted with. Shared as one line of text ("DW1.…"), by
// e-mail, chat or read out.

export type Method = "lan" | "relay" | "folder";

export interface Invite {
  method: Method;
  /** The hub's address (lan, relay); for a folder, its name as a hint. */
  address: string;
  room: string;
  secret: string;
}

const PREFIX = "DW1.";

export function encodeInvite(invite: Invite): string {
  const json = JSON.stringify({ m: invite.method, a: invite.address, r: invite.room, k: invite.secret });
  const b64 = btoa(String.fromCharCode(...new TextEncoder().encode(json)));
  return PREFIX + b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

export function decodeInvite(text: string): Invite {
  const code = text.trim().replace(/\s+/g, "");
  const at = code.indexOf(PREFIX);
  if (at < 0) throw new Error("That isn't a DuckWriter invitation (it starts with DW1.).");
  try {
    const b64 = code.slice(at + PREFIX.length).replace(/-/g, "+").replace(/_/g, "/");
    const json = JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))));
    if (!["lan", "relay", "folder"].includes(json.m) || !json.r || !json.k) throw new Error();
    return { method: json.m, address: json.a ?? "", room: json.r, secret: json.k };
  } catch {
    throw new Error("The invitation is incomplete. Copy all of it, from DW1. to the end.");
  }
}
