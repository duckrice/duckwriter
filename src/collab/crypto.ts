// Messages between people writing together are encrypted (AES-GCM) with a key
// made from the invitation's secret. The hub, relay server or shared folder
// in between only ever sees the encrypted text.

const b64 = (bytes: Uint8Array) => {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
};
const unb64 = (text: string) => Uint8Array.from(atob(text), (c) => c.charCodeAt(0));

export const toBase64 = b64;
export const fromBase64 = unb64;

/** A new random secret, for an invitation. */
export function newSecret(bytes = 16): string {
  return b64(crypto.getRandomValues(new Uint8Array(bytes))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

/** The room's key, from its secret and name. */
export async function roomKey(secret: string, room: string): Promise<CryptoKey> {
  const material = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(`duckwriter:${room}:${secret}`));
  return crypto.subtle.importKey("raw", material, "AES-GCM", false, ["encrypt", "decrypt"]);
}

export async function seal(key: CryptoKey, text: string): Promise<string> {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const sealed = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, new TextEncoder().encode(text)));
  const out = new Uint8Array(iv.length + sealed.length);
  out.set(iv);
  out.set(sealed, iv.length);
  return b64(out);
}

/** Throws if the message wasn't sealed with this key (a wrong code, or tampering). */
export async function open(key: CryptoKey, sealed: string): Promise<string> {
  const bytes = unb64(sealed);
  const plain = await crypto.subtle.decrypt({ name: "AES-GCM", iv: bytes.subarray(0, 12) }, key, bytes.subarray(12));
  return new TextDecoder().decode(plain);
}
