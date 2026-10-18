// Matching keyboard shortcuts written like "Ctrl+Shift+L", on any keyboard
// layout. A key is recognised by the character it types or by where it sits
// (so Ctrl+] works on a Norwegian keyboard, where ] needs AltGr), and typing
// with AltGr (@, €, {, [ …) never sets off a Ctrl+Alt shortcut.

/** The character a physical key has on a US keyboard, which shortcuts are written with. */
const CODE_CHARS: Record<string, string> = {
  BracketLeft: "[",
  BracketRight: "]",
  Equal: "=",
  Minus: "-",
  Period: ".",
  Comma: ",",
  Slash: "/",
  Backslash: "\\",
  IntlBackslash: "<",
  Semicolon: ";",
  Quote: "'",
  Backquote: "`",
  Space: "Space",
  NumpadAdd: "+",
  NumpadSubtract: "-",
};

function codeChar(code: string): string | null {
  if (code.startsWith("Key")) return code.slice(3);
  if (code.startsWith("Digit")) return code.slice(5);
  return CODE_CHARS[code] ?? null;
}

export interface KeyEventLike {
  key: string;
  code: string;
  ctrlKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
  metaKey?: boolean;
}

interface Parsed {
  ctrl: boolean;
  shift: boolean;
  alt: boolean;
  key: string;
}

function parse(spec: string): Parsed {
  const parts = spec.split("+");
  // "Ctrl++" splits into ["Ctrl", "", ""]: the key is "+".
  let key = parts.pop()!;
  if (key === "" && parts[parts.length - 1] === "") {
    parts.pop();
    key = "+";
  }
  return { ctrl: parts.includes("Ctrl"), shift: parts.includes("Shift"), alt: parts.includes("Alt"), key };
}

const isLetterOrDigit = (key: string) => /^[A-Za-z0-9]$/.test(key);
const isNamed = (key: string) => key.length > 1;

/** Whether a key press is the shortcut \`spec\`, such as "Ctrl+B", "Ctrl+Alt+1", "Ctrl+>" or "F12". */
export function matches(spec: string, event: KeyEventLike): boolean {
  const want = parse(spec);
  const ctrl = event.ctrlKey || !!event.metaKey;
  if (ctrl !== want.ctrl || event.altKey !== want.alt) return false;
  const typed = event.key;
  // AltGr (Ctrl+Alt on Windows) typing a character: that is text, not a shortcut.
  const place = codeChar(event.code);
  if (event.ctrlKey && event.altKey && typed.length === 1 && typed.toUpperCase() !== (place === "Space" ? " " : (place ?? "")).toUpperCase()) return false;

  if (isNamed(want.key)) {
    // F12, Enter, Space, ArrowUp…
    const name = want.key === "Space" ? " " : want.key;
    return event.shiftKey === want.shift && (typed === name || codeChar(event.code) === want.key);
  }
  if (isLetterOrDigit(want.key)) {
    if (event.shiftKey !== want.shift) return false;
    return typed.toUpperCase() === want.key.toUpperCase() || codeChar(event.code) === want.key.toUpperCase();
  }
  // A sign such as > < + = [ ]: by the character typed, whatever it took to type it,
  if (typed === want.key) return !want.shift || event.shiftKey;
  // or by the key's place on a US keyboard.
  return codeChar(event.code) === want.key && event.shiftKey === want.shift;
}
