import { describe, expect, it } from "vitest";
import { matches, type KeyEventLike } from "../src/keys";
import { SHORTCUTS, shortcutFor } from "../src/ui/shortcuts";

/** A key press: modifiers as a string ("ctrl shift"), the character typed, and the physical key. */
const press = (mods: string, key: string, code: string): KeyEventLike => ({
  key,
  code,
  ctrlKey: mods.includes("ctrl"),
  shiftKey: mods.includes("shift"),
  altKey: mods.includes("alt"),
});

const which = (event: KeyEventLike) => shortcutFor(event as KeyboardEvent)?.label ?? null;

describe("keyboard shortcuts", () => {
  it("match letters whatever the case, and need exactly the modifiers given", () => {
    expect(matches("Ctrl+B", press("ctrl", "b", "KeyB"))).toBe(true);
    expect(matches("Ctrl+B", press("ctrl shift", "B", "KeyB"))).toBe(false);
    expect(matches("Ctrl+Shift+L", press("ctrl shift", "L", "KeyL"))).toBe(true);
    expect(matches("F12", press("", "F12", "F12"))).toBe(true);
    // Ctrl+Alt+Space (Read aloud) is a shortcut, not AltGr typing a space.
    expect(matches("Ctrl+Alt+Space", press("ctrl alt", " ", "Space"))).toBe(true);
    expect(matches("Ctrl+Shift+Space", press("ctrl shift", " ", "Space"))).toBe(true);
  });

  it("change the font size on a US keyboard", () => {
    expect(which(press("ctrl shift", ">", "Period"))).toBe("Font size up, to the next size");
    expect(which(press("ctrl shift", "<", "Comma"))).toBe("Font size down, to the next size");
    expect(which(press("ctrl", "]", "BracketRight"))).toBe("Font size up by one point");
    expect(which(press("ctrl", "[", "BracketLeft"))).toBe("Font size down by one point");
  });

  it("change the font size on a Norwegian keyboard, where [ and ] need AltGr", () => {
    // The key left of Z types < and, with Shift, >.
    expect(which(press("ctrl", "<", "IntlBackslash"))).toBe("Font size down, to the next size");
    expect(which(press("ctrl shift", ">", "IntlBackslash"))).toBe("Font size up, to the next size");
    // Ctrl with the keys where US keyboards have [ and ] (å and ¨ on Norwegian ones).
    expect(which(press("ctrl", "å", "BracketLeft"))).toBe("Font size down by one point");
    expect(which(press("ctrl", "¨", "BracketRight"))).toBe("Font size up by one point");
    // Superscript and subscript: + is its own key, = is Shift+0.
    expect(which(press("ctrl", "+", "Minus"))).toBe("Superscript (x²)");
    expect(which(press("ctrl shift", "=", "Digit0"))).toBe("Subscript (x₂)");
    // Numbering: Shift+7 types "/".
    expect(which(press("ctrl shift", "/", "Digit7"))).toBe("Numbering");
  });

  it("never take characters typed with AltGr", () => {
    // AltGr is Ctrl+Alt on Windows: @, €, [, ], { and } are text, not shortcuts.
    expect(which(press("ctrl alt", "@", "Digit2"))).toBeNull();
    expect(which(press("ctrl alt", "€", "KeyE"))).toBeNull();
    expect(which(press("ctrl alt", "]", "Digit9"))).toBeNull();
    expect(which(press("ctrl alt", "}", "Digit0"))).toBeNull();
    // But a real Ctrl+Alt+1 is Heading 1.
    expect(which(press("ctrl alt", "1", "Digit1"))).toBe("Heading 1");
  });

  it("give each key press to one action only", () => {
    const seen = new Map<string, string>();
    for (const s of SHORTCUTS) {
      for (const k of s.keys) {
        expect(seen.get(k), `${k} is both "${seen.get(k)}" and "${s.label}"`).toBeUndefined();
        seen.set(k, s.label);
      }
    }
  });

  it("cover every toolbar action", () => {
    const labels = SHORTCUTS.map((s) => s.label.toLowerCase()).join("\n");
    for (const action of ["new document", "open", "save", "undo", "redo", "bold", "italic", "underline", "strikethrough", "superscript", "subscript", "font size", "font", "text colour", "highlight", "clear formatting", "style", "centre", "justify", "bullets", "numbering", "indent", "line spacing", "page break", "equation", "symbol", "picture", "footnote", "caption", "cross-reference", "table of contents", "numbered headings", "header and footer", "zoom", "grid"]) {
      expect(labels, action).toContain(action);
    }
  });
});
