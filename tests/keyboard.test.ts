import { TextSelection } from "prosemirror-state";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { createState } from "../src/editor/editor";
import { moveByParagraph } from "../src/editor/commands";
import { schema } from "../src/editor/schema";
import { loadKeyLayout, localKey, localKeys } from "../src/keys";

describe("Ctrl+Up and Ctrl+Down", () => {
  it("go by paragraph, as in Word, and select with Shift", async () => {
    const { context } = await readDocx(await blankDocx());
    const p = (t: string) => schema.nodes.paragraph.create(null, schema.text(t));
    let state = createState(schema.nodes.doc.create(null, [p("One two"), p("Three four"), p("Five")]), () => context);
    // In the middle of "Three four".
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 14)));
    const go = (cmd: ReturnType<typeof moveByParagraph>) => {
      let next = state;
      cmd(state, (tr) => (next = state.apply(tr)));
      state = next;
      return state.selection;
    };
    expect(go(moveByParagraph(-1)).head).toBe(10); // start of "Three four"
    expect(go(moveByParagraph(-1)).head).toBe(1); // start of "One two"
    expect(go(moveByParagraph(1)).head).toBe(10);
    expect(go(moveByParagraph(1)).head).toBe(22); // start of "Five"
    expect(go(moveByParagraph(1)).head).toBe(26); // end of the last
    const sel = go(moveByParagraph(-1, true));
    expect([sel.anchor, sel.head]).toEqual([26, 22]);
  });
});

describe("shortcuts as on this keyboard", () => {
  it("are named by the key that types them on a Norwegian keyboard", async () => {
    // A Norwegian keyboard: the keys where [ ] and \ are on an English one show Å, ¨ and '.
    const map = new Map([["Equal", "\\"], ["BracketLeft", "å"], ["BracketRight", "¨"], ["Backslash", "'"], ["Minus", "+"], ["Slash", "-"], ["IntlBackslash", "<"], ["KeyB", "b"]]);
    Object.defineProperty(navigator, "keyboard", { value: { getLayoutMap: async () => map }, configurable: true });
    await loadKeyLayout();
    expect(localKey("Ctrl+[")).toBe("Ctrl+Å");
    expect(localKey("Ctrl+]")).toBe("Ctrl+¨");
    expect(localKey("Ctrl+<")).toBe("Ctrl+<");
    expect(localKey("Ctrl+-")).toBe("Ctrl+-");
    expect(localKey("Alt+=")).toBe("Alt+=");
    expect(localKeys("Equation (Alt+=)")).toBe("Equation (Alt+=)");
    expect(localKey("Ctrl+B")).toBe("Ctrl+B");
    expect(localKey("Ctrl+ArrowUp")).toBe("Ctrl+ArrowUp");
    expect(localKeys("Smaller text (Ctrl+< or Ctrl+[)")).toBe("Smaller text (Ctrl+< or Ctrl+Å)");
  });
});
