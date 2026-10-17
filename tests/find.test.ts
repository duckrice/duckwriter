import { TextSelection, type EditorState } from "prosemirror-state";
import { beforeAll, describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import type { DocContext } from "../src/docx/model";
import { blankDocx } from "../src/docx/template";
import { createState } from "../src/editor/editor";
import { findMatches, findState, replaceAll, replaceCurrent, search, step } from "../src/editor/find";
import { schema } from "../src/editor/schema";

const any = { caseSensitive: false, wholeWord: false };
let context: DocContext;

beforeAll(async () => {
  context = (await readDocx(await blankDocx())).context;
});

const para = (...parts: (string | [string, "bold"])[]) =>
  schema.nodes.paragraph.create(
    null,
    parts.map((p) => (typeof p === "string" ? schema.text(p) : schema.text(p[0], [schema.marks.bold.create()]))),
  );

function stateOf(...paragraphs: ReturnType<typeof para>[]): EditorState {
  return createState(schema.nodes.doc.create(null, paragraphs), () => context);
}

const text = (state: EditorState, m: { from: number; to: number }) => state.doc.textBetween(m.from, m.to);

describe("find", () => {
  it("finds every match, ignoring case unless asked", () => {
    const state = stateOf(para("The cat sat. THE end."), para("Then the dog."));
    expect(findMatches(state.doc, "the", any).map((m) => text(state, m))).toEqual(["The", "THE", "The", "the"]);
    expect(findMatches(state.doc, "the", { ...any, caseSensitive: true })).toHaveLength(1);
  });

  it("finds whole words only, also with Norwegian letters", () => {
    const state = stateOf(para("The theme, the. Bær bæret bær."));
    expect(findMatches(state.doc, "the", { ...any, wholeWord: true })).toHaveLength(2);
    expect(findMatches(state.doc, "bær", { ...any, wholeWord: true })).toHaveLength(2);
  });

  it("finds text that runs across formatting, and special characters literally", () => {
    const state = stateOf(para("a bo", ["ld w", "bold"], "ord (x+1)"));
    expect(findMatches(state.doc, "bold word", any).map((m) => text(state, m))).toEqual(["bold word"]);
    expect(findMatches(state.doc, "(x+1)", any)).toHaveLength(1);
  });

  it("steps forwards and backwards, wrapping round", () => {
    let state = stateOf(para("one two one two one"));
    state = state.apply(search(state, "one", any));
    const at = () => state.selection.from;
    const go = (d: 1 | -1) => (state = state.apply(step(state, d)!));
    go(1);
    expect(at()).toBe(1);
    expect(findState(state)!.current).toBe(0);
    go(1);
    go(1);
    expect(findState(state)!.current).toBe(2);
    go(1);
    expect(findState(state)!.current).toBe(0);
    go(-1);
    expect(findState(state)!.current).toBe(2);
  });

  it("replaces the selected match keeping its formatting, then all the rest", () => {
    let state = stateOf(para(["colour", "bold"], " and colour, colour."));
    state = state.apply(search(state, "colour", any));
    // The first press only shows the match; the second replaces it.
    state = state.apply(replaceCurrent(state, "color")!);
    expect(state.doc.textContent).toBe("colour and colour, colour.");
    state = state.apply(replaceCurrent(state, "color")!);
    expect(state.doc.textContent).toBe("color and colour, colour.");
    expect(state.doc.firstChild!.firstChild!.marks.map((m) => m.type.name)).toEqual(["bold"]);
    expect(findState(state)!.matches).toHaveLength(2);
    const all = replaceAll(state, "color")!;
    expect(all.count).toBe(2);
    state = state.apply(all.tr);
    expect(state.doc.textContent).toBe("color and color, color.");
    expect(findState(state)!.matches).toHaveLength(0);
  });

  it("replace all is one undo step and highlights nothing once stopped", () => {
    let state = stateOf(para("x x x"));
    state = state.apply(search(state, "x", any));
    expect(findState(state)!.decorations.find()).toHaveLength(3);
    state = state.apply(replaceAll(state, "yy")!.tr);
    expect(state.doc.textContent).toBe("yy yy yy");
    state = state.apply(search(state, "", any));
    expect(findState(state)!.decorations.find()).toHaveLength(0);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 1)));
    expect(step(state, 1)).toBeNull();
  });
});
