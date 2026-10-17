import { TextSelection, type EditorState, type Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { afterEach, beforeAll, describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import type { DocContext } from "../src/docx/model";
import { blankDocx } from "../src/docx/template";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";
import { addToDictionary, correct, ignoreWord, mistakes, nextMistake, setSpellChecker, spellingPlugin, wordsOf } from "../src/editor/spelling";

let context: DocContext;
// Knows a little English and a little Norwegian.
const WORDS = new Set("the weather is fine today denne setningen er riktig and".split(" "));
const fake = { check: async (words: string[]) => words.map((w) => WORDS.has(w.toLowerCase())), suggest: async () => ["the"] };

beforeAll(async () => {
  context = (await readDocx(await blankDocx())).context;
});
afterEach(() => setSpellChecker(null));

const para = (text: string) => schema.nodes.paragraph.create(null, text ? schema.text(text) : null);

/** A pretend editor that the spelling plugin can update. */
function editor(...texts: string[]) {
  const view = {
    state: createState(schema.nodes.doc.create(null, texts.map(para)), () => context),
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
  };
  // The plugin's view hook, so rechecks reach this editor.
  const plugin = view.state.plugins.find((p) => p.spec.key === spellingPlugin().spec.key)!;
  plugin.spec.view!(view as unknown as EditorView);
  return view as { state: EditorState; dispatch(tr: Transaction): void };
}

const settle = () => new Promise((r) => setTimeout(r, 300));
const marked = (state: EditorState) => mistakes(state).map((m) => m.word);

describe("spelling", () => {
  it("picks out the words, leaving numbers, abbreviations and addresses alone", () => {
    const doc = schema.nodes.doc.create(null, [para("Rock'n'roll at 5pm with DNA, see www.duck.no or me@duck.no: ‘quoted’ x")]);
    expect(wordsOf(doc.firstChild!, 0).map((w) => w.word)).toEqual(["Rock'n'roll", "at", "with", "see", "or", "quoted"]);
    const w = wordsOf(doc.firstChild!, 0)[1];
    expect(doc.textBetween(w.from, w.to)).toBe("at");
  });

  it("marks words no language knows, in Norwegian and English together", async () => {
    setSpellChecker(fake);
    const view = editor("The weather is fien today.", "Denne setningen er riktgi.");
    await settle();
    expect(marked(view.state)).toEqual(["fien", "riktgi"]);
  });

  it("doesn't mark the word being typed until the cursor leaves it", async () => {
    setSpellChecker(fake);
    const view = editor("The weather");
    await settle();
    const end = view.state.doc.content.size - 1;
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, end)));
    view.dispatch(view.state.tr.insertText(" wrnog"));
    await settle();
    expect(marked(view.state)).toEqual([]);
    view.dispatch(view.state.tr.insertText(" "));
    expect(marked(view.state)).toEqual(["wrnog"]);
  });

  it("finds the next mistake, corrects it in place, and can ignore or learn words", async () => {
    setSpellChecker(fake);
    const view = editor("Teh weather is fien and fien.");
    await settle();
    let next = nextMistake(view.state)!;
    expect(next.word).toBe("Teh");
    view.dispatch(correct(view.state, next, "The")!);
    expect(view.state.doc.textContent).toBe("The weather is fien and fien.");
    expect(marked(view.state)).toEqual(["fien", "fien"]);
    ignoreWord("fien");
    expect(marked(view.state)).toEqual([]);
    view.dispatch(view.state.tr.insertText(" Quackery ", view.state.doc.content.size - 1));
    await settle();
    next = nextMistake(view.state)!;
    expect(next.word).toBe("Quackery");
    addToDictionary("Quackery");
    expect(marked(view.state)).toEqual([]);
  });
});
