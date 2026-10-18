import type { Node } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { compareDocs, countChanges } from "../src/editor/compare";
import { changes } from "../src/editor/review";
import { schema } from "../src/editor/schema";

const doc = (...texts: string[]) => schema.nodes.doc.create(null, texts.map((t) => schema.nodes.paragraph.create(null, t ? schema.text(t) : null)));
let n = 0;
const options = { author: "Compare", newId: () => String(++n) };

/** The text with changes shown as [+added+] and [-removed-]. */
function show(d: Node): string[] {
  const out: string[] = [];
  d.forEach((p) => {
    let s = p.attrs.change ? `${p.attrs.change.kind}¶ ` : "";
    p.forEach((c) => {
      const t = c.text ?? "";
      if (schema.marks.insertion.isInSet(c.marks)) s += `[+${t}+]`;
      else if (schema.marks.deletion.isInSet(c.marks)) s += `[-${t}-]`;
      else s += t;
    });
    out.push(s);
  });
  return out;
}

describe("comparing versions", () => {
  it("finds changed words inside a paragraph", () => {
    const result = compareDocs(doc("The rate was high in May."), doc("The rate is high in June."), options);
    expect(show(result)).toEqual(["The rate [-was-][+is+] high in [-May-][+June+]."]);
    expect(countChanges(result)).toBe(4);
  });

  it("finds paragraphs added, removed and kept", () => {
    const result = compareDocs(doc("Intro.", "Old method, now gone.", "Results are clear.", "End."), doc("Intro.", "Results are clear.", "A brand new paragraph here.", "End."), options);
    expect(show(result)).toEqual(["Intro.", "del¶ [-Old method, now gone.-]", "Results are clear.", "ins¶ [+A brand new paragraph here.+]", "End."]);
  });

  it("keeps the formatting of the newer version, and its changes can be accepted or rejected", () => {
    const bold = schema.marks.bold.create();
    const older = doc("Mass and speed.");
    const newer = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, [schema.text("Mass", [bold]), schema.text(" and velocity.")]));
    const result = compareDocs(older, newer, options);
    expect(show(result)).toEqual(["Mass and [-speed-][+velocity+]."]);
    expect(result.firstChild!.firstChild!.marks.map((m) => m.type.name)).toEqual(["bold"]);
    expect(changes(result).map((c) => c.kind)).toEqual(["del", "ins"]);
  });

  it("finds nothing when nothing changed", () => {
    const result = compareDocs(doc("Same.", "Text."), doc("Same.", "Text."), options);
    expect(countChanges(result)).toBe(0);
  });
});
