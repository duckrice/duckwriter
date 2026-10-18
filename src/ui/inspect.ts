// Check before sending (Alt+Shift+I): what is easy to forget before a
// document goes to a journal, a supervisor or a customer. Each finding has a
// Show button that goes from one place to the next.

import type { Node } from "prosemirror-model";
import { TextSelection, NodeSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { h } from "../dom";
import { parseCitation } from "../editor/citations";
import { bookmarkOf, fieldWords } from "../editor/fields";
import { changes, commentAnchors } from "../editor/review";
import { fieldKind, schema } from "../editor/schema";
import { mistakes } from "../editor/spelling";
import { headingLevel } from "./references";

export interface Finding {
  /** "warning" needs a look; "info" is just to know. */
  kind: "warning" | "info";
  text: string;
  /** Where (positions in the text), to go through with Show. */
  places: number[];
  /** Pictures are selected whole; text gets the cursor. */
  select?: "node";
}

const plural = (n: number, one: string, many: string) => `${n} ${n === 1 ? one : many}`;

/** Everything worth a look before the document goes out. */
export function inspect(doc: Node, ctx: DocContext, spellingMistakes: number[] = []): Finding[] {
  const out: Finding[] = [];
  const open = changes(doc);
  if (open.length) out.push({ kind: "warning", text: `${plural(open.length, "tracked change is", "tracked changes are")} still there to accept or reject.`, places: open.map((c) => c.from) });

  const comments = commentAnchors(doc).filter((a) => {
    const item = ctx.comments.items.get(a.id);
    return item && !item.parent && !item.done;
  });
  if (comments.length) out.push({ kind: "warning", text: `${plural(comments.length, "comment is", "comments are")} not resolved (they go with the document).`, places: comments.map((a) => a.from + 1) });

  const undescribed: number[] = [];
  const bookmarks = new Set<string>();
  const refs: { name: string; pos: number }[] = [];
  const missingSources: number[] = [];
  const emptyHeadings: number[] = [];
  const known = new Set(ctx.bibliography.sources.map((s) => s.tag));
  doc.descendants((node, pos) => {
    if (node.type === schema.nodes.image && !node.attrs.alt?.trim()) undescribed.push(pos);
    const mark = bookmarkOf(node);
    if (mark?.kind === "start" && mark.name) bookmarks.add(mark.name);
    if (node.type === schema.nodes.field) {
      const kind = fieldKind(node.attrs.instr);
      if (kind === "REF") refs.push({ name: fieldWords(node.attrs.instr)[1] ?? "", pos });
      if (kind === "CITATION" && parseCitation(node.attrs.instr).tags.some((t) => !known.has(t))) missingSources.push(pos);
    }
    const link = schema.marks.link.isInSet(node.marks);
    if (node.isText && link?.attrs.anchor && !link.attrs.href) refs.push({ name: link.attrs.anchor, pos });
    if (node.type === schema.nodes.paragraph && headingLevel(ctx, node) && !node.textContent.trim()) emptyHeadings.push(pos + 1);
    return true;
  });
  const broken = refs.filter((r) => r.name && !bookmarks.has(r.name) && !r.name.startsWith("_Toc"));
  if (undescribed.length) out.push({ kind: "warning", text: `${plural(undescribed.length, "picture has", "pictures have")} no description for people who can't see it (select a picture → Description).`, places: undescribed, select: "node" });
  if (missingSources.length) out.push({ kind: "warning", text: `${plural(missingSources.length, "citation refers", "citations refer")} to a source that isn't in the list of sources.`, places: missingSources });
  if (broken.length) out.push({ kind: "warning", text: `${plural(broken.length, "reference or link goes", "references or links go")} to a place that isn't in the document any more.`, places: broken.map((r) => r.pos) });
  if (emptyHeadings.length) out.push({ kind: "warning", text: `${plural(emptyHeadings.length, "heading is", "headings are")} empty.`, places: emptyHeadings });
  if (spellingMistakes.length) out.push({ kind: "info", text: `${plural(spellingMistakes.length, "word is", "words are")} marked as misspelt.`, places: spellingMistakes });

  // Names the file carries: who commented and who made changes.
  const names = new Set<string>();
  for (const c of ctx.comments.items.values()) if (c.author) names.add(c.author);
  for (const c of open) if (c.author) names.add(c.author);
  if (names.size) out.push({ kind: "info", text: `Names in the document (comments and changes): ${[...names].sort().join(", ")}.`, places: [] });
  return out;
}

/** Shows the findings, with buttons to go through the places. */
export function showInspection(view: EditorView, ctx: DocContext) {
  const spelling = mistakes(view.state).map((m) => m.from);
  const findings = inspect(view.state.doc, ctx, spelling);
  const go = (f: Finding, i: number) => {
    const pos = f.places[i % f.places.length];
    const doc = view.state.doc;
    const sel = f.select === "node" && doc.nodeAt(pos) ? NodeSelection.create(doc, pos) : TextSelection.near(doc.resolve(Math.min(pos + (f.select ? 0 : 0), doc.content.size)));
    view.dispatch(view.state.tr.setSelection(sel).scrollIntoView());
    (view.domAtPos(pos).node as HTMLElement)?.parentElement?.scrollIntoView?.({ block: "center" });
  };
  const rows = findings.map((f) => {
    let next = 0;
    const show = h("button", { type: "button", class: "btn small" }, f.places.length > 1 ? `Show (1 of ${f.places.length})` : "Show");
    show.addEventListener("click", () => {
      go(f, next);
      next = (next + 1) % f.places.length;
      show.textContent = f.places.length > 1 ? `Show (${next + 1} of ${f.places.length})` : "Show";
    });
    return h("div", { class: `inspect-item ${f.kind}` }, h("span", { class: "inspect-mark" }, f.kind === "warning" ? "!" : "i"), h("span", { class: "source-label inspect-text" }, f.text), f.places.length ? show : null);
  });
  const dialog = h(
    "dialog",
    { class: "dialog wide inspect-dialog" },
    h("h2", {}, "Check before sending"),
    findings.length ? h("div", { class: "source-list inspect-list" }, ...rows) : h("p", {}, "Nothing to look at: no open changes or comments, every picture is described, and every reference has its target."),
    h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn primary", onclick: () => dialog.close() }, "Close")),
  ) as HTMLDialogElement;
  // Not modal: the document can be looked at and fixed while it is open.
  dialog.addEventListener("close", () => {
    dialog.remove();
    view.focus();
  });
  document.body.append(dialog);
  dialog.show();
}
