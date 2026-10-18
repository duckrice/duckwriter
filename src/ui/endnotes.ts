// Endnotes: notes collected at the end of the document, numbered i, ii, iii
// as in Word (Ctrl+Alt+D). The list at the end is drawn from the notes;
// clicking a mark or a note opens it to edit.

import type { Node } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import { DOMSerializer } from "prosemirror-model";
import type { EditorView, NodeView } from "prosemirror-view";
import { ensureStyles } from "../docx/builtins";
import type { DocContext } from "../docx/model";
import { h } from "../dom";
import { createState, createView } from "../editor/editor";
import { schema } from "../editor/schema";

/** The endnotes in the order their marks come in the text. */
export function endnoteOrder(doc: Node): string[] {
  const ids: string[] = [];
  doc.descendants((n) => {
    if (n.type === schema.nodes.endnote && !ids.includes(n.attrs.id)) ids.push(n.attrs.id);
  });
  return ids;
}

const roman = (n: number) => {
  const map: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let out = "";
  for (const [v, s] of map) while (n >= v) (out += s), (n -= v);
  return out;
};

/** The list of endnotes at the end of the document. */
export class EndnotesView implements NodeView {
  dom = h("div", { class: "endnotes", contenteditable: "false" });

  constructor(
    node: Node,
    private view: EditorView,
    private context: () => DocContext,
  ) {
    this.render();
    void node;
  }

  render() {
    const ctx = this.context();
    const order = endnoteOrder(this.view.state.doc);
    const serializer = DOMSerializer.fromSchema(schema);
    const items = order.map((id, i) => {
      const note = ctx.endnotes.notes.get(id);
      const body = h("div", { class: "endnote-text" });
      if (note) body.append(serializer.serializeFragment(note.content));
      const item = h("div", { class: "endnote-item", "data-id": id, title: "Click to edit the note" }, h("span", { class: "endnote-number" }, roman(i + 1)), body);
      item.addEventListener("mousedown", (event) => {
        event.preventDefault();
        editEndnote(this.view, ctx, id);
      });
      return item;
    });
    this.dom.replaceChildren(h("div", { class: "endnotes-rule" }), ...(items.length ? items : [h("div", { class: "toc-empty" }, "Endnotes (Ctrl+Alt+D) are listed here.")]));
  }

  update(node: Node) {
    if (node.type !== schema.nodes.endnotes) return false;
    this.render();
    return true;
  }

  stopEvent() {
    return true;
  }

  ignoreMutation() {
    return true;
  }
}

/** Redraws the list (after a note was edited). */
function refresh(view: EditorView) {
  const tr = view.state.tr;
  view.state.doc.forEach((node, pos) => {
    if (node.type === schema.nodes.endnotes) tr.setNodeMarkup(pos, undefined, { version: (node.attrs.version ?? 0) + 1 });
  });
  if (tr.docChanged) view.dispatch(tr.setMeta("addToHistory", false).setMeta("noTrack", true));
}

/** Adds an endnote at the cursor and opens it to write in (Ctrl+Alt+D). */
export function insertEndnote(view: EditorView, ctx: DocContext) {
  ensureStyles(ctx, "EndnoteText", "EndnoteReference");
  const used = [...ctx.endnotes.notes.keys()].map(Number).filter((n) => n > 0);
  const id = String(Math.max(0, ...used) + 1);
  ctx.endnotes.notes.set(id, schema.nodes.doc.create(null, schema.nodes.paragraph.create({ style: ctx.styles.byId.has("EndnoteText") ? "EndnoteText" : null })));
  const mark = schema.nodes.endnote.create({ id });
  const tr = view.state.tr.replaceSelectionWith(mark, false);
  let hasList = false;
  tr.doc.forEach((n) => {
    if (n.type === schema.nodes.endnotes) hasList = true;
  });
  if (!hasList) tr.insert(tr.doc.content.size, schema.nodes.endnotes.create());
  view.dispatch(tr.scrollIntoView());
  editEndnote(view, ctx, id);
}

/** Edits a note in a small window; Esc or Done puts it back. */
export function editEndnote(view: EditorView, ctx: DocContext, id: string) {
  const note = ctx.endnotes.notes.get(id) ?? schema.nodes.doc.create(null, schema.nodes.paragraph.create());
  const number = roman(endnoteOrder(view.state.doc).indexOf(id) + 1 || 1);
  const host = h("div", { class: "endnote-editor" });
  const dialog = h(
    "dialog",
    { class: "dialog wide" },
    h("h2", {}, `Endnote ${number}`),
    host,
    h(
      "div",
      { class: "dialog-footer" },
      h(
        "button",
        {
          type: "button",
          class: "btn",
          onclick: () => {
            // Removing the note removes its mark too.
            const tr = view.state.tr;
            const marks: number[] = [];
            view.state.doc.descendants((n, pos) => {
              if (n.type === schema.nodes.endnote && n.attrs.id === id) marks.push(pos);
            });
            for (const pos of marks.reverse()) tr.delete(pos, pos + 1);
            ctx.endnotes.notes.delete(id);
            view.dispatch(tr);
            dialog.close();
          },
        },
        "Delete the note",
      ),
      h("span", { class: "dialog-spacer" }),
      h("button", { type: "button", class: "btn primary", onclick: () => dialog.close() }, "Done"),
    ),
  ) as HTMLDialogElement;
  document.body.append(dialog);
  const editor = createView(host, createState(note, () => ctx), { onChange: () => undefined }, "hf-doc en-doc");
  // The editor keeps Esc from the dialog, so it closes it itself.
  editor.dom.addEventListener("keydown", (event) => {
    if ((event as KeyboardEvent).key === "Escape") {
      event.preventDefault();
      dialog.close();
    }
  });
  dialog.addEventListener("close", () => {
    if (ctx.endnotes.notes.has(id)) ctx.endnotes.notes.set(id, editor.state.doc);
    editor.destroy();
    dialog.remove();
    refresh(view);
    view.focus();
  });
  dialog.showModal();
  editor.focus();
  editor.dispatch(editor.state.tr.setSelection(TextSelection.atEnd(editor.state.doc)));
}

/**
 * Turns every footnote into an endnote, or every endnote into a footnote, as
 * journals often ask. The notes keep their text; their numbers follow from
 * their new places. Undo puts them back (the old notes are kept until saving,
 * which writes only the ones the text refers to).
 */
export function convertNotes(view: EditorView, ctx: DocContext, to: "endnote" | "footnote"): number {
  const fromType = to === "endnote" ? schema.nodes.footnote : schema.nodes.endnote;
  const toType = to === "endnote" ? schema.nodes.endnote : schema.nodes.footnote;
  const from = to === "endnote" ? ctx.footnotes : ctx.endnotes;
  const into = to === "endnote" ? ctx.endnotes : ctx.footnotes;
  const [oldText, newText, oldRef, newRef] = to === "endnote"
    ? ["FootnoteText", "EndnoteText", "FootnoteReference", "EndnoteReference"]
    : ["EndnoteText", "FootnoteText", "EndnoteReference", "FootnoteReference"];
  ensureStyles(ctx, newText, newRef);
  const marks: { pos: number; node: Node }[] = [];
  view.state.doc.descendants((node, pos) => {
    if (node.type === fromType) marks.push({ pos, node });
    return !node.isInline;
  });
  if (!marks.length) return 0;
  let next = Math.max(0, ...[...into.notes.keys()].map(Number).filter(Number.isFinite)) + 1;
  const newIds = new Map<string, string>();
  const restyle = (doc: Node): Node => {
    const content: Node[] = [];
    doc.forEach((block) => {
      content.push(block.type === schema.nodes.paragraph && block.attrs.style === oldText ? block.type.create({ ...block.attrs, style: newText }, block.content, block.marks) : block);
    });
    return doc.type.create(doc.attrs, content);
  };
  for (const { node } of marks) {
    const id = node.attrs.id as string;
    if (newIds.has(id)) continue;
    const newId = String(next++);
    newIds.set(id, newId);
    const note = from.notes.get(id);
    into.notes.set(newId, note ? restyle(note) : schema.nodes.doc.create(null, schema.nodes.paragraph.create({ style: newText })));
  }
  // Links and pictures in the notes keep their relationships.
  for (const [rid, rel] of from.rels) if (!into.rels.has(rid)) into.rels.set(rid, rel);
  const tr = view.state.tr;
  for (const { pos, node } of marks.reverse()) {
    let kept = node.marks.filter((m) => !(m.type === schema.marks.charStyle && m.attrs.id === oldRef));
    if (to === "footnote") kept = [...schema.marks.charStyle.create({ id: newRef }).addToSet(kept)];
    tr.replaceWith(pos, pos + node.nodeSize, toType.create({ id: newIds.get(node.attrs.id) }, null, kept));
  }
  // The list of endnotes at the end: added for the first endnotes, gone with the last.
  let list = -1;
  tr.doc.forEach((n, pos) => {
    if (n.type === schema.nodes.endnotes) list = pos;
  });
  if (to === "endnote" && list < 0) tr.insert(tr.doc.content.size, schema.nodes.endnotes.create());
  if (to === "footnote" && list >= 0) tr.delete(list, list + tr.doc.nodeAt(list)!.nodeSize);
  view.dispatch(tr.setMeta("noTrack", true));
  refresh(view);
  return newIds.size;
}
