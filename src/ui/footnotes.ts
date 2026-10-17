// Footnotes at the bottom of each page, numbered in the order their marks
// appear. Click a note (or its mark in the text) to edit it; Ctrl+Alt+F adds
// a new one, as in Word.

import { DOMSerializer, type Node } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { h } from "../dom";
import { ensureStyles } from "../docx/builtins";
import type { DocContext } from "../docx/model";
import { createState, createView } from "../editor/editor";
import { NOTE_SEPARATOR, PAGE_GAP, relayout, type PageLayout } from "../editor/pagination";
import { schema } from "../editor/schema";

export interface FootnoteHost {
  readonly ctx: DocContext;
  readonly body: EditorView;
  /** Called when a note changed, or editing started or stopped. */
  storiesChanged(): void;
}

const PX = 4 / 3;

export class Footnotes {
  readonly layer = h("div", { class: "fn-layer" });
  private measurer = h("div", { class: "fn-measure", "aria-hidden": "true" });
  private heights = new WeakMap<Node, number>();
  private layout: PageLayout = { count: 1, notes: [[]], pageOfBlock: new Map() };
  private editing: { id: string; view: EditorView } | null = null;

  constructor(private host: FootnoteHost) {
    this.layer.append(this.measurer);
    this.layer.addEventListener("mousedown", (event) => {
      const note = (event.target as HTMLElement).closest<HTMLElement>(".fn-note");
      if (!note || note.classList.contains("editing")) return;
      event.preventDefault();
      this.start(note.dataset.id!);
    });
  }

  get view(): EditorView | null {
    return this.editing?.view ?? null;
  }

  /** The note ids in the order their marks appear in the text. */
  private order(): string[] {
    const ids: string[] = [];
    this.host.body.state.doc.descendants((n) => {
      if (n.type === schema.nodes.footnote && !ids.includes(n.attrs.id)) ids.push(n.attrs.id);
    });
    return ids;
  }

  private numberOf(id: string): number {
    return this.order().indexOf(id) + 1;
  }

  /** A note as shown at the bottom of a page: its number, then its text. */
  private draw(id: string, doc: Node): HTMLElement {
    const content = h("div", { class: "hf-doc fn-doc" });
    content.append(DOMSerializer.fromSchema(schema).serializeFragment(doc.content));
    const first = content.querySelector(".para");
    const number = h("span", { class: "fn-number" }, String(this.numberOf(id)));
    if (first) first.prepend(number);
    else content.prepend(number);
    return content;
  }

  /** How tall a note is, in pixels, for the page layout. */
  height(id: string): number {
    const doc = this.host.ctx.footnotes.notes.get(id);
    if (!doc) return 0;
    const known = this.heights.get(doc);
    if (known !== undefined) return known;
    const p = this.host.ctx.page;
    this.measurer.style.width = `${p.width - p.marginLeft - p.marginRight}pt`;
    this.measurer.replaceChildren(this.draw(id, doc));
    const height = this.measurer.offsetHeight;
    this.measurer.replaceChildren();
    this.heights.set(doc, height);
    return height;
  }

  /** Draws the notes at the bottom of each page. */
  render(layout: PageLayout) {
    this.layout = layout;
    const p = this.host.ctx.page;
    const pitch = p.height * PX + PAGE_GAP;
    const areas = [...this.layer.querySelectorAll<HTMLElement>(".fn-area")];
    while (areas.length > layout.count) areas.pop()!.remove();
    for (let k = 0; k < layout.count; k++) {
      let area = areas[k];
      if (!area) {
        area = h("div", { class: "fn-area" });
        this.layer.append(area);
        areas.push(area);
      }
      area.style.left = `${p.marginLeft}pt`;
      area.style.right = `${p.marginRight}pt`;
      area.style.top = `${k * pitch + (p.height - p.marginBottom) * PX}px`;
      const ids = layout.notes[k] ?? [];
      area.hidden = !ids.length;
      const editing = this.editing ? area.querySelector<HTMLElement>(`.fn-note.editing[data-id="${this.editing.id}"]`) : null;
      // The note being typed in stays where it is: moving it would lose the cursor.
      const separator = area.querySelector(".fn-separator") ?? h("div", { class: "fn-separator" });
      if (!separator.parentNode) area.prepend(separator);
      for (const old of [...area.querySelectorAll<HTMLElement>(".fn-note")]) if (old !== editing) old.remove();
      for (const id of ids) {
        if (editing && editing.dataset.id === id) continue;
        const doc = this.host.ctx.footnotes.notes.get(id);
        const note = h("div", { class: "fn-note", "data-id": id, title: "Click to edit the footnote" });
        if (doc) note.append(this.draw(id, doc));
        const before = editing && ids.indexOf(editing.dataset.id!) > ids.indexOf(id) ? editing : null;
        area.insertBefore(note, before);
      }
      if (editing && !ids.includes(editing.dataset.id!)) editing.remove();
      area.style.minHeight = `${NOTE_SEPARATOR}px`;
    }
  }

  /** Edits a note where it is shown. */
  start(id: string) {
    this.stop();
    const doc = this.host.ctx.footnotes.notes.get(id);
    const note = this.layer.querySelector<HTMLElement>(`.fn-note[data-id="${id}"]`);
    if (!doc || !note) return;
    note.classList.add("editing");
    note.replaceChildren();
    const number = h("span", { class: "fn-number editing-number" }, String(this.numberOf(id)));
    note.append(number);
    const view = createView(
      note,
      createState(doc, () => this.host.ctx),
      {
        onChange: (v, docChanged) => {
          if (docChanged) {
            this.host.ctx.footnotes.notes.set(id, v.state.doc);
            // The note may now take more or less room on its page.
            relayout(this.host.body);
          }
          this.host.storiesChanged();
        },
      },
      "hf-doc fn-doc",
    );
    view.dom.addEventListener("keydown", (event) => {
      if ((event as KeyboardEvent).key === "Escape") {
        event.preventDefault();
        this.stop();
        // Back to the text, after the note's mark.
        this.host.body.focus();
      }
    });
    view.dom.addEventListener("focusout", () => {
      // Clicking elsewhere ends editing, once the click has had its effect.
      setTimeout(() => {
        if (this.editing?.view === view && !view.hasFocus()) this.stop();
      }, 150);
    });
    this.editing = { id, view };
    view.focus();
    // The cursor goes at the end of the note.
    const end = view.state.doc.content.size - 1;
    view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(Math.max(1, end)))));
    this.host.storiesChanged();
  }

  stop() {
    const editing = this.editing;
    if (!editing) return;
    this.editing = null;
    this.host.ctx.footnotes.notes.set(editing.id, editing.view.state.doc);
    editing.view.destroy();
    this.render(this.layout);
    this.host.storiesChanged();
  }

  /** Puts a new footnote at the cursor in the text, and opens it for typing. */
  insert() {
    const { ctx, body } = this.host;
    ensureStyles(ctx, "FootnoteText", "FootnoteReference");
    const ids = [...ctx.footnotes.notes.keys(), ...this.order()].map(Number).filter(Number.isFinite);
    const id = String(Math.max(0, ...ids) + 1);
    ctx.footnotes.notes.set(id, schema.nodes.doc.create(null, schema.nodes.paragraph.create({ style: "FootnoteText" }, schema.text(" "))));
    const mark = schema.marks.charStyle.create({ id: "FootnoteReference" });
    const node = schema.nodes.footnote.create({ id }, null, [mark]);
    const tr = body.state.tr.replaceSelectionWith(node, false);
    // Text typed after the mark isn't superscript.
    tr.setStoredMarks(body.state.selection.$from.marks().filter((m) => m.type !== schema.marks.charStyle));
    body.dispatch(tr.scrollIntoView());
    // Once the page layout has placed it, open the note.
    setTimeout(() => {
      this.start(id);
      this.layer.querySelector(`.fn-note[data-id="${id}"]`)?.scrollIntoView({ block: "nearest" });
    }, 50);
  }
}
