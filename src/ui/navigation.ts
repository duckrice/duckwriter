// The navigation pane: the document's headings, as an outline beside the page.
// Clicking one (or Enter) goes there; the one the cursor is under is marked.

import type { Command } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { h } from "../dom";
import { listLabels } from "../editor/lists";
import { moveSection, moveSectionCommand, shiftSection } from "../editor/outline";
import { schema } from "../editor/schema";
import { headingLevel } from "./references";

const OPEN_KEY = "duckwriter.navigation";

export interface NavigationHost {
  readonly body: EditorView;
  readonly ctx: DocContext;
  goTo(pos: number): void;
  /** Runs a command on the text, as if typed there. */
  run(command: Command): boolean;
}

interface Entry {
  pos: number;
  level: number;
  text: string;
}

export class NavigationPane {
  readonly el = h("nav", { class: "nav-pane", hidden: true, "aria-label": "Headings" });
  private list = h("div", { class: "nav-list", role: "tree" });
  private entries: Entry[] = [];
  private timer: number | undefined;
  private shown: unknown = null;

  constructor(private host: NavigationHost) {
    this.el.append(h("div", { class: "nav-title" }, "Headings"), this.list);
    try {
      this.el.hidden = localStorage.getItem(OPEN_KEY) !== "open";
    } catch {
      // Closed.
    }
    this.list.addEventListener("keydown", (event) => {
      const items = [...this.list.querySelectorAll<HTMLElement>(".nav-item")];
      const i = items.indexOf(document.activeElement as HTMLElement);
      if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        event.preventDefault();
        items[Math.max(0, Math.min(items.length - 1, i + (event.key === "ArrowDown" ? 1 : -1)))]?.focus();
      } else if (event.key === "Home" || event.key === "End") {
        event.preventDefault();
        items[event.key === "Home" ? 0 : items.length - 1]?.focus();
      } else if ((event.key === "Enter" || event.key === " ") && i >= 0) {
        event.preventDefault();
        items[i].click();
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.host.body.focus();
      } else if (event.altKey && event.shiftKey && !event.ctrlKey && event.key.startsWith("Arrow") && i >= 0) {
        // Moving the chosen section, or taking its headings a level up or down.
        event.preventDefault();
        const { ctx } = this.host;
        const pos = this.entries[i].pos;
        const command =
          event.key === "ArrowUp" ? moveSection(ctx, -1, pos)
          : event.key === "ArrowDown" ? moveSection(ctx, 1, pos)
          : shiftSection(ctx, event.key === "ArrowLeft" ? -1 : 1, pos);
        this.act(command, event.key === "ArrowUp" || event.key === "ArrowDown" ? null : i);
      }
    });
    // Dragging a heading moves its section: before the heading it is dropped on, or to the end.
    this.list.addEventListener("dragstart", (event) => {
      const item = (event.target as HTMLElement).closest<HTMLElement>(".nav-item");
      if (!item) return;
      this.dragging = [...this.list.querySelectorAll(".nav-item")].indexOf(item);
      event.dataTransfer?.setData("text/x-duckwriter-section", String(this.dragging));
      if (event.dataTransfer) event.dataTransfer.effectAllowed = "move";
    });
    this.list.addEventListener("dragover", (event) => {
      if (this.dragging < 0) return;
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = "move";
      this.showDrop(this.dropIndex(event.clientY));
    });
    this.list.addEventListener("dragleave", (event) => {
      if (!this.list.contains(event.relatedTarget as globalThis.Node | null)) this.showDrop(-1);
    });
    this.list.addEventListener("dragend", () => {
      this.dragging = -1;
      this.showDrop(-1);
    });
    this.list.addEventListener("drop", (event) => {
      if (this.dragging < 0) return;
      event.preventDefault();
      const from = this.entries[this.dragging]?.pos;
      const j = this.dropIndex(event.clientY);
      this.dragging = -1;
      this.showDrop(-1);
      if (from === undefined) return;
      const to = j < this.entries.length ? this.entries[j].pos : this.endOfText();
      this.act(moveSectionCommand(this.host.ctx, from, to), null);
    });
  }

  private dragging = -1;

  /** Where the end of the text is, before the list of endnotes if there is one. */
  private endOfText(): number {
    const doc = this.host.body.state.doc;
    const last = doc.lastChild;
    return last?.type === schema.nodes.endnotes ? doc.content.size - last.nodeSize : doc.content.size;
  }

  /** The heading a drop at this height goes before (the number of headings: at the end). */
  private dropIndex(y: number): number {
    const items = [...this.list.querySelectorAll<HTMLElement>(".nav-item")];
    const j = items.findIndex((el) => {
      const r = el.getBoundingClientRect();
      return y < r.top + r.height / 2;
    });
    return j < 0 ? items.length : j;
  }

  private showDrop(j: number) {
    const items = [...this.list.querySelectorAll<HTMLElement>(".nav-item")];
    items.forEach((el, k) => {
      el.classList.toggle("drop-before", k === j);
      el.classList.toggle("drop-after", j === items.length && k === items.length - 1);
    });
  }

  /** Runs an outline command, then shows the headings again with the moved (or same) one chosen. */
  private act(command: Command, keep: number | null): boolean {
    if (!this.host.run(command)) return false;
    this.render();
    const items = [...this.list.querySelectorAll<HTMLElement>(".nav-item")];
    const target = keep ?? this.currentIndex();
    items[Math.max(0, Math.min(items.length - 1, target))]?.focus();
    return true;
  }

  private currentIndex(): number {
    const at = this.host.body.state.selection.from;
    let current = 0;
    this.entries.forEach((e, i) => {
      if (e.pos < at) current = i;
    });
    return current;
  }

  get open(): boolean {
    return !this.el.hidden;
  }

  /** Shows or hides the pane; shown, the keyboard goes to the current heading. */
  toggle(focus = true) {
    this.el.hidden = !this.el.hidden;
    try {
      localStorage.setItem(OPEN_KEY, this.el.hidden ? "closed" : "open");
    } catch {
      // Not remembered.
    }
    if (this.el.hidden) return this.host.body.focus();
    this.render();
    if (focus) (this.list.querySelector<HTMLElement>(".nav-item.current") ?? this.list.querySelector<HTMLElement>(".nav-item"))?.focus();
  }

  /** After a change: the headings again, soon; the current one at once. */
  update(docChanged: boolean) {
    if (this.el.hidden) return;
    if (docChanged) {
      clearTimeout(this.timer);
      this.timer = window.setTimeout(() => this.render(), 250);
    } else this.markCurrent();
  }

  render() {
    const { body, ctx } = this.host;
    const doc = body.state.doc;
    if (this.shown === doc && this.list.childElementCount) return this.markCurrent();
    this.shown = doc;
    const labels = listLabels(ctx, doc);
    this.entries = [];
    doc.forEach((para, pos) => {
      if (para.type !== schema.nodes.paragraph) return;
      const level = para.attrs.style === "Title" ? 1 : headingLevel(ctx, para);
      const text = para.textContent.trim();
      if (!level || level > 6 || !text) return;
      const number = labels.get(pos)?.label;
      this.entries.push({ pos, level, text: number ? `${number} ${text}` : text });
    });
    const focused = this.list.contains(document.activeElement) ? [...this.list.children].indexOf(document.activeElement as Element) : -1;
    this.list.replaceChildren(
      ...(this.entries.length
        ? this.entries.map((e) => {
            const item = h("button", { type: "button", class: `nav-item level-${Math.min(e.level, 4)}`, role: "treeitem", title: `${e.text}
Drag to move this section; Alt+Shift+arrows move it or change its level`, draggable: "true" }, e.text);
            item.addEventListener("click", () => this.host.goTo(e.pos));
            return item;
          })
        : [h("div", { class: "nav-empty" }, "Headings (Heading 1, 2, 3…) are listed here, to go to with a click.")]),
    );
    if (focused >= 0) (this.list.children[Math.min(focused, this.list.children.length - 1)] as HTMLElement | undefined)?.focus();
    this.markCurrent();
  }

  /** Marks the heading of the part the cursor is in. */
  private markCurrent() {
    const at = this.host.body.state.selection.from;
    let current = -1;
    this.entries.forEach((e, i) => {
      if (e.pos < at) current = i;
    });
    [...this.list.querySelectorAll(".nav-item")].forEach((el, i) => el.classList.toggle("current", i === current));
  }
}
