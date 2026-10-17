// The navigation pane: the document's headings, as an outline beside the page.
// Clicking one (or Enter) goes there; the one the cursor is under is marked.

import type { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { h } from "../dom";
import { listLabels } from "../editor/lists";
import { schema } from "../editor/schema";
import { headingLevel } from "./references";

const OPEN_KEY = "duckwriter.navigation";

export interface NavigationHost {
  readonly body: EditorView;
  readonly ctx: DocContext;
  goTo(pos: number): void;
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
      }
    });
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
            const item = h("button", { type: "button", class: `nav-item level-${Math.min(e.level, 4)}`, role: "treeitem", title: e.text }, e.text);
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
