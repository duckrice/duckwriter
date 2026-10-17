// The find (Ctrl+F) and replace (Ctrl+H) bar above the page.

import type { EditorView } from "prosemirror-view";
import { h } from "../dom";
import { findState, replaceAll, replaceCurrent, search, step, type FindOptions } from "../editor/find";

export interface FindHost {
  readonly body: EditorView;
  toast(text: string): void;
}

export class FindBar {
  readonly el: HTMLElement;
  private findInput = h("input", { class: "tb-input find-input", placeholder: "Find", "aria-label": "Find", spellcheck: "false" });
  private replaceInput = h("input", { class: "tb-input find-input", placeholder: "Replace with", "aria-label": "Replace with", spellcheck: "false" });
  private count = h("span", { class: "find-count" });
  private caseBox = h("input", { type: "checkbox" }) as HTMLInputElement;
  private wordBox = h("input", { type: "checkbox" }) as HTMLInputElement;
  private replaceRow: HTMLElement;

  constructor(private host: FindHost) {
    const button = (label: string, title: string, onClick: () => void, cls = "btn small") => {
      const b = h("button", { type: "button", class: cls, title }, label);
      b.addEventListener("mousedown", (event) => event.preventDefault());
      b.addEventListener("click", onClick);
      return b;
    };
    this.replaceRow = h(
      "div",
      { class: "find-row" },
      this.replaceInput,
      button("Replace", "Replace this one and go to the next (Enter)", () => this.replace()),
      button("Replace all", "Replace every one", () => this.replaceEvery()),
    );
    const bar = h(
      "div",
      { class: "find-bar", role: "search" },
      h(
        "div",
        { class: "find-row" },
        this.findInput,
        this.count,
        button("↑", "Previous (Shift+Enter, Shift+F3)", () => this.go(-1), "btn small find-step"),
        button("↓", "Next (Enter, F3)", () => this.go(1), "btn small find-step"),
        h("label", { class: "find-option", title: "Match upper and lower case" }, this.caseBox, "Match case"),
        h("label", { class: "find-option", title: "Only whole words" }, this.wordBox, "Whole words"),
        button("×", "Close (Esc)", () => this.close(), "btn subtle small find-close"),
      ),
      this.replaceRow,
    );
    // A zero-height strip that stays at the top while the page scrolls under it.
    this.el = h("div", { class: "find-dock", hidden: true }, bar);
    this.findInput.addEventListener("input", () => this.run());
    this.caseBox.addEventListener("change", () => this.run());
    this.wordBox.addEventListener("change", () => this.run());
    this.findInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.go(event.shiftKey ? -1 : 1);
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.close();
      }
    });
    this.replaceInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.replace();
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.close();
      }
    });
  }

  get isOpen(): boolean {
    return !this.el.hidden;
  }

  private options(): FindOptions {
    return { caseSensitive: this.caseBox.checked, wholeWord: this.wordBox.checked };
  }

  /** Opens the bar (with the replace row for Ctrl+H), starting from the selected text. */
  open(replace: boolean) {
    const { state } = this.host.body;
    const { from, to, empty } = state.selection;
    const selected = empty ? "" : state.doc.textBetween(from, to, " ");
    if (selected && selected.length < 200 && !selected.includes("\n")) this.findInput.value = selected;
    this.el.hidden = false;
    this.replaceRow.hidden = !replace;
    this.run();
    const input = replace && this.findInput.value ? this.replaceInput : this.findInput;
    input.focus();
    input.select();
  }

  close() {
    if (this.el.hidden) return;
    this.el.hidden = true;
    const view = this.host.body;
    view.dispatch(search(view.state, "", this.options()));
    view.focus();
  }

  private run() {
    const view = this.host.body;
    view.dispatch(search(view.state, this.findInput.value, this.options()));
    this.update();
  }

  /** Goes to the next (1) or previous (-1) match. */
  go(direction: 1 | -1) {
    const view = this.host.body;
    if (this.el.hidden) return this.open(false);
    const tr = step(view.state, direction);
    if (tr) view.dispatch(tr);
    this.update();
  }

  private replace() {
    const view = this.host.body;
    const tr = replaceCurrent(view.state, this.replaceInput.value);
    if (!tr) return;
    view.dispatch(tr);
    // On to the next one, selected, ready to be replaced in turn.
    const next = step(view.state, 1);
    if (next) view.dispatch(next);
    this.update();
  }

  private replaceEvery() {
    const view = this.host.body;
    const result = replaceAll(view.state, this.replaceInput.value);
    if (!result) return;
    view.dispatch(result.tr);
    this.host.toast(`Replaced ${result.count} ${result.count === 1 ? "time" : "times"}`);
    this.update();
  }

  /** Shows "3 of 12", or that there is nothing. */
  update() {
    if (this.el.hidden) return;
    const find = findState(this.host.body.state);
    const n = find?.matches.length ?? 0;
    this.count.textContent = !this.findInput.value ? "" : n ? `${(find?.current ?? -1) + 1 || "–"} of ${n}` : "Not found";
    this.count.classList.toggle("none", !!this.findInput.value && !n);
  }
}
