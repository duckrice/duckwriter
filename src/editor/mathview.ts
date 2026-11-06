// Shows an equation in the text as typeset math, and lets you write in it
// straight away, as in Word: click into it (or arrow into it) and type. The
// shorthand builds up as you go (math/edit.ts); the arrows move between the
// boxes, and moving past either end goes back to the text.

import { NodeSelection, TextSelection } from "prosemirror-state";
import { redo, undo } from "prosemirror-history";
import type { Node } from "prosemirror-model";
import type { EditorView, NodeView } from "prosemirror-view";
import { MathSyntaxError, parseMath } from "../math/ast";
import { BIG, MathEdit, fromLatex, fromNode, type Item, type Row } from "../math/edit";
import { toMathML } from "../math/mathml";

export type EditMath = (pos: number) => void;
type Where = "start" | "end";

const NS = "http://www.w3.org/1998/Math/MathML";
function el(tag: string, attrs: Record<string, string> = {}, ...children: (Element | string)[]): Element {
  const e = document.createElementNS(NS, tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  e.append(...children);
  return e;
}
const BRACKETS = "()[]{}|‖⟨⟩⌊⌋⌈⌉";
const ACCENT_SHOWN: Record<string, string> = { "⃗": "→", "̂": "^", "̅": "‾", "̇": "˙", "̈": "¨", "̃": "~" };

/** Starts writing in the equation at `pos`, with the cursor at its start or end. */
export function enterMath(view: EditorView, pos: number, where: Where): boolean {
  const dom = view.nodeDOM(pos) as (HTMLElement & { mathView?: MathView }) | null;
  if (!dom?.mathView) return false;
  dom.mathView.activate(where);
  return true;
}

export class MathView implements NodeView {
  dom: HTMLElement & { mathView?: MathView };
  private edit: MathEdit | null = null;
  private input: HTMLTextAreaElement | null = null;
  /** The LaTeX this view last put in the document, to tell its own changes from undo or another writer's. */
  private written = "";

  constructor(
    private node: Node,
    private view: EditorView,
    private getPos: () => number | undefined,
    private editLatex: EditMath,
  ) {
    this.dom = document.createElement("span");
    this.dom.mathView = this;
    this.dom.addEventListener("mousedown", (event) => {
      if (event.button !== 0 || !this.view.editable) return;
      event.preventDefault();
      this.activate("end");
      this.placeAt(event.clientX, event.clientY);
      this.afterChange(false);
    });
    this.draw();
  }

  // ---- writing in it --------------------------------------------------------------

  activate(where: Where) {
    if (this.edit) return;
    const pos = this.getPos();
    if (pos === undefined) return;
    this.edit = new MathEdit(fromLatex(this.node.attrs.latex));
    if (where === "start") this.edit.caret = { path: [], index: 0 };
    this.written = this.node.attrs.latex;
    const input = document.createElement("textarea");
    input.className = "math-input";
    input.setAttribute("aria-label", "Equation");
    input.spellcheck = false;
    input.autocapitalize = "off";
    document.body.append(input);
    this.input = input;
    input.addEventListener("keydown", (event) => this.key(event));
    input.addEventListener("input", (event) => {
      if (!(event as InputEvent).isComposing) this.takeText();
    });
    input.addEventListener("compositionend", () => this.takeText());
    input.addEventListener("paste", (event) => {
      event.preventDefault();
      const text = event.clipboardData?.getData("text/plain") ?? "";
      try {
        this.edit?.insertRow(fromNode(parseMath(text)));
      } catch {
        this.edit?.type(text);
      }
      this.afterChange(true);
    });
    // A click in the text leaves the equation; a click on the ribbon keeps it selected, to format it.
    input.addEventListener("blur", (event) => {
      const to = event.relatedTarget as globalThis.Node | null;
      this.deactivate(null, !!to && this.view.dom.contains(to));
    });
    // Selected in the document, so the ribbon knows which equation it is.
    this.view.dispatch(this.view.state.tr.setSelection(NodeSelection.create(this.view.state.doc, pos)));
    this.dom.classList.add("editing");
    input.focus({ preventScroll: true });
    this.afterChange(false);
  }

  /** Stops writing; `side` puts the text cursor before or after the equation (null: leave it where a click put it). */
  deactivate(side: "before" | "after" | null, toText = true) {
    const edit = this.edit;
    if (!edit) return;
    edit.finish();
    this.commit();
    this.edit = null;
    const input = this.input;
    this.input = null;
    input?.remove();
    this.dom.classList.remove("editing");
    this.draw();
    const pos = this.getPos();
    if (pos === undefined) return;
    const tr = this.view.state.tr;
    if (edit.isEmpty()) tr.delete(pos, pos + this.node.nodeSize);
    const at = edit.isEmpty() || side === "before" ? pos : pos + this.node.nodeSize;
    // Clicked away: the click places the cursor, but never leave the equation selected (typing would replace it).
    const stillSelected = toText && tr.selection instanceof NodeSelection && tr.selection.from === pos;
    if (side || edit.isEmpty() || stillSelected) tr.setSelection(TextSelection.create(tr.doc, Math.min(at, tr.doc.content.size)));
    if (tr.docChanged || tr.selectionSet) this.view.dispatch(tr.scrollIntoView());
    if (side) this.view.focus();
  }

  private takeText() {
    const input = this.input;
    if (!input || !this.edit || !input.value) return;
    const text = input.value;
    input.value = "";
    this.edit.type(text);
    this.afterChange(true);
  }

  private key(event: KeyboardEvent) {
    const edit = this.edit;
    if (!edit || event.isComposing) return;
    const ctrl = event.ctrlKey || event.metaKey;
    if (ctrl && !event.altKey && (event.key === "z" || event.key === "y" || event.key === "Z")) {
      event.preventDefault();
      this.deactivate("after");
      (event.key === "y" || event.shiftKey ? redo : undo)(this.view.state, this.view.dispatch);
      return;
    }
    if (ctrl || event.altKey) return;
    let changed = false;
    switch (event.key) {
      case "ArrowLeft":
        if (!edit.left()) return this.leave(event, "before");
        break;
      case "ArrowRight":
        if (!edit.right()) return this.leave(event, "after");
        break;
      case "ArrowUp":
      case "ArrowDown":
        if (!edit.vertical(event.key === "ArrowUp" ? -1 : 1)) return this.leave(event, event.key === "ArrowUp" ? "before" : "after");
        break;
      case "Home":
        edit.home();
        break;
      case "End":
        edit.end();
        break;
      case "Backspace":
        if (edit.isEmpty()) return this.leave(event, "before");
        edit.backspace();
        changed = true;
        break;
      case "Delete":
        edit.del();
        changed = true;
        break;
      case "Tab":
        if (!edit.nextCell()) return this.leave(event, "after");
        break;
      case "Enter":
        if (edit.addMatrixRow()) {
          changed = true;
          break;
        }
        // Finishes the equation; Enter again starts a new paragraph.
        return this.leave(event, "after");
      case "Escape":
        return this.leave(event, "after");
      default:
        return;
    }
    event.preventDefault();
    this.afterChange(changed);
  }

  private leave(event: KeyboardEvent, side: "before" | "after") {
    event.preventDefault();
    this.deactivate(side);
  }

  /** Puts the change in the document (so it is saved and can be undone) and redraws. */
  private afterChange(changed: boolean) {
    if (changed) this.commit();
    this.draw();
    // The hidden box follows the cursor, so the IME window opens in the right place.
    const caret = this.dom.querySelector(".m-caret");
    if (caret && this.input) {
      const r = caret.getBoundingClientRect();
      this.input.style.left = `${r.left}px`;
      this.input.style.top = `${r.top}px`;
    }
  }

  private commit() {
    const pos = this.getPos();
    if (!this.edit || pos === undefined) return;
    const latex = this.edit.latex();
    if (latex === this.node.attrs.latex) return;
    this.written = latex;
    this.view.dispatch(this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, latex }));
  }

  /** The cursor where the equation was clicked: before or after the nearest thing in it. */
  private placeAt(x: number, y: number) {
    const edit = this.edit;
    if (!edit) return;
    let best: { d: number; path: number[]; index: number } | null = null;
    for (const target of this.dom.querySelectorAll<Element>("[data-at]")) {
      const r = target.getBoundingClientRect();
      const dy = y < r.top ? r.top - y : y > r.bottom ? y - r.bottom : 0;
      const dx = x < r.left ? r.left - x : x > r.right ? x - r.right : 0;
      // Smaller things (inside boxes) win when the click is on them.
      const d = dy * 4 + dx + (r.width * r.height) / 1e5;
      if (!best || d < best.d) {
        const [path, index] = JSON.parse(target.getAttribute("data-at")!) as [number[], number];
        const after = target.hasAttribute("data-item") && x > r.left + r.width / 2;
        best = { d, path, index: index + (after ? 1 : 0) };
      }
    }
    if (best) edit.caret = { path: best.path, index: best.index };
    edit.marked = false;
  }

  // ---- drawing --------------------------------------------------------------------------

  private draw() {
    const { latex, display } = this.node.attrs;
    this.dom.className = `math${display ? " display" : ""}${this.edit ? " editing" : ""}`;
    this.dom.title = "";
    this.dom.replaceChildren();
    if (this.edit) {
      this.dom.append(el("math", display ? { display: "block" } : {}, this.drawRow(this.edit.root, [])));
      return;
    }
    if (!latex.trim()) {
      this.dom.classList.add("empty");
      this.dom.textContent = "Type an equation";
      return;
    }
    try {
      this.dom.append(toMathML(parseMath(latex), display));
    } catch (error) {
      this.dom.classList.add("error");
      this.dom.textContent = latex;
      this.dom.title = error instanceof MathSyntaxError ? `${error.message}. Click to fix.` : "";
    }
  }

  private drawRow(row: Row, path: number[]): Element {
    const edit = this.edit!;
    const here = edit.caret.path.length === path.length && edit.caret.path.every((v, i) => v === path[i]);
    const at = (index: number) => JSON.stringify([path, index]);
    const children: Element[] = [];
    const caret = () => el("mspace", { class: "m-caret", width: "0", height: "0.8em", depth: "0.2em" });
    row.forEach((item, i) => {
      if (here && edit.caret.index === i) children.push(caret());
      const e = this.drawItem(item, [...path], i);
      e.setAttribute("data-at", at(i));
      e.setAttribute("data-item", "");
      if (here && edit.marked && edit.caret.index === i + 1) e.classList.add("m-marked");
      children.push(e);
    });
    if (here && edit.caret.index === row.length) children.push(caret());
    // An empty box shows as a dotted square to type into.
    if (!row.length) return el("mrow", { class: "m-slot", "data-at": at(0) }, ...children, el("mspace", { width: "0.5em", height: "0.7em", depth: "0.1em" }));
    return el("mrow", {}, ...children);
  }

  private drawItem(item: Item, path: number[], i: number): Element {
    if (item.k === "ch") {
      if (/^[0-9.]$/.test(item.v)) return el("mn", {}, item.v);
      if (/^\p{L}$/u.test(item.v)) return el("mi", {}, item.v);
      if (item.v === " ") return el("mspace", { width: "0.5em" });
      const v = item.v === "-" ? "−" : item.v === "*" ? "∗" : item.v;
      if (BIG.has(v)) return el("mo", { largeop: "true", movablelimits: "true" }, v);
      return el("mo", BRACKETS.includes(v) ? { lspace: "0", rspace: "0" } : {}, v);
    }
    if (item.k === "sym") return item.text ? el("mtext", {}, item.v) : el("mi", { mathvariant: "normal" }, item.v);
    const slot = (s: number) => this.drawRow(item.slots[s], [...path, i, s]);
    const display = !!this.node.attrs.display;
    // ∑ and lim take their limits above and below when on their own line.
    const first = item.slots[0]?.[0];
    const under = display && item.slots[0]?.length === 1 && ((first?.k === "ch" && BIG.has(first.v) && first.v[0] !== "∫") || (first?.k === "sym" && !first.text));
    switch (item.kind) {
      case "frac":
        return el("mfrac", {}, slot(0), slot(1));
      case "sqrt":
        return el("msqrt", {}, slot(0));
      case "root":
        return el("mroot", {}, slot(1), slot(0));
      case "sub":
        return el(under ? "munder" : "msub", {}, slot(0), slot(1));
      case "sup":
        return el(under ? "mover" : "msup", {}, slot(0), slot(1));
      case "subsup":
        return el(under ? "munderover" : "msubsup", {}, slot(0), slot(1), slot(2));
      case "delim":
        return el("mrow", {}, el("mo", { stretchy: "true" }, item.open ?? "("), slot(0), el("mo", { stretchy: "true" }, item.close ?? ")"));
      case "accent":
        return el("mover", { accent: "true" }, slot(0), el("mo", { stretchy: item.chr === "̅" ? "true" : "false" }, ACCENT_SHOWN[item.chr ?? ""] ?? item.chr ?? ""));
      case "matrix": {
        const cols = item.cols ?? 1;
        const rows: Element[] = [];
        for (let r = 0; r * cols < item.slots.length; r++)
          rows.push(el("mtr", {}, ...Array.from({ length: cols }, (_, c) => el("mtd", {}, slot(r * cols + c)))));
        const parts: Element[] = [el("mtable", {}, ...rows)];
        if (item.open) parts.unshift(el("mo", { stretchy: "true" }, item.open));
        if (item.close) parts.push(el("mo", { stretchy: "true" }, item.close));
        return el("mrow", {}, ...parts);
      }
    }
  }

  // ---- ProseMirror -----------------------------------------------------------------------

  update(node: Node): boolean {
    if (node.type !== this.node.type) return false;
    const changed = node.attrs.latex !== this.node.attrs.latex || node.attrs.display !== this.node.attrs.display;
    this.node = node;
    // Undo, or someone else editing the same equation: start again from what is there.
    if (this.edit && node.attrs.latex !== this.written) {
      this.edit = new MathEdit(fromLatex(node.attrs.latex));
      this.written = node.attrs.latex;
    }
    if (changed || this.edit) this.draw();
    return true;
  }

  selectNode() {
    this.dom.classList.add("selected");
  }

  deselectNode() {
    this.dom.classList.remove("selected");
  }

  ignoreMutation() {
    return true;
  }

  stopEvent(event: Event) {
    return event.type.startsWith("mouse") || event.type === "click" || event.type === "dblclick" || event.type.startsWith("drag");
  }

  destroy() {
    if (this.edit) {
      this.input?.remove();
      this.input = null;
      this.edit = null;
    }
  }

  /** The LaTeX dialog, for those who'd rather write the shorthand in one go. */
  editAsLatex() {
    const pos = this.getPos();
    if (pos !== undefined) this.editLatex(pos);
  }
}
