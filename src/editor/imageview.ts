// Pictures on the page. A picture in line with the text sits in its line; a
// floating one is placed where the document says, with text wrapping around
// it (square), above and below it, or not at all (in front of or behind the
// text). Pictures resize from their corners and floating ones move by
// dragging; both snap to a grid (hold Alt to place freely).

import type { Node } from "prosemirror-model";
import { NodeSelection } from "prosemirror-state";
import type { EditorView, NodeView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import type { Wrap } from "../docx/drawing";
import { PAGE_GAP } from "./pagination";

const PX = 4 / 3;
/** The grid pictures snap to: half a centimetre, in points. */
export const GRID = 14.173;

export interface PictureHooks {
  context(): DocContext;
  /** Whether pictures snap to the grid. */
  snap(): boolean;
}

/** All floating pictures on screen, to place again after each page layout. */
const floating = new Set<ImageView>();

export function placeFloatingPictures() {
  for (const view of floating) view.place();
}

const WRAPS: { wrap: Wrap | "inline"; label: string }[] = [
  { wrap: "inline", label: "In line with text" },
  { wrap: "square", label: "Square" },
  { wrap: "topAndBottom", label: "Top and bottom" },
  { wrap: "none", label: "In front of text" },
  { wrap: "behind", label: "Behind text" },
];

export class ImageView implements NodeView {
  dom: HTMLElement;
  private img: HTMLImageElement;
  private bar: HTMLElement | null = null;

  constructor(
    private node: Node,
    private view: EditorView,
    private getPos: () => number | undefined,
    private hooks: PictureHooks,
  ) {
    this.dom = document.createElement("span");
    this.img = document.createElement("img");
    this.img.draggable = false;
    this.img.alt = "";
    this.dom.append(this.img);
    for (const corner of ["nw", "ne", "sw", "se"]) {
      const handle = document.createElement("span");
      handle.className = `resize-handle ${corner}`;
      handle.addEventListener("mousedown", (event) => this.resize(event, corner));
      this.dom.append(handle);
    }
    this.img.addEventListener("mousedown", (event) => this.press(event));
    this.draw();
  }

  private draw() {
    const a = this.node.attrs;
    this.img.src = a.src;
    this.img.style.width = `${a.width}pt`;
    this.img.style.height = `${a.height}pt`;
    this.dom.className = a.floating ? `picture-box floating wrap-${a.wrap}` : "picture-box inline";
    if (a.floating) floating.add(this);
    else {
      floating.delete(this);
      this.dom.style.cssText = "";
      this.img.style.left = this.img.style.top = "";
    }
    this.place();
  }

  /** Puts a floating picture where its position says, by adjusting how its box sits in the text. */
  place() {
    const a = this.node.attrs;
    if (!a.floating || !this.dom.isConnected) return;
    const doc = this.view.dom as HTMLElement;
    const para = this.dom.closest(".para") as HTMLElement | null;
    if (!para) return;
    const ctx = this.hooks.context();
    const docRect = doc.getBoundingClientRect();
    const scale = doc.offsetWidth ? docRect.width / doc.offsetWidth : 1;
    const page = ctx.page;
    const pitch = page.height * PX + PAGE_GAP;
    const contentWidth = (page.width - page.marginLeft - page.marginRight) * PX;

    // Where it should be, in pixels from the top left of the text on the first page.
    const paraTop = (para.getBoundingClientRect().top - docRect.top) / scale;
    const pageIndex = Math.max(0, Math.floor(paraTop / pitch));
    const pageTop = pageIndex * pitch - page.marginTop * PX;
    let top = a.y * PX;
    if (a.relY === "paragraph" || a.relY === "line") top += paraTop;
    else if (a.relY === "page") top += pageTop;
    else top += pageIndex * pitch;
    let left = a.x * PX;
    if (a.relX === "page") left -= page.marginLeft * PX;

    // Where it is without adjustment.
    this.dom.style.marginTop = this.dom.style.marginLeft = this.dom.style.marginRight = "";
    const box = this.dom.getBoundingClientRect();
    const naturalTop = (box.top - docRect.top) / scale;
    const naturalLeft = (box.left - docRect.left) / scale;
    const width = a.width * PX;
    if (a.wrap === "square" || a.wrap === "topAndBottom") {
      // A box in the text that the text flows around.
      const shift = top - naturalTop;
      this.dom.style.marginTop = `${shift}px`;
      // Only the picture keeps the text away, not the space above it.
      this.dom.style.shapeOutside = a.wrap === "square" && shift > 0 ? `inset(${shift}px 0 0 0) margin-box` : "";
      if (a.wrap === "square" && left + width / 2 > contentWidth / 2) {
        this.dom.style.cssFloat = "right";
        this.dom.style.marginRight = `${Math.max(0, contentWidth - left - width)}px`;
        this.dom.style.marginLeft = "8px";
      } else {
        this.dom.style.cssFloat = a.wrap === "square" ? "left" : "none";
        const lineLeft = a.wrap === "square" ? naturalLeft : (para.getBoundingClientRect().left - docRect.left) / scale;
        this.dom.style.marginLeft = `${left - lineLeft}px`;
        this.dom.style.marginRight = a.wrap === "square" ? "8px" : "0";
      }
    } else {
      // In front of or behind the text: drawn there, taking no room.
      this.img.style.left = `${left - naturalLeft}px`;
      this.img.style.top = `${top - naturalTop}px`;
    }
  }

  // ---- moving and resizing ----------------------------------------------------------

  private select() {
    const pos = this.getPos();
    if (pos === undefined) return;
    this.view.dispatch(this.view.state.tr.setSelection(NodeSelection.create(this.view.state.doc, pos)));
  }

  private scale(): number {
    const doc = this.view.dom as HTMLElement;
    return doc.offsetWidth ? doc.getBoundingClientRect().width / doc.offsetWidth : 1;
  }

  private press(event: MouseEvent) {
    if (event.button !== 0) return;
    this.select();
    if (!this.node.attrs.floating) return;
    event.preventDefault();
    this.drag(event);
  }

  /** Drags a floating picture, snapping to the grid and to the margins and middle of the page. */
  private drag(start: MouseEvent) {
    const a = this.node.attrs;
    const scale = this.scale();
    const startX = a.x;
    const startY = a.y;
    const grid = showGrid(this.view, this.hooks.context());
    let x = startX;
    let y = startY;
    const move = (e: MouseEvent) => {
      x = startX + (e.clientX - start.clientX) / scale / PX;
      y = startY + (e.clientY - start.clientY) / scale / PX;
      if (this.hooks.snap() && !e.altKey) {
        const page = this.hooks.context().page;
        const contentWidth = page.width - page.marginLeft - page.marginRight;
        const offset = a.relX === "page" ? page.marginLeft : 0;
        const inMargin = x - offset;
        // The edges and the middle of the text area win over the grid when close.
        const targets = [0, contentWidth - a.width, (contentWidth - a.width) / 2];
        const near = targets.find((t) => Math.abs(inMargin - t) < 6);
        x = (near ?? Math.round(inMargin / GRID) * GRID) + offset;
        y = Math.round(y / GRID) * GRID;
      }
      this.dom.style.transform = `translate(${(x - startX) * PX}px, ${(y - startY) * PX}px)`;
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      grid.remove();
      this.dom.style.transform = "";
      if (x !== startX || y !== startY) this.change({ x: round(x), y: round(y) });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  /** Resizes from a corner, keeping the picture's proportions; the width snaps to the grid. */
  private resize(start: MouseEvent, corner: string) {
    start.preventDefault();
    start.stopPropagation();
    this.select();
    const a = this.node.attrs;
    const scale = this.scale();
    const ratio = a.height / a.width;
    const direction = corner.includes("w") ? -1 : 1;
    let width = a.width;
    const move = (e: MouseEvent) => {
      width = Math.max(12, a.width + (direction * (e.clientX - start.clientX)) / scale / PX);
      if (this.hooks.snap() && !e.altKey) width = Math.max(GRID, Math.round(width / GRID) * GRID);
      this.img.style.width = `${width}pt`;
      this.img.style.height = `${width * ratio}pt`;
    };
    const up = () => {
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      if (width !== a.width) this.change({ width: round(width), height: round(width * ratio) });
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  /** Changes the picture in the document (from a drag, a resize or the wrapping menu). */
  private change(attrs: Record<string, unknown>) {
    const pos = this.getPos();
    if (pos === undefined) return;
    const tr = this.view.state.tr.setNodeMarkup(pos, undefined, { ...this.node.attrs, ...attrs });
    tr.setSelection(NodeSelection.create(tr.doc, pos));
    this.view.dispatch(tr);
  }

  // ---- the picture bar ---------------------------------------------------------------

  private showBar() {
    this.bar?.remove();
    const bar = document.createElement("span");
    bar.className = "picture-bar";
    bar.contentEditable = "false";
    const current = this.node.attrs.floating ? this.node.attrs.wrap : "inline";
    for (const { wrap, label } of WRAPS) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = label;
      b.className = wrap === current ? "active" : "";
      b.addEventListener("mousedown", (event) => {
        event.preventDefault();
        event.stopPropagation();
        this.setWrap(wrap);
      });
      bar.append(b);
    }
    // A description for people who can't see the picture (alt text).
    const describe = document.createElement("button");
    describe.type = "button";
    describe.textContent = this.node.attrs.alt ? "Description ✓" : "Description…";
    describe.title = this.node.attrs.alt || "Describe the picture for people who can't see it (alt text)";
    describe.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      this.describe();
    });
    bar.append(describe);
    this.dom.append(bar);
    this.bar = bar;
  }

  /** Asks for the picture's description (alt text). */
  describe() {
    const pos = this.getPos();
    if (pos === undefined) return;
    const text = document.createElement("textarea");
    text.className = "tb-input dialog-input wide";
    text.rows = 3;
    text.value = this.node.attrs.alt;
    text.placeholder = "What the picture shows, in a sentence or two";
    const dialog = document.createElement("dialog");
    dialog.className = "dialog";
    const title = document.createElement("h2");
    title.textContent = "Describe the picture";
    const hint = document.createElement("p");
    hint.className = "dialog-hint";
    hint.textContent = "Read aloud by screen readers, and shown where the picture can't be. Word calls it alt text.";
    const footer = document.createElement("div");
    footer.className = "dialog-footer";
    const cancel = document.createElement("button");
    cancel.type = "button";
    cancel.className = "btn";
    cancel.textContent = "Cancel";
    cancel.onclick = () => dialog.close();
    const ok = document.createElement("button");
    ok.type = "button";
    ok.className = "btn primary";
    ok.textContent = "OK";
    ok.onclick = () => {
      const at = this.getPos();
      if (at !== undefined) this.view.dispatch(this.view.state.tr.setNodeMarkup(at, undefined, { ...this.node.attrs, alt: text.value.trim() }));
      dialog.close();
    };
    footer.append(cancel, ok);
    dialog.append(title, hint, text, footer);
    dialog.addEventListener("close", () => {
      dialog.remove();
      this.view.focus();
    });
    document.body.append(dialog);
    dialog.showModal();
    text.focus();
  }

  /** Changes how text goes around the picture, keeping it where it is on the page. */
  private setWrap(wrap: Wrap | "inline") {
    const a = this.node.attrs;
    if (wrap === "inline") return this.change({ floating: false });
    if (a.floating) return this.change({ wrap });
    // From in line to floating: where it is now, measured from the paragraph and the margin.
    const doc = (this.view.dom as HTMLElement).getBoundingClientRect();
    const para = (this.dom.closest(".para") as HTMLElement).getBoundingClientRect();
    const box = this.img.getBoundingClientRect();
    const scale = this.scale();
    this.change({ floating: true, wrap, relX: "margin", relY: "paragraph", x: round((box.left - doc.left) / scale / PX), y: round((box.top - para.top) / scale / PX) });
  }

  update(node: Node): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.draw();
    if (this.bar) this.showBar();
    return true;
  }

  selectNode() {
    this.dom.classList.add("selected");
    this.showBar();
  }

  deselectNode() {
    this.dom.classList.remove("selected");
    this.bar?.remove();
    this.bar = null;
  }

  stopEvent(event: Event) {
    return event.type === "mousedown" && (event.target as HTMLElement).closest(".resize-handle, .picture-bar") !== null;
  }

  ignoreMutation() {
    return true;
  }

  destroy() {
    floating.delete(this);
  }
}

const round = (points: number) => Math.round(points * 100) / 100;

/** Shows the snapping grid over the text area of the pages while a picture moves. */
function showGrid(view: EditorView, ctx: DocContext): HTMLElement {
  const grid = document.createElement("div");
  grid.className = "snap-grid";
  const step = GRID * PX;
  grid.style.backgroundSize = `${step}px ${step}px`;
  const doc = view.dom as HTMLElement;
  grid.style.left = `${doc.offsetLeft}px`;
  grid.style.top = `${doc.offsetTop}px`;
  grid.style.width = `${(ctx.page.width - ctx.page.marginLeft - ctx.page.marginRight) * PX}px`;
  grid.style.height = `${doc.offsetHeight}px`;
  // On the sheet, not in the text being edited.
  doc.parentElement!.append(grid);
  return grid;
}
