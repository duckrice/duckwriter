import "./styles.css";
import type { Node } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { h } from "./dom";
import type { DocContext } from "./docx/model";
import { readDocx } from "./docx/read";
import { blankDocx } from "./docx/template";
import { writeDocx } from "./docx/write";
import { createState, createView } from "./editor/editor";
import { PAGE_GAP, relayout, type PageLayout } from "./editor/pagination";
import { stylesheetFor } from "./editor/stylesheet";
import { baseName, isDesktop, onCloseRequested, pickDocument, pickSavePath, readPath, setTitle, startupFiles, writePath } from "./platform";
import { Footnotes } from "./ui/footnotes";
import { updateTocs } from "./ui/references";
import { HeadersAndFooters } from "./ui/headers";
import { icon } from "./ui/icons";
import { Toolbar } from "./ui/toolbar";

const ZOOMS = [0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];

class App {
  ctx!: DocContext;
  /** The editor of the main text. */
  body!: EditorView;
  private path: string | null = null;
  private name = "Document1.docx";
  /** The document as last opened or saved, to tell whether there are unsaved changes. */
  private savedDoc: Node | null = null;
  private zoom = 1;

  private toolbar: Toolbar;
  private styleEl = h("style");
  /** The page sheets, drawn behind the text. */
  private pagesLayer = h("div", { class: "pages", "aria-hidden": "true" });
  private sheet = h("div", { class: "sheet" }, this.pagesLayer);
  private workspace = h("main", { class: "workspace" }, this.sheet);
  private pageCount = h("span", { class: "status-item" });
  private words = h("span", { class: "status-item" });
  private zoomLabel = h("button", { type: "button", class: "status-button", title: "Reset zoom to 100%" });
  private toasts = h("div", { class: "toasts", role: "status" });
  private wordTimer = 0;
  private headers: HeadersAndFooters;
  private footnotes: Footnotes;
  /** Whether the first page had its own header when last opened or saved. */
  private savedTitlePage = false;

  constructor(root: HTMLElement) {
    this.toolbar = new Toolbar(this);
    this.headers = new HeadersAndFooters(this, this.workspace);
    this.footnotes = new Footnotes(this);
    this.sheet.append(this.headers.layer, this.footnotes.layer);
    // Clicking a footnote's mark in the text opens the note; clicking a contents entry goes to the heading.
    this.sheet.addEventListener("click", (event) => {
      const target = event.target as HTMLElement;
      const mark = target.closest<HTMLElement>(".doc .footnote-ref");
      if (mark?.dataset.id) this.footnotes.start(mark.dataset.id);
      const entry = target.closest<HTMLElement>(".doc .toc-entry");
      if (entry?.dataset.pos) this.goTo(Number(entry.dataset.pos));
    });
    const zoomOut = h("button", { type: "button", class: "status-button", title: "Zoom out", html: icon("zoomOut", 16), onclick: () => this.stepZoom(-1) });
    const zoomIn = h("button", { type: "button", class: "status-button", title: "Zoom in", html: icon("zoomIn", 16), onclick: () => this.stepZoom(1) });
    this.zoomLabel.addEventListener("click", () => this.setZoom(1));
    const status = h("footer", { class: "statusbar" }, this.pageCount, this.words, h("span", { class: "status-spacer" }), zoomOut, this.zoomLabel, zoomIn);
    document.head.append(this.styleEl);
    root.append(this.toolbar.el, this.workspace, status, this.toasts);

    this.workspace.addEventListener("wheel", (event) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      this.stepZoom(event.deltaY < 0 ? 1 : -1);
    }, { passive: false });
    // Clicking the grey area or the margins puts the cursor in the text.
    this.workspace.addEventListener("mousedown", (event) => {
      if (event.target === this.workspace || event.target === this.sheet || this.pagesLayer.contains(event.target as HTMLElement)) {
        event.preventDefault();
        this.view.focus();
      }
    });
    // Double-clicking the text while editing a header goes back to the text.
    this.sheet.addEventListener("dblclick", (event) => {
      if (this.headers.view && !this.headers.layer.contains(event.target as HTMLElement)) this.headers.stop();
    });
    window.addEventListener("keydown", (event) => this.onKey(event));
    onCloseRequested(() => this.confirmDiscard());
    window.addEventListener("beforeunload", (event) => {
      if (!isDesktop && this.dirty) event.preventDefault();
    });
    this.setZoom(1);
  }

  /** The editor being typed in: the main text, or a header or footer. */
  get view(): EditorView {
    return this.headers.view ?? this.footnotes.view ?? this.body;
  }

  get dirty(): boolean {
    if (!this.savedDoc) return false;
    if (!this.body.state.doc.eq(this.savedDoc) || this.ctx.titlePage !== this.savedTitlePage) return true;
    if (this.ctx.stories.some((s) => !s.saved || !s.doc.eq(s.saved))) return true;
    const fn = this.ctx.footnotes;
    return [...fn.notes].some(([id, doc]) => {
      const saved = fn.saved.get(id);
      return !saved || !saved.eq(doc);
    });
  }

  /** Puts the cursor at the start of the paragraph at `pos` and shows it. */
  goTo(pos: number) {
    const view = this.body;
    const at = Math.min(pos + 1, view.state.doc.content.size);
    view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(at))));
    (view.nodeDOM(pos) as HTMLElement | null)?.scrollIntoView({ block: "start", behavior: "smooth" });
    view.focus();
  }

  insertFootnote() {
    this.headers.stop();
    this.footnotes.insert();
  }

  editHeader() {
    this.headers.start("header", 0);
  }

  storiesChanged() {
    this.changed(false);
  }

  /** Shows a document; `path` is where Save writes it (null for a new one). */
  private show(doc: Node, ctx: DocContext, name: string, path: string | null) {
    this.headers.stop();
    this.footnotes.stop();
    this.ctx = ctx;
    this.path = path;
    this.name = name;
    this.savedDoc = doc;
    this.savedTitlePage = ctx.titlePage;
    delete this.styleEl.dataset.key;
    this.refreshStyles();
    const p = ctx.page;
    this.sheet.style.width = `${p.width}pt`;
    const state = createState(doc, () => this.ctx, { onLayout: (layout) => this.drawPages(layout), noteHeight: (id) => this.footnotes.height(id) });
    if (this.body) this.body.updateState(state);
    else this.body = createView(this.sheet, state, { onChange: (_view, docChanged) => this.changed(docChanged) });
    const dom = this.body.dom as HTMLElement;
    dom.style.margin = `${p.marginTop}pt ${p.marginRight}pt 0 ${p.marginLeft}pt`;
    dom.style.minHeight = `${p.height - p.marginTop - p.marginBottom}pt`;
    relayout(this.body);
    this.workspace.scrollTop = 0;
    this.changed(true);
    this.body.focus();
  }

  /** Draws a sheet of paper behind each page of text. */
  private drawPages(layout: PageLayout) {
    const p = this.ctx.page;
    const pitch = p.height * (4 / 3) + PAGE_GAP;
    while (this.pagesLayer.children.length < layout.count) this.pagesLayer.append(h("div", { class: "page-sheet" }));
    while (this.pagesLayer.children.length > layout.count) this.pagesLayer.lastElementChild!.remove();
    [...this.pagesLayer.children].forEach((el, i) => {
      const sheet = el as HTMLElement;
      sheet.style.top = `${i * pitch}px`;
      sheet.style.height = `${p.height}pt`;
    });
    this.sheet.style.height = `${layout.count * pitch - PAGE_GAP}px`;
    this.pageCount.textContent = `${layout.count} ${layout.count === 1 ? "page" : "pages"}`;
    this.headers.render(layout.count);
    this.footnotes.render(layout);
    // After this layout, so the table of contents gets the page numbers just worked out.
    queueMicrotask(() => updateTocs(this.body, this.ctx, layout));
  }

  /** The CSS for the document's styles; again whenever a style was added or changed. */
  private refreshStyles() {
    const key = `${this.ctx.styles.byId.size}|${this.ctx.addedStyles.length}|${this.ctx.styleNumbering.size}`;
    if (this.styleEl.dataset.key === key && this.styleEl.dataset.doc === this.name) return;
    this.styleEl.dataset.key = key;
    this.styleEl.dataset.doc = this.name;
    this.styleEl.textContent = `${stylesheetFor(this.ctx, ".sheet .doc")}\n${stylesheetFor(this.ctx, ".hf-doc")}`;
  }

  private changed(docChanged: boolean) {
    this.refreshStyles();
    this.toolbar.update();
    setTitle(`${this.dirty ? "• " : ""}${this.name} – DuckWriter`);
    if (docChanged) {
      clearTimeout(this.wordTimer);
      this.wordTimer = window.setTimeout(() => this.countWords(), 300);
    }
  }

  private countWords() {
    const doc = this.body.state.doc;
    const text = doc.textBetween(0, doc.content.size, " ", " ");
    const count = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’\-.]*/gu) ?? []).length;
    this.words.textContent = `${count.toLocaleString()} ${count === 1 ? "word" : "words"}`;
  }

  // ---- files ----------------------------------------------------------------

  async newDocument(ask = true) {
    if (ask && !(await this.confirmDiscard())) return;
    const { doc, context } = await readDocx(await blankDocx());
    this.show(doc, context, "Document1.docx", null);
  }

  async open() {
    if (!(await this.confirmDiscard())) return;
    let file;
    try {
      file = await pickDocument();
    } catch (error) {
      return this.toast(`Couldn't open the file: ${message(error)}`, "error");
    }
    if (file) await this.load(file.bytes, file.name, file.path);
  }

  async openPath(path: string) {
    try {
      const file = await readPath(path);
      await this.load(file.bytes, file.name, file.path);
    } catch (error) {
      this.toast(`Couldn't open ${baseName(path)}: ${message(error)}`, "error");
    }
  }

  private async load(bytes: Uint8Array, name: string, path: string | null) {
    if (!/\.docx$/i.test(name)) {
      this.toast(/\.doc$/i.test(name) ? `${name} is an old Word file (.doc). Open it in Word and save it as .docx first.` : `${name} isn't a Word document (.docx).`, "error");
      return;
    }
    try {
      const { doc, context } = await readDocx(bytes);
      this.show(doc, context, name, path);
    } catch (error) {
      this.toast(message(error), "error");
    }
  }

  /** The document as .docx bytes. */
  async build(): Promise<Uint8Array> {
    return writeDocx(this.body.state.doc, this.ctx);
  }

  /** Saves; resolves to false if cancelled or failed. */
  async save(saveAs = false): Promise<boolean> {
    let path = this.path;
    if (!path || saveAs) {
      path = await pickSavePath(this.name);
      if (!path) return false;
    }
    const doc = this.body.state.doc;
    try {
      await writePath(path, await this.build());
    } catch (error) {
      this.toast(`Couldn't save: ${message(error)}`, "error");
      return false;
    }
    if (isDesktop) this.path = path;
    this.name = baseName(path);
    this.savedDoc = doc;
    this.savedTitlePage = this.ctx.titlePage;
    this.changed(false);
    this.toast(`Saved ${this.name}`);
    return true;
  }

  /** Offers to save unsaved changes. Resolves to true when it's fine to continue. */
  async confirmDiscard(): Promise<boolean> {
    if (!this.dirty) return true;
    const answer = await ask("Save your changes?", `"${this.name}" has changes that haven't been saved.`, [
      { label: "Cancel", value: "cancel" },
      { label: "Don't save", value: "discard" },
      { label: "Save", value: "save", primary: true },
    ]);
    if (answer === "save") return this.save();
    return answer === "discard";
  }

  // ---- view -----------------------------------------------------------------

  private setZoom(zoom: number) {
    this.zoom = Math.min(5, Math.max(0.25, zoom));
    this.sheet.style.zoom = String(this.zoom);
    this.zoomLabel.textContent = `${Math.round(this.zoom * 100)}%`;
  }

  private stepZoom(direction: 1 | -1) {
    const next = direction > 0 ? ZOOMS.find((z) => z > this.zoom + 0.001) : [...ZOOMS].reverse().find((z) => z < this.zoom - 0.001);
    if (next) this.setZoom(next);
  }

  private onKey(event: KeyboardEvent) {
    if (event.ctrlKey && event.altKey && event.key.toLowerCase() === "f") {
      // Ctrl+Alt+F: a new footnote, as in Word.
      event.preventDefault();
      this.insertFootnote();
      return;
    }
    if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
    const key = event.key.toLowerCase();
    const handled = () => {
      event.preventDefault();
      event.stopPropagation();
    };
    if (key === "s") {
      handled();
      void this.save(event.shiftKey);
    } else if (key === "o") {
      handled();
      void this.open();
    } else if (key === "n" && !event.shiftKey) {
      handled();
      void this.newDocument();
    } else if (key === "0") {
      handled();
      this.setZoom(1);
    }
  }

  toast(text: string, kind: "info" | "error" = "info") {
    const el = h("div", { class: `toast ${kind}` }, text);
    this.toasts.append(el);
    setTimeout(() => el.remove(), kind === "error" ? 8000 : 3000);
  }
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** Shows a modal question and resolves to the chosen button's value, or null if dismissed. */
function ask(title: string, text: string, buttons: { label: string; value: string; primary?: boolean }[]) {
  return new Promise<string | null>((resolve) => {
    let answer: string | null = null;
    const dialog = h(
      "dialog",
      { class: "dialog" },
      h("h2", {}, title),
      h("p", {}, text),
      h(
        "div",
        { class: "dialog-footer" },
        ...buttons.map((b) =>
          h(
            "button",
            {
              type: "button",
              class: `btn${b.primary ? " primary" : ""}`,
              onclick: () => {
                answer = b.value;
                dialog.close();
              },
            },
            b.label,
          ),
        ),
      ),
    );
    dialog.addEventListener("close", () => {
      dialog.remove();
      resolve(answer);
    });
    document.body.append(dialog);
    dialog.showModal();
    (dialog.querySelector(".btn.primary") as HTMLElement | null)?.focus();
  });
}

const app = new App(document.querySelector<HTMLElement>("#app")!);
const [first] = await startupFiles();
if (first) {
  await app.newDocument(false);
  await app.openPath(first);
} else await app.newDocument(false);

if (localStorage.getItem("duckwriter.debug") === "1") {
  Object.assign(window, { __dw: app });
}
