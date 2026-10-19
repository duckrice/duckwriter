import "./styles.css";
import type { Node } from "prosemirror-model";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { h } from "./dom";
import { resolvedPara, resolvedRun, type DocContext } from "./docx/model";
import { readDocx } from "./docx/read";
import { ensureStyles } from "./docx/builtins";
import { blankDocx } from "./docx/template";
import { writeDocx } from "./docx/write";
import { createState, createView, setLinkOpener } from "./editor/editor";
import { bookmarkPosition } from "./ui/references";
import { addRecent, removeRecent } from "./recent";
import { pageSetup } from "./ui/pagesetup";
import { showVersions, type Version } from "./ui/versions";
import { TEMPLATES, templateDoc, type TemplateKind } from "./templates";
import { NavigationPane } from "./ui/navigation";
import { count, showCounts } from "./ui/wordcount";
import { offerUpdate } from "./updates";
import type { Plugin } from "prosemirror-state";
import type { CollabSession } from "./collab/session";
import { joinSharing, showSharing, startSharing, stopSharing } from "./ui/share";
import type { Method } from "./collab/invite";
import { schema } from "./editor/schema";
import { PAGE_GAP, pageAt, pageBox, pageOffset, relayout, setPageGap, type PageLayout } from "./editor/pagination";
import { invoke } from "@tauri-apps/api/core";
import { stylesheetFor } from "./editor/stylesheet";
import { pickExportPath, setFullscreen } from "./platform";
import { compareDocs, countChanges } from "./editor/compare";
import { newAnnotationId } from "./docx/model";
import { toHtml, toMarkdown, toText } from "./convert/export";
import { toLatex } from "./convert/latex";
import { authorName } from "./author";
import { showProperties } from "./ui/properties";
import { isReading, toggleReadAloud } from "./ui/readaloud";
import { stamped } from "./docx/properties";
import { checkPackage } from "./docx/check";
import { asPackage, convertPackage, isWordFile, packageKind, toFlatOpc, wordXml } from "./docx/package";
import { htmlToDoc, markdownToHtml, textFormat, textToHtml } from "./convert/import";
import { baseName, isDesktop, onCloseRequested, openWindow, pickDocument, pickPdfPath, pickSavePath, readPath, setTitle, startupFiles, windowOpening, windowsSpelling, writePath } from "./platform";
import { Footnotes } from "./ui/footnotes";
import { updateTocs } from "./ui/references";
import { insertPictureFiles, setSnap, snapOn } from "./ui/pictures";
import { handleShortcuts, showShortcutHelp } from "./ui/shortcuts";
import { AutoSave } from "./autosave";
import { FindBar } from "./ui/findbar";
import { mistakeAt, nextMistake, openMenuFor, setSpellChecker, setSpelling, spellingOn } from "./editor/spelling";
import { closeSpellMenu, showSpellMenu } from "./ui/spellmenu";
import { CommentsPane, askName } from "./ui/comments";
import { changes, changesAt, commentAnchors, decide, decideAll, nextChange, reviewKey } from "./editor/review";
import { placeFloatingPictures } from "./editor/imageview";
import { HeadersAndFooters } from "./ui/headers";
import { icon } from "./ui/icons";
import { Toolbar } from "./ui/toolbar";

const ZOOMS = [0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3];

class App {
  ctx!: DocContext;
  /** The editor of the main text. */
  body!: EditorView;
  path: string | null = null;
  name = "Document1.docx";
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
  private words = h("button", { type: "button", class: "status-button words", title: "Word count (Ctrl+Shift+G)" });
  private totalWords = 0;
  private lastLayout: PageLayout | null = null;
  private pages = 1;
  private zoomLabel = h("button", { type: "button", class: "status-button", title: "Reset zoom to 100%" });
  private toasts = h("div", { class: "toasts", role: "status" });
  private saveStatus = h("button", { type: "button", class: "status-button save-status" });
  private spellStatus = h("button", { type: "button", class: "status-button spell-status", hidden: true });
  private spellLanguages: string[] = [];
  private autosave: AutoSave;
  private wordTimer = 0;
  private headers: HeadersAndFooters;
  private findBar: FindBar;
  private comments: CommentsPane;
  /** Writing together, when this document is shared or joined. */
  collab: CollabSession | null = null;
  /** The file as last opened or saved: someone joining starts from it. */
  private base: Uint8Array | null = null;
  private shareStatus = h("button", { type: "button", class: "status-button share-status", hidden: true, title: "Writing together (Alt+Shift+W)" });
  private navigation: NavigationPane;
  private navButton = h("button", { type: "button", class: "status-button nav-button", title: "Headings: the navigation pane (Alt+Shift+N)" }, "☰ Headings");
  private trackStatus = h("button", { type: "button", class: "status-button track-status", hidden: true, title: "Changes are tracked (Ctrl+Shift+E to stop)" }, "Tracking changes");
  /** The page size for printing, as an @page rule. */
  private pageRule = h("style");
  /** Resolves when the next page layout is done. */
  private laidOut: (() => void) | null = null;
  private footnotes: Footnotes;
  /** Whether the first page had its own header when last opened or saved. */
  private savedTitlePage = false;
  private lineGridSaved = true;

  constructor(root: HTMLElement) {
    this.toolbar = new Toolbar(this);
    const app = this;
    this.autosave = new AutoSave({
      get dirty() {
        return app.body ? app.dirty : false;
      },
      get path() {
        return app.path;
      },
      get name() {
        return app.name;
      },
      build: () => this.build(),
      saveQuietly: () => this.saveQuietly(),
      status: (text, title) => {
        this.saveStatus.textContent = text;
        this.saveStatus.title = title;
      },
    });
    this.saveStatus.addEventListener("click", () => this.autosave.toggle());
    this.spellStatus.addEventListener("click", () => this.toggleSpelling());
    if (isDesktop) {
      setSpellChecker(windowsSpelling, showSpellMenu);
      void windowsSpelling.languages().then((languages) => {
        this.spellLanguages = languages;
        this.showSpelling();
      }, () => setSpellChecker(null));
    }
    this.headers = new HeadersAndFooters(this, this.workspace);
    this.findBar = new FindBar(this);
    this.comments = new CommentsPane({
      get body() {
        return app.body;
      },
      get ctx() {
        return app.ctx;
      },
      sheet: this.sheet,
      get zoom() {
        return app.zoom;
      },
      commentsChanged: () => {
        this.changed(false, true);
        this.collab?.commentsChanged();
      },
    });
    this.sheet.append(this.comments.el);
    this.navigation = new NavigationPane({
      get body() {
        return app.body;
      },
      get ctx() {
        return app.ctx;
      },
      goTo: (pos) => this.goTo(pos),
      run: (command) => command(this.body.state, this.body.dispatch, this.body),
    });
    this.navButton.addEventListener("click", () => this.toggleNavigation());
    this.navButton.classList.toggle("active", this.navigation.open);
    this.trackStatus.addEventListener("click", () => this.toggleTracking());
    this.shareStatus.addEventListener("click", () => this.share());
    this.workspace.addEventListener("duckwriter-sources", () => this.collab?.sourcesChanged());
    setLinkOpener((href, anchor) => {
      if (href) {
        if (isDesktop) invoke("open_url", { url: href }).catch((error) => this.toast(message(error), "error"));
        else window.open(href, "_blank", "noopener");
        return;
      }
      const pos = anchor ? bookmarkPosition(this.body.state.doc, anchor) : null;
      if (pos === null) return this.toast("The place this link goes to isn't in the document");
      this.goTo(this.body.state.doc.resolve(pos).before());
    });
    window.addEventListener("resize", () => this.placeSheet());
    try {
      if (localStorage.getItem("duckwriter.darkPage") === "on") this.sheet.classList.add("dark-page");
    } catch {
      // Light, then.
    }
    // Esc leaves focus mode (when nothing else wants it: a menu, a dialog, the find bar).
    window.addEventListener(
      "keydown",
      (event) => {
        if (event.key !== "Escape" || document.querySelector("dialog[open], .menu") || (event.target as HTMLElement).closest?.(".find-bar, .comments-pane")) return;
        if (this.leaveFocus()) {
          event.preventDefault();
          event.stopPropagation();
        }
      },
      true,
    );
    document.addEventListener("duckwriter-toast", (event) => this.toast((event as CustomEvent<string>).detail));
    // A document dropped on the window opens (Word, Markdown, a web page, text); pictures dropped in the text are the editor's.
    const documentFile = (f: File) => /\.(docx?|docm|dotx|dotm|xml|md|markdown|html?|txt)$/i.test(f.name);
    const hasFiles = (event: DragEvent) => [...(event.dataTransfer?.items ?? [])].some((i) => i.kind === "file");
    window.addEventListener("dragover", (event) => {
      if (!hasFiles(event)) return;
      event.preventDefault();
      if (event.dataTransfer) event.dataTransfer.dropEffect = "copy";
    });
    window.addEventListener("drop", (event) => {
      const file = [...(event.dataTransfer?.files ?? [])].find(documentFile);
      if (!file) return;
      event.preventDefault();
      event.stopPropagation();
      void this.openDropped(file);
    }, true);
    // Anything else dropped outside the text mustn't open in the window in place of the document.
    window.addEventListener("drop", (event) => {
      if (event.defaultPrevented || !event.dataTransfer?.files.length) return;
      event.preventDefault();
      this.toast("Drop documents on the window to open them, and pictures into the text");
    });
    this.workspace.prepend(this.findBar.el);
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
    this.zoomLabel.addEventListener("click", () => this.cycleZoomFit());
    this.zoomLabel.title = "100% → page width → whole page (Alt+Shift+Z)";
    this.words.addEventListener("click", () => this.wordCount());
    const status = h("footer", { class: "statusbar" }, this.navButton, this.pageCount, this.words, h("span", { class: "status-spacer" }), this.shareStatus, this.trackStatus, this.spellStatus, this.saveStatus, zoomOut, this.zoomLabel, zoomIn);
    document.head.append(this.styleEl, this.pageRule);
    root.append(this.toolbar.el, h("div", { class: "main-row" }, this.navigation.el, this.workspace), status, this.toasts);

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
    handleShortcuts(this);
    onCloseRequested(async () => {
      const close = await this.confirmDiscard();
      if (close) this.leaveSharing();
      // Closed on purpose: the recovery copy isn't needed.
      if (close) await this.autosave.clearRecovery();
      return close;
    });
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
    if (!this.body.state.doc.eq(this.savedDoc) || this.ctx.titlePage !== this.savedTitlePage || !this.lineGridSaved) return true;
    if ([...this.ctx.stories, ...(this.ctx.sectionStories?.values() ?? [])].some((s) => !s.saved || !s.doc.eq(s.saved))) return true;
    if (this.ctx.comments.version !== this.ctx.comments.savedVersion || this.ctx.trackChanges !== this.ctx.trackChangesSaved) return true;
    if (this.ctx.bibliography.version !== this.ctx.bibliography.savedVersion) return true;
    if (this.ctx.styleChanges.size || JSON.stringify(this.ctx.properties) !== this.savedProperties) return true;
    const en = this.ctx.endnotes;
    if ([...en.notes].some(([id, doc]) => !en.saved.get(id)?.eq(doc))) return true;
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

  /** The page's size and margins, on screen. */
  private applyPage() {
    const p = this.ctx.page;
    this.sheet.style.width = `${p.width}pt`;
    const dom = this.body.dom as HTMLElement;
    dom.style.margin = `${p.marginTop}pt ${p.marginRight}pt 0 ${p.marginLeft}pt`;
    dom.style.minHeight = `${p.height - p.marginTop - p.marginBottom}pt`;
    this.pageRule.textContent = "";
    relayout(this.body);
  }

  // ---- markup, pages ------------------------------------------------------------------

  /** Tracked changes shown with all markup, as they would be once accepted (final), or as they were (original). */
  markup: "all" | "final" | "original" = "all";

  setMarkup(mode: "all" | "final" | "original") {
    this.markup = mode;
    this.sheet.dataset.markup = mode;
    relayout(this.body);
    this.toast(mode === "all" ? "Showing all changes" : mode === "final" ? "Showing the text as it would be with all changes accepted" : "Showing the text as it was before the changes");
    this.toolbar.update();
  }

  /** All markup → final → original (Alt+Shift+M). */
  cycleMarkup() {
    this.setMarkup(this.markup === "all" ? "final" : this.markup === "final" ? "original" : "all");
  }

  /** Asks which page, and goes there (Ctrl+G). */
  goToPage() {
    const count = this.pages;
    const input = h("input", { class: "tb-input dialog-input number", type: "number", min: "1", max: String(count), value: "", "aria-label": "Page" }) as HTMLInputElement;
    const go = () => {
      const n = Math.max(1, Math.min(count, Math.round(Number(input.value) || 1)));
      dialog.close();
      const layout = this.lastLayout;
      let target: number | null = null;
      for (const [pos, page] of layout?.pageOfBlock ?? []) if (page >= n - 1 && (target === null || pos < target)) target = pos;
      if (target !== null) this.goTo(target);
    };
    const dialog = h(
      "dialog",
      { class: "dialog" },
      h("h2", {}, "Go to page"),
      h("div", { class: "dialog-form" }, h("label", {}, `Page (1–${count})`, input)),
      h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"), h("button", { type: "button", class: "btn primary", onclick: go }, "Go")),
    ) as HTMLDialogElement;
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        go();
      }
    });
    dialog.addEventListener("close", () => {
      dialog.remove();
      this.view.focus();
    });
    document.body.append(dialog);
    dialog.showModal();
    input.focus();
  }

  /** Page numbers in the footer ("Page 2 of 9", centred) or at the top or bottom right. */
  pageNumbers(where: "bottom-center" | "bottom-right" | "top-right") {
    const kind = where.startsWith("top") ? "header" : "footer";
    const style = kind === "header" ? "Header" : "Footer";
    ensureStyles(this.ctx, style);
    const field = (instr: string, result: string) => schema.nodes.field.create({ instr, result });
    const content = where === "bottom-center" ? [schema.text("Page "), field(" PAGE ", "1"), schema.text(" of "), field(" NUMPAGES ", String(this.pages))] : [field(" PAGE ", "1")];
    const para = schema.nodes.paragraph.create({ style: this.ctx.styles.byId.has(style) ? style : null, align: where === "bottom-center" ? "center" : "right" }, content);
    let story = this.ctx.stories.find((s) => s.kind === kind && s.type === "default");
    if (!story) {
      story = { kind, type: "default", path: null, rid: null, rootStart: null, prefixes: new Set(), rels: new Map(), doc: schema.nodes.doc.create(null, para), saved: null };
      this.ctx.stories.push(story);
    } else {
      // Page numbers already there are replaced; anything else written there stays.
      const kept: Node[] = [];
      story.doc.forEach((p) => {
        let hasPage = false;
        p.descendants((n) => {
          if (n.type === schema.nodes.field && /^\s*(PAGE|NUMPAGES)\b/.test(n.attrs.instr)) hasPage = true;
        });
        if (!hasPage && (p.textContent.trim() || p.childCount)) kept.push(p);
      });
      story.doc = schema.nodes.doc.create(null, [...kept, para]);
    }
    relayout(this.body);
    this.changed(false, true);
    this.toast("Page numbers added");
  }

  private lineNumberLayer = h("div", { class: "line-numbers", "aria-hidden": "true" });

  /** Numbers the lines in the left margin, every one (or every n-th), from 1 or from 1 on each page. */
  private drawLineNumbers() {
    const numbering = this.ctx.page.lineNumbers;
    if (!this.lineNumberLayer.isConnected) this.sheet.append(this.lineNumberLayer);
    if (!numbering) return void this.lineNumberLayer.replaceChildren();
    const zoom = this.zoom || 1;
    const origin = this.sheet.getBoundingClientRect();
    const lines: number[] = [];
    for (const p of this.body.dom.querySelectorAll<HTMLElement>(":scope > p.para")) {
      const range = document.createRange();
      range.selectNodeContents(p);
      const rects = [...range.getClientRects()].filter((r) => r.height >= 2 && r.width > 0).sort((a, b) => a.top - b.top);
      // An empty paragraph is one line.
      if (!rects.length) {
        const r = p.getBoundingClientRect();
        lines.push((r.top - origin.top) / zoom + (r.height / zoom - parseFloat(getComputedStyle(p).paddingBottom || "0")) / 2);
        continue;
      }
      let last = -Infinity;
      for (const r of rects) {
        const mid = (r.top + r.bottom) / 2;
        // Pieces of the same line (formatting, a page gap) overlap.
        if (mid - last < 4) continue;
        last = r.bottom;
        lines.push((mid - origin.top) / zoom);
      }
    }
    const items: HTMLElement[] = [];
    let n = 0;
    let page = -1;
    for (const y of lines) {
      const k = pageAt(this.lastLayout, y, this.ctx.page);
      if (numbering.restart === "newPage" && k !== page) n = 0;
      page = k;
      n++;
      if (n % numbering.countBy) continue;
      const label = h("span", { class: "line-number" }, String(n));
      label.style.top = `${y}px`;
      items.push(label);
    }
    this.lineNumberLayer.style.width = `${this.ctx.page.marginLeft - 14}pt`;
    this.lineNumberLayer.replaceChildren(...items);
  }

  /** Line numbers off → every line, continuous → from 1 on each page → off (Alt+Shift+L). */
  cycleLineNumbers() {
    const now = this.ctx.page.lineNumbers;
    this.setLineNumbers(!now ? "continuous" : now.restart === "continuous" ? "newPage" : null);
  }

  setLineNumbers(restart: "continuous" | "newPage" | null) {
    this.ctx.page = { ...this.ctx.page, lineNumbers: restart ? { countBy: this.ctx.page.lineNumbers?.countBy ?? 1, restart } : null };
    this.ctx.pageChanged = true;
    this.lineGridSaved = false;
    this.drawLineNumbers();
    this.changed(false, true);
    this.toast(restart === "continuous" ? "Lines are numbered through the document" : restart === "newPage" ? "Lines are numbered from 1 on each page" : "No line numbers");
  }

  /** A new document from a template: here if this window is unused, else in a window of its own. */
  async newFromTemplate(kind: TemplateKind) {
    if (isDesktop && !this.blank) return void openWindow(`template:${kind}`);
    if (!(await this.confirmDiscard())) return;
    this.leaveSharing();
    const bytes = await blankDocx();
    const { context } = await readDocx(bytes);
    const doc = templateDoc(kind, context);
    this.base = bytes;
    this.show(doc, context, TEMPLATES.find((t) => t.kind === kind)?.file ?? "Document1.docx", null);
  }

  /** Markdown, a web page or text, opened as a new Word document (saving asks where). */
  private async importText(bytes: Uint8Array, name: string, format: "markdown" | "html" | "text") {
    try {
      const text = new TextDecoder().decode(bytes);
      const blank = await blankDocx();
      const { context } = await readDocx(blank);
      const html = format === "markdown" ? markdownToHtml(text) : format === "html" ? text : textToHtml(text);
      const doc = htmlToDoc(html, context);
      this.leaveSharing();
      this.base = blank;
      this.show(doc, context, name.replace(/\.[^.]+$/, ".docx"), null);
      // Not saved as a Word document yet.
      this.savedDoc = schema.nodes.doc.create(null, schema.nodes.paragraph.create());
      this.changed(true);
      this.toast(`Opened ${name} as a new Word document`);
    } catch (error) {
      this.toast(`Couldn't open ${name}: ${message(error)}`, "error");
    }
  }

  /** Saves a copy as Markdown, a web page or text; the document stays as it is. */
  async exportAs(format: "markdown" | "html" | "text" | "latex") {
    const info = { markdown: { ext: "md", name: "Markdown" }, html: { ext: "html", name: "Web page" }, text: { ext: "txt", name: "Plain text" }, latex: { ext: "tex", name: "LaTeX" } }[format];
    const base = this.name.replace(/\.docx$/i, "");
    const path = await pickExportPath(`${base}.${info.ext}`, info.name, info.ext);
    if (path) await this.exportTo(format, path);
  }

  /** Writes the copy (exportAs, once it knows where; tests call it directly). */
  async exportTo(format: "markdown" | "html" | "text" | "latex", path: string) {
    const base = this.name.replace(/\.docx$/i, "");
    try {
      const doc = this.body.state.doc;
      if (format === "latex") return await this.exportLatex(path);
      const content = format === "markdown" ? await toMarkdown(doc, this.ctx) : format === "html" ? await toHtml(doc, this.ctx, base) : toText(doc, this.ctx);
      // Text files start with a byte-order mark so Windows programs read æ, ø and å right.
      const bytes = new TextEncoder().encode(format === "text" ? `\ufeff${content}` : content);
      await writePath(path, bytes);
      this.toast(`Saved ${baseName(path)}`);
    } catch (error) {
      this.toast(`Couldn't save it: ${message(error)}`, "error");
    }
  }

  /** The .tex file, with the sources (.bib) and pictures beside it, named after it. */
  private async exportLatex(path: string) {
    const dir = path.slice(0, Math.max(path.lastIndexOf(String.fromCharCode(92)), path.lastIndexOf("/")) + 1);
    const name = baseName(path).replace(/[.]tex$/i, "");
    // LaTeX can't take spaces in the names of the files it reads.
    const base = [...name].map((c) => (c === "_" || c === "-" || (c >= "0" && c <= "9") || c.toLowerCase() !== c.toUpperCase() ? c : "-")).join("");
    const { tex, bib, pictures } = await toLatex(this.body.state.doc, this.ctx, base);
    const encode = (s: string) => new TextEncoder().encode(s);
    await writePath(path, encode(tex));
    if (bib) await writePath(`${dir}${base}.bib`, encode(bib));
    for (const p of pictures) await writePath(dir + p.name, p.bytes);
    const extra = [bib ? "the sources" : "", pictures.length ? `${pictures.length} picture${pictures.length === 1 ? "" : "s"}` : ""].filter(Boolean).join(" and ");
    this.toast(`Saved ${baseName(path)}${extra ? `, with ${extra} beside it` : ""}`);
  }

  /** Compares this document with another version of it, chosen from a file (Ctrl+Alt+Shift+C). */
  async compareWithFile() {
    let file;
    try {
      file = await pickDocument();
    } catch (error) {
      return this.toast(`Couldn't open it: ${message(error)}`, "error");
    }
    if (file) await this.compareWith(file.bytes, file.name);
  }

  /** Compares an earlier version (kept on this computer) with the document as it is. */
  async compareVersion(version: Version) {
    try {
      const file = await readPath(version.path);
      await this.compareWith(file.bytes, `the version from ${new Date(version.time).toLocaleString([], { dateStyle: "medium", timeStyle: "short" })}`);
    } catch (error) {
      this.toast(`Couldn't compare: ${message(error)}`, "error");
    }
  }

  /** Shows this document with what changed since `older` as tracked changes. */
  async compareWith(older: Uint8Array, label: string) {
    let original: Node;
    try {
      original = (await readDocx(await asPackage(older))).doc;
    } catch (error) {
      return this.toast(message(error), "error");
    }
    const result = compareDocs(original, this.body.state.doc, { author: "Comparison", newId: () => newAnnotationId(this.ctx) });
    const count = countChanges(result);
    if (!count) return this.toast(`No differences from ${label}`);
    if (!(await this.confirmDiscard())) return;
    const ctx = this.ctx;
    const name = this.name.replace(/\.docx$/i, "");
    this.show(result, ctx, `${name} (compared).docx`, null);
    this.savedDoc = schema.nodes.doc.create(null, schema.nodes.paragraph.create());
    this.changed(true);
    this.toast(`${count} ${count === 1 ? "difference" : "differences"} from ${label}, shown as tracked changes`);
  }

  // ---- the user's own templates -------------------------------------------------------

  private async templatePath(name: string): Promise<string> {
    return `${await invoke<string>("templates_dir")}\\${name}.docx`;
  }

  /** Keeps this document as a template of one's own, to start new documents from. */
  async saveAsTemplate() {
    if (!isDesktop) return this.toast("Templates of your own need the DuckWriter app.");
    const input = h("input", { class: "tb-input dialog-input wide", value: this.name.replace(/\.docx$/i, ""), "aria-label": "Template name" }) as HTMLInputElement;
    const ok = async () => {
      const name = input.value.trim().replace(/[\\/:*?"<>|]/g, "");
      if (!name) return input.focus();
      dialog.close();
      try {
        const path = await this.templatePath(name);
        // Written as it is now, without making this document count as saved.
        const unsaved = this.dirty;
        const bytes = await this.build();
        if (unsaved) this.lineGridSaved = false;
        await writePath(path, bytes);
        this.toast(`Saved as the template "${name}" (New ▸ My templates)`);
      } catch (error) {
        this.toast(`Couldn't save the template: ${message(error)}`, "error");
      }
    };
    const dialog = h(
      "dialog",
      { class: "dialog" },
      h("h2", {}, "Save as a template"),
      h("p", { class: "dialog-hint" }, "New documents can then start from this one: its text, styles, headers and page setup. Kept on this computer."),
      h("div", { class: "dialog-form" }, h("label", {}, "Name", input)),
      h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"), h("button", { type: "button", class: "btn primary", onclick: () => void ok() }, "Save")),
    ) as HTMLDialogElement;
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        void ok();
      }
    });
    dialog.addEventListener("close", () => dialog.remove());
    document.body.append(dialog);
    dialog.showModal();
    input.select();
  }

  /** A new document from one's own template: here if this window is unused, else in a window of its own. */
  async newFromMyTemplate(name: string) {
    if (isDesktop && !this.blank) return void openWindow(`mytemplate:${name}`);
    if (!(await this.confirmDiscard())) return;
    try {
      const file = await readPath(await this.templatePath(name));
      const { doc, context } = await readDocx(file.bytes);
      this.leaveSharing();
      this.base = file.bytes;
      this.show(doc, context, "Document1.docx", null);
    } catch (error) {
      this.toast(`Couldn't open the template: ${message(error)}`, "error");
    }
  }

  /** Earlier versions of this document (Alt+Shift+V). */
  versions() {
    void showVersions(this);
  }

  /** Reads the selection, or from the cursor on, aloud; again to stop (Ctrl+Alt+Space). */
  readAloud() {
    const started = toggleReadAloud(this.body, (on) => {
      if (!on) this.toast("Stopped reading aloud");
    });
    if (started) this.toast("Reading aloud; Ctrl+Alt+Space stops");
    else if (!isReading() && !window.speechSynthesis) this.toast("Windows has no voice to read with", "error");
  }

  /** Title, author, keywords… (Alt+Shift+P). */
  properties() {
    showProperties(
      this.ctx,
      () => this.changed(false, true),
      () => this.body.focus(),
    );
  }

  /** The properties as last saved (they aren't in the text), to tell whether they were edited. */
  private savedProperties = "";

  /** Shows an earlier version as the document; saving keeps it, closing without saving goes back. */
  async restoreVersion(version: Version) {
    if (!(await this.confirmDiscard())) return;
    try {
      const file = await readPath(version.path);
      const bytes = await asPackage(file.bytes, file.name);
      const { doc, context } = await readDocx(bytes);
      this.show(doc, context, this.name, this.path);
      this.base = bytes;
      // Not saved yet: it differs from the file.
      this.savedDoc = schema.nodes.doc.create(null, schema.nodes.paragraph.create());
      this.changed(true);
      this.toast(`Restored the version from ${new Date(version.time).toLocaleString([], { dateStyle: "medium", timeStyle: "short" })}. Save to keep it.`);
    } catch (error) {
      this.toast(`Couldn't restore it: ${message(error)}`, "error");
    }
  }

  /** A style changed: the page shows it at once, and laid out again. */
  styleChanged(id: string) {
    this.refreshStyles();
    relayout(this.body);
    this.changed(false, true);
    const name = this.ctx.styles.byId.get(id)?.name ?? id;
    this.toast(`Style "${name}" changed`);
  }

  sortRows() {
    this.toolbar.sortRows();
  }

  selectTab(n: number) {
    this.toolbar.selectTab(n);
  }

  // ---- focus mode, dark page ---------------------------------------------------------

  /** Only the page, full screen (F11; Esc or F11 to leave). */
  toggleFocus() {
    const on = !document.documentElement.classList.contains("focus-mode");
    document.documentElement.classList.toggle("focus-mode", on);
    void setFullscreen(on);
    this.placeSheet();
    if (on) this.toast("Focus mode: Esc or F11 to leave");
    this.view.focus();
  }

  leaveFocus(): boolean {
    if (!document.documentElement.classList.contains("focus-mode")) return false;
    this.toggleFocus();
    return true;
  }

  /** The page dark, for writing at night (Alt+Shift+D). Only on screen: printing and the file are as they were. */
  toggleDarkPage(on = !this.sheet.classList.contains("dark-page")) {
    this.sheet.classList.toggle("dark-page", on);
    try {
      localStorage.setItem("duckwriter.darkPage", on ? "on" : "off");
    } catch {
      // For now only.
    }
  }

  /** Shows or hides the headings beside the page (Alt+Shift+N). */
  toggleNavigation() {
    this.navigation.toggle();
    this.navButton.classList.toggle("active", this.navigation.open);
    this.placeSheet();
  }

  /** Paper size, orientation and margins (Alt+P). */
  pageSetup() {
    pageSetup(this.ctx.page, (page) => {
      this.ctx.page = page;
      this.ctx.pageChanged = true;
      this.lineGridSaved = false;
      this.applyPage();
      this.headers.stop();
      this.changed(false, true);
      this.view.focus();
    });
  }

  /** Lines of text snap to an evenly spaced grid (Word's document grid), or stop doing so. */
  setLineGrid(on: boolean) {
    const run = resolvedRun(this.ctx.styles, null);
    const para = resolvedPara(this.ctx.styles, null);
    // One step: a normal line of the document's normal text, to the half point.
    const natural = (run.size ?? 11) * 1.17 * (para.lineSpacing ?? 1);
    this.ctx.page.lineGrid = on ? Math.round(natural * 2) / 2 : null;
    this.ctx.lineGridChanged = true;
    delete this.styleEl.dataset.key;
    this.refreshStyles();
    relayout(this.body);
    this.lineGridSaved = false;
    this.changed(false, true);
  }

  insertFootnote() {
    this.headers.stop();
    this.footnotes.insert();
  }

  editHeader() {
    this.headers.start("header", 0);
  }

  storiesChanged() {
    this.changed(false, true);
  }

  /** Shows a document; `path` is where Save writes it (null for a new one). */
  private show(doc: Node, ctx: DocContext, name: string, path: string | null, plugins: Plugin[] = []) {
    this.headers.stop();
    this.footnotes.stop();
    this.ctx = ctx;
    this.path = path;
    this.name = name;
    this.savedDoc = doc;
    this.savedTitlePage = ctx.titlePage;
    this.lineGridSaved = true;
    this.savedProperties = JSON.stringify(ctx.properties);
    delete this.styleEl.dataset.key;
    this.refreshStyles();
    this.sheet.style.width = `${ctx.page.width}pt`;
    const state = createState(doc, () => this.ctx, { onLayout: (layout) => this.drawPages(layout), noteHeight: (id) => this.footnotes.height(id), measureNotes: (ids) => this.footnotes.measure(ids) }, plugins);
    if (this.body) this.body.updateState(state);
    else {
      this.body = createView(this.sheet, state, {
        onChange: (_view, docChanged) => this.changed(docChanged),
        pictures: { context: () => this.ctx, snap: snapOn },
        context: () => this.ctx,
        onImageFiles: (files, view, pos) => void insertPictureFiles(files, view, this.ctx, pos).catch((error) => this.toast(message(error), "error")),
      });
    }
    this.applyPage();
    this.workspace.scrollTop = 0;
    this.changed(true);
    this.body.focus();
  }

  /** Draws a sheet of paper behind each page of text. */
  private drawPages(layout: PageLayout) {
    const p = this.ctx.page;
    while (this.pagesLayer.children.length < layout.count) this.pagesLayer.append(h("div", { class: "page-sheet" }));
    while (this.pagesLayer.children.length > layout.count) this.pagesLayer.lastElementChild!.remove();
    let height = 0;
    [...this.pagesLayer.children].forEach((el, i) => {
      const sheet = el as HTMLElement;
      // A section may have pages of its own size (a landscape table).
      const { top, setup } = pageBox(layout, i, p);
      sheet.style.top = `${top}px`;
      sheet.style.height = `${setup.height}pt`;
      sheet.style.width = setup.width === p.width ? "" : `${setup.width}pt`;
      sheet.style.left = setup.width === p.width ? "" : `${pageOffset(setup, p)}pt`;
      height += setup.height;
    });
    const last = pageBox(layout, layout.count - 1, p);
    // On paper the pages follow each other exactly; a hair less keeps a blank page from spilling over.
    this.sheet.style.height = PAGE_GAP ? `${last.top + last.setup.height * (4 / 3)}px` : `calc(${height}pt - 1px)`;
    this.showPage();
    this.pages = layout.count;
    this.lastLayout = layout;
    this.headers.render(layout.count, layout.labels, layout.sections, layout);
    this.footnotes.render(layout);
    this.comments.layout();
    this.drawLineNumbers();
    this.laidOut?.();
    this.laidOut = null;
    // After this layout, so the table of contents gets the page numbers just worked out.
    queueMicrotask(() => {
      updateTocs(this.body, this.ctx, layout);
      placeFloatingPictures();
    });
  }

  /** The CSS for the document's styles; again whenever a style was added or changed. */
  private refreshStyles() {
    const key = `${this.ctx.styles.byId.size}|${this.ctx.addedStyles.length}|${this.ctx.styleNumbering.size}|${this.ctx.styleVersion ?? 0}`;
    if (this.styleEl.dataset.key === key && this.styleEl.dataset.doc === this.name) return;
    this.styleEl.dataset.key = key;
    this.styleEl.dataset.doc = this.name;
    this.styleEl.textContent = `${stylesheetFor(this.ctx, ".sheet .doc")}\n${stylesheetFor(this.ctx, ".hf-doc")}`;
  }

  /** @param contentChanged something that is saved changed (the text, a header, a note, the grid). */
  private changed(docChanged: boolean, contentChanged = docChanged) {
    if (contentChanged) this.autosave?.changed();
    this.findBar?.update();
    this.comments?.render();
    if (this.collab?.active) this.collab.afterChange();
    if (this.body) this.showPage();
    this.navigation?.update(docChanged);
    this.placeSheet();
    this.trackStatus.hidden = !this.ctx?.trackChanges;
    this.refreshStyles();
    this.toolbar.update();
    setTitle(`${this.dirty ? "• " : ""}${this.name} – DuckWriter`);
    if (docChanged) {
      clearTimeout(this.wordTimer);
      this.wordTimer = window.setTimeout(() => this.countWords(), 300);
    } else this.showWords();
  }

  /** The word count, with the selection's, and the goal (Ctrl+Shift+G). */
  wordCount() {
    const { selection, doc } = this.body.state;
    showCounts(doc, selection, this.pages, { goal: this.wordGoal, today: this.wordsToday() }, (goal) => {
      this.wordGoal = goal;
      this.showWords();
    });
  }

  /** Remembered on this computer, for each document. */
  private get progressKey(): string {
    return this.path ?? `unsaved:${this.name}`;
  }

  /** The number of words this document should reach, if one was set. */
  get wordGoal(): number | null {
    try {
      const goal = Number(localStorage.getItem(`duckwriter.goal:${this.progressKey}`));
      return goal > 0 ? goal : null;
    } catch {
      return null;
    }
  }

  set wordGoal(goal: number | null) {
    try {
      if (goal && goal > 0) localStorage.setItem(`duckwriter.goal:${this.progressKey}`, String(Math.round(goal)));
      else localStorage.removeItem(`duckwriter.goal:${this.progressKey}`);
    } catch {
      // Not remembered.
    }
  }

  /** Words written today: the count now, less the count when the document was first opened today. */
  private wordsToday(): number {
    try {
      const was = JSON.parse(localStorage.getItem(`duckwriter.today:${this.progressKey}`) ?? "null") as { day: string; start: number } | null;
      return was?.day === new Date().toLocaleDateString("sv") ? this.totalWords - was.start : 0;
    } catch {
      return 0;
    }
  }

  /** The first count of the day is where today's words start from. */
  private startToday() {
    const day = new Date().toLocaleDateString("sv");
    const key = `duckwriter.today:${this.progressKey}`;
    try {
      const was = JSON.parse(localStorage.getItem(key) ?? "null") as { day: string } | null;
      if (was?.day !== day) localStorage.setItem(key, JSON.stringify({ day, start: this.totalWords }));
    } catch {
      // Not remembered.
    }
  }

  /** "340 words", "12 of 340 words" with text selected; with a goal, how far along, and today's words. */
  private showWords() {
    const { selection, doc } = this.body.state;
    const words = (n: number) => `${n.toLocaleString()} ${n === 1 ? "word" : "words"}`;
    const goal = this.wordGoal;
    const total = goal ? `${this.totalWords.toLocaleString()} of ${words(goal)}` : words(this.totalWords);
    const today = this.wordsToday();
    const progress = today > 0 ? ` · ${today.toLocaleString()} today` : "";
    this.words.textContent = selection.empty ? `${total}${progress}` : `${count(doc, selection.from, selection.to).words.toLocaleString()} of ${total}`;
    const share = goal ? Math.min(1, this.totalWords / goal) : 0;
    this.words.classList.toggle("has-goal", !!goal);
    this.words.classList.toggle("goal-reached", !!goal && share >= 1);
    this.words.style.setProperty("--goal", `${Math.round(share * 100)}%`);
    this.words.title = goal
      ? `${Math.round(share * 100)}% of the goal of ${words(goal)}${today > 0 ? `; ${words(today)} written today` : ""}. Ctrl+Shift+G for the word count and the goal.`
      : "Ctrl+Shift+G for the word count, and to set a goal";
  }

  private countWords() {
    this.totalWords = count(this.body.state.doc).words;
    this.startToday();
    this.showWords();
  }

  // ---- files ----------------------------------------------------------------

  /** An empty document nobody has typed in: the next one opened can take its place. */
  get blank(): boolean {
    return !this.path && !this.dirty && this.body.state.doc.textContent === "" && this.body.state.doc.childCount === 1;
  }

  // ---- writing together -----------------------------------------------------------

  /** Write together (Alt+Shift+W): share, join, or see who is here. */
  share() {
    showSharing(this);
  }

  /** As the dialog does (and for tests): share by `method`, or join from an invitation. */
  shareBy(method: Method, place: string): Promise<string> {
    return startSharing(this, method, place);
  }

  joinBy(invite: string, folder = ""): Promise<void> {
    return joinSharing(this, invite, folder);
  }

  baseBytes(): Uint8Array {
    return this.base ?? new Uint8Array();
  }

  get folder(): string | null {
    return this.path ? this.path.slice(0, Math.max(this.path.lastIndexOf("\\"), this.path.lastIndexOf("/"))) : null;
  }

  openShared(doc: Node, ctx: DocContext, name: string, plugins: Plugin[]) {
    this.show(doc, ctx, `${name.replace(/\.docx$/i, "")} (shared).docx`, null, plugins);
  }

  setPlugins(add: Plugin[], remove: Plugin[]) {
    const state = this.body.state;
    this.body.updateState(state.reconfigure({ plugins: [...state.plugins.filter((p) => !remove.includes(p)), ...add] }));
  }

  sharingChanged() {
    const session = this.collab?.active ? this.collab : null;
    this.shareStatus.hidden = !session;
    if (session) this.shareStatus.textContent = `👥 ${session.role === "host" ? "Sharing" : "Shared"} · ${session.people.size}`;
    this.toolbar.update();
  }

  commentsUpdated() {
    this.comments.render();
    this.changed(false, true);
  }

  sideUpdated() {
    // Laying the pages out again draws the headers, footers and footnotes anew.
    relayout(this.body);
    this.changed(false, true);
  }

  /** Opening another document here ends writing together on this one. */
  private leaveSharing() {
    if (this.collab?.active) stopSharing(this);
  }

  async newDocument(ask = true) {
    // Another document gets a window of its own, as in Word.
    if (ask && isDesktop && !this.blank) return void openWindow(null).catch((error) => this.toast(message(error), "error"));
    if (ask && !(await this.confirmDiscard())) return;
    this.leaveSharing();
    const bytes = await blankDocx();
    const { doc, context } = await readDocx(bytes);
    this.base = bytes;
    this.show(doc, context, "Document1.docx", null);
  }

  async open() {
    const here = !isDesktop || this.blank;
    if (here && !(await this.confirmDiscard())) return;
    let file;
    try {
      file = await pickDocument();
    } catch (error) {
      return this.toast(`Couldn't open the file: ${message(error)}`, "error");
    }
    if (!file) return;
    if (!here && file.path) return void openWindow(file.path);
    await this.load(file.bytes, file.name, file.path);
  }

  async openPath(path: string) {
    try {
      const file = await readPath(path);
      await this.load(file.bytes, file.name, file.path);
    } catch (error) {
      this.toast(`Couldn't open ${baseName(path)}: ${message(error)}`, "error");
    }
  }

  /** Opens a document from the recent list; one that is gone is taken off it. */
  async openRecent(path: string) {
    if (isDesktop && !this.blank) return void openWindow(path);
    if (!(await this.confirmDiscard())) return;
    try {
      const file = await readPath(path);
      await this.load(file.bytes, file.name, file.path);
    } catch {
      removeRecent(path);
      this.toast(`${baseName(path)} isn't there any more (moved, renamed or deleted)`, "error");
    }
  }

  /** A document dropped on the window. The browser doesn't say where it is, so Save asks. */
  private async openDropped(file: File) {
    if (!(await this.confirmDiscard())) return;
    await this.load(new Uint8Array(await file.arrayBuffer()), file.name, null);
  }

  private async load(bytes: Uint8Array, name: string, path: string | null) {
    const format = textFormat(name);
    if (format) return this.importText(bytes, name, format);
    if (!isWordFile(name)) {
      this.toast(/\.doc$/i.test(name) ? `${name} is an old Word file (.doc). Open it in Word and save it as .docx first.` : `${name} isn't a Word document (.docx).`, "error");
      return;
    }
    try {
      // A Word XML document holds the same parts as a .docx, in one file.
      if (/\.xml$/i.test(name)) bytes = await wordXml(bytes, name);
      const { doc, context } = await readDocx(bytes);
      this.leaveSharing();
      this.base = bytes;
      this.show(doc, context, name, path);
      if (path) addRecent(path);
    } catch (error) {
      this.toast(message(error), "error");
    }
  }

  /** The document as .docx bytes. */
  async build(): Promise<Uint8Array> {
    // "Last saved by" in the properties: the name your comments and changes carry, if given.
    this.ctx.savedBy = authorName() || null;
    const bytes = await writeDocx(this.body.state.doc, this.ctx);
    // While the app checks itself: anything Word would call damage is an error.
    if ((window as { __DUCKWRITER_TEST?: boolean }).__DUCKWRITER_TEST) {
      for (const problem of await checkPackage(bytes)) console.error(`Saved file: ${problem}`);
    }
    return bytes;
  }

  /** Saves; resolves to false if cancelled or failed. */
  async save(saveAs = false): Promise<boolean> {
    let path = this.path;
    if (!path || saveAs) {
      path = await pickSavePath(this.name);
      if (!path) return false;
    }
    if (!(await this.writeTo(path))) return false;
    this.toast(`Saved ${this.name}`);
    return true;
  }

  /** Auto-save: to the document's own file, without asking or a message. */
  async saveQuietly(): Promise<boolean> {
    return !!this.path && (await this.writeTo(this.path, true));
  }

  private async writeTo(path: string, quiet = false): Promise<boolean> {
    const doc = this.body.state.doc;
    const titlePage = this.ctx.titlePage;
    try {
      const bytes = await this.build();
      // As the kind of file it is saved as: with or without macros, a template, or Word XML in one file.
      let out = await convertPackage(bytes, packageKind(path) ?? "docx");
      if (/\.xml$/i.test(path)) out = new TextEncoder().encode(await toFlatOpc(out));
      // The file as it was is kept first (auto-save: at most every ten minutes).
      if (isDesktop) await invoke("versions_keep", { path, minAgeSecs: quiet ? 600 : 0 }).catch((error) => console.warn("No earlier version kept", error));
      await writePath(path, out);
      this.base = bytes;
    } catch (error) {
      if (!quiet) this.toast(`Couldn't save: ${message(error)}`, "error");
      return false;
    }
    // The goal and today's words go along to the new name.
    const goal = this.wordGoal;
    const today = (() => {
      try {
        return localStorage.getItem(`duckwriter.today:${this.progressKey}`);
      } catch {
        return null;
      }
    })();
    if (isDesktop) this.path = path;
    this.name = baseName(path);
    if (goal && !this.wordGoal) this.wordGoal = goal;
    try {
      if (today && !localStorage.getItem(`duckwriter.today:${this.progressKey}`)) localStorage.setItem(`duckwriter.today:${this.progressKey}`, today);
    } catch {
      // Not remembered.
    }
    if (isDesktop && !quiet) addRecent(path);
    // What was written; changes made while saving still count as unsaved.
    this.savedDoc = doc;
    this.savedTitlePage = titlePage;
    this.lineGridSaved = true;
    // As written: last saved by, now, the next revision.
    this.ctx.properties = stamped(this.ctx.properties, this.ctx.savedBy ?? null);
    this.savedProperties = JSON.stringify(this.ctx.properties);
    this.autosave.saved();
    this.changed(false);
    return true;
  }

  toggleAutoSave() {
    this.autosave.toggle();
    this.toast(this.autosave.enabled ? "Auto-save is on" : "Auto-save is off");
  }

  /** Offers back work left by a DuckWriter that closed before it was saved. */
  async offerRecovery() {
    for (const found of await AutoSave.abandoned()) {
      const when = new Date(found.time).toLocaleString([], { dateStyle: "medium", timeStyle: "short" });
      const answer = await ask(
        "Restore unsaved work?",
        `DuckWriter closed before "${found.name}" was saved (${when}). Do you want it back?`,
        [
          { label: "Delete it", value: "delete" },
          { label: "Restore", value: "restore", primary: true },
        ],
      );
      if (answer === "restore") {
        try {
          const file = await readPath(found.docx);
          const { doc, context } = await readDocx(file.bytes);
          this.show(doc, context, found.name, found.path);
          // It isn't saved yet: Save puts it back in its file.
          this.savedDoc = schema.nodes.doc.create(null, schema.nodes.paragraph.create());
          this.changed(true);
        } catch (error) {
          this.toast(`Couldn't restore it: ${message(error)}`, "error");
          continue;
        }
      }
      if (answer !== null) await AutoSave.discard(found);
      if (answer === "restore") break;
    }
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
    if (answer === "discard") await this.autosave.clearRecovery();
    return answer === "discard";
  }

  // ---- view -----------------------------------------------------------------

  // ---- printing and PDF ---------------------------------------------------------

  /**
   * Lays the pages out for paper (no gaps between them, nothing that only
   * belongs on screen), runs `job`, and puts the screen back.
   */
  private async onPaper(job: () => Promise<void>) {
    this.headers.stop();
    this.footnotes.stop();
    const zoom = this.zoom;
    const p = this.ctx.page;
    this.pageRule.textContent = `@page { size: ${p.width}pt ${p.height}pt; margin: 0 }`;
    document.documentElement.classList.add("printing");
    // No blinking cursor on paper.
    (document.activeElement as HTMLElement | null)?.blur();
    this.setZoom(1);
    const done = new Promise<void>((resolve) => (this.laidOut = resolve));
    setPageGap(0);
    relayout(this.body);
    await done;
    // Pictures and contents place themselves after the layout.
    await new Promise((resolve) => setTimeout(resolve, 50));
    try {
      await job();
    } finally {
      document.documentElement.classList.remove("printing");
      setPageGap(24);
      this.setZoom(zoom);
      relayout(this.body);
    }
  }

  print() {
    void this.onPaper(async () => window.print());
  }

  /** Asks where, then saves the document as a PDF. */
  async savePdf() {
    const path = await pickPdfPath(this.name.replace(/\.docx$/i, ".pdf"));
    if (path) await this.savePdfTo(path);
  }

  async savePdfTo(path: string) {
    const p = this.ctx.page;
    try {
      await this.onPaper(() => invoke("save_pdf", { path, width: p.width / 72, height: p.height / 72 }));
      this.toast(`Saved ${baseName(path)}`);
    } catch (error) {
      this.toast(`Couldn't save the PDF: ${message(error)}`, "error");
    }
  }

  // ---- for the keyboard shortcuts ----------------------------------------------

  find(replace: boolean) {
    this.headers.stop();
    this.footnotes.stop();
    this.findBar.open(replace);
  }

  findNext(direction: 1 | -1) {
    this.findBar.go(direction);
  }

  // ---- reviewing ----------------------------------------------------------------

  /** Moves the page left to make room for the comments beside it, when the window is narrow. */
  private placeSheet() {
    if (!this.sheet || !this.comments) return;
    if (!this.comments.visible) {
      this.sheet.style.marginLeft = "";
      return;
    }
    const zoom = this.zoom || 1;
    const room = this.workspace.clientWidth - 48;
    const page = this.sheet.offsetWidth * zoom;
    const pane = 300 * zoom;
    const left = Math.max(0, Math.min((room - page) / 2, room - page - pane));
    this.sheet.style.marginLeft = `${left / zoom}px`;
  }

  async newComment() {
    if (!(await askName("Your name is shown with your comments."))) return;
    this.headers.stop();
    this.footnotes.stop();
    this.comments.add();
  }

  async toggleTracking() {
    if (!this.ctx.trackChanges && !(await askName("Your name is shown with the changes you make."))) return;
    this.ctx.trackChanges = !this.ctx.trackChanges;
    this.toast(this.ctx.trackChanges ? "Changes are tracked" : "Changes aren't tracked any more");
    this.changed(false, true);
    this.view.focus();
  }

  /** Goes to the next (1) or previous (-1) tracked change or comment. */
  stepChange(direction: 1 | -1) {
    const view = this.body;
    const { state } = view;
    const targets = [
      ...changes(state.doc).map((c) => ({ from: c.from, to: c.to, comment: null as string | null })),
      ...commentAnchors(state.doc)
        .filter((a) => !this.ctx.comments.items.get(a.id)?.parent)
        .map((a) => ({ from: a.from + 1, to: a.from + 1, comment: a.id })),
    ].sort((a, b) => a.from - b.from);
    if (!targets.length) return this.toast("No changes or comments");
    const { from, to } = state.selection;
    const target =
      direction > 0
        ? (targets.find((t) => t.from > from || (t.from === from && t.to > to)) ?? targets[0])
        : ([...targets].reverse().find((t) => t.from < from) ?? targets[targets.length - 1]);
    const tr = state.tr.setSelection(TextSelection.create(state.doc, target.from, target.to)).scrollIntoView();
    if (target.comment) tr.setMeta(reviewKey, { active: target.comment });
    view.dispatch(tr);
    view.focus();
  }

  /** Accepts or rejects the change at the cursor (or those selected), then goes to the next one. */
  decideChange(accept: boolean) {
    const view = this.body;
    const here = changesAt(view.state);
    if (!here.length) {
      const next = nextChange(view.state, 1);
      if (!next) return this.toast("No tracked changes");
      return this.stepChange(1);
    }
    const tr = decide(view.state, here, accept, this.ctx);
    if (!tr) return;
    view.dispatch(tr);
    const next = nextChange(view.state, 1);
    if (next) view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, next.from, next.to)).scrollIntoView());
    else this.toast("That was the last change");
    view.focus();
  }

  decideAllChanges(accept: boolean, author?: string) {
    const view = this.body;
    const count = changes(view.state.doc).filter((c) => author === undefined || c.author === author).length;
    const tr = decideAll(view.state, accept, this.ctx, author);
    if (!tr) return this.toast("No tracked changes");
    view.dispatch(tr);
    this.toast(`${accept ? "Accepted" : "Rejected"} ${count} ${count === 1 ? "change" : "changes"}${author !== undefined ? ` by ${author || "an unnamed reviewer"}` : ""}`);
    view.focus();
  }

  /** Selects the next word with a red line and opens its suggestions (F7). */
  nextSpellingMistake() {
    const view = this.view;
    if (!spellingOn()) {
      this.toast("Spelling is off (Ctrl+F7 turns it on)");
      return;
    }
    const here = mistakeAt(view.state, view.state.selection.from);
    const mistake = here && !view.state.selection.empty ? nextMistake(view.state) : (here ?? nextMistake(view.state));
    if (!mistake) {
      this.toast("No spelling mistakes");
      return;
    }
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, mistake.from, mistake.to)).scrollIntoView());
    // After the page has scrolled to it.
    requestAnimationFrame(() => openMenuFor(view, mistake));
  }

  toggleSpelling() {
    closeSpellMenu();
    setSpelling(!spellingOn());
    this.showSpelling();
    this.toast(spellingOn() ? "Spelling is checked" : "Spelling isn't checked");
  }

  private showSpelling() {
    const names = new Intl.DisplayNames(undefined, { type: "language" });
    const languages = this.spellLanguages.map((tag) => names.of(tag) ?? tag).join(", ");
    this.spellStatus.hidden = !this.spellLanguages.length;
    this.spellStatus.textContent = spellingOn() ? "Spelling on" : "Spelling off";
    this.spellStatus.title = `${spellingOn() ? "Checking spelling" : "Not checking spelling"} in ${languages} (Ctrl+F7 to switch; F7 for the next mistake).\nThe languages are the ones in Windows' language settings.`;
  }

  /** Zooms in (1), out (-1) or back to 100% (0). */
  zoomBy(direction: 1 | -1 | 0) {
    if (direction === 0) this.setZoom(1);
    else this.stepZoom(direction);
  }

  toggleSnap() {
    setSnap(!snapOn());
    this.toast(snapOn() ? "Pictures snap to the grid" : "Pictures move freely");
  }

  focusBox(box: "font" | "size" | "style") {
    this.toolbar.focusBox(box);
  }

  openMenu(menu: "symbol" | "references" | "layout" | "textColor" | "highlight" | "table" | "open" | "newDoc") {
    this.toolbar.openMenuByKeyboard(menu);
  }

  toggleTableLines() {
    this.toolbar.toggleTableLines();
  }

  applyTextColor() {
    this.toolbar.applyTextColor();
  }

  applyHighlight() {
    this.toolbar.applyHighlight();
  }

  showShortcuts() {
    showShortcutHelp();
  }

  private setZoom(zoom: number) {
    this.zoom = Math.min(5, Math.max(0.25, zoom));
    this.sheet.style.zoom = String(this.zoom);
    this.placeSheet();
    this.zoomLabel.textContent = `${Math.round(this.zoom * 100)}%`;
  }

  /** The zoom at which the page fills the window's width, or all of the page shows. */
  private fitZoom(whole: boolean): number {
    const p = this.ctx.page;
    const room = this.workspace.clientWidth - 48 - (this.comments.visible ? 300 : 0);
    const width = room / (p.width * (4 / 3));
    if (!whole) return Math.round(width * 100) / 100;
    const height = (this.workspace.clientHeight - 48) / (p.height * (4 / 3));
    return Math.round(Math.min(width, height) * 100) / 100;
  }

  /** 100% → page width → whole page → 100% (Alt+Shift+Z, or a click on the zoom). */
  cycleZoomFit() {
    const width = this.fitZoom(false);
    const whole = this.fitZoom(true);
    const at = (z: number) => Math.abs(this.zoom - z) < 0.011;
    if (at(width) && !at(whole)) this.setZoom(whole);
    else if (at(1) || (!at(width) && !at(whole))) this.setZoom(width);
    else this.setZoom(1);
    this.body.dispatch(this.body.state.tr.scrollIntoView());
  }

  private pageShown = false;

  /**
   * "Page 3 of 9": the page the cursor is on. Once per frame, after the page
   * layout: measuring right after a change lays out the whole document an extra time.
   */
  private showPage() {
    if (this.pageShown) return;
    this.pageShown = true;
    requestAnimationFrame(() => {
      this.pageShown = false;
      this.showPageNow();
    });
  }

  private showPageNow() {
    const count = this.pages;
    let page = 1;
    try {
      const top = this.body.coordsAtPos(this.body.state.selection.head).top;
      const origin = this.sheet.getBoundingClientRect().top;
      page = Math.min(count, Math.max(1, pageAt(this.lastLayout, (top - origin) / (this.zoom || 1), this.ctx.page) + 1));
    } catch {
      // Before the first layout.
    }
    this.pageCount.textContent = count > 1 ? `Page ${page} of ${count}` : "1 page";
  }

  private stepZoom(direction: 1 | -1) {
    const next = direction > 0 ? ZOOMS.find((z) => z > this.zoom + 0.001) : [...ZOOMS].reverse().find((z) => z < this.zoom - 0.001);
    if (next) this.setZoom(next);
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
// A window opened for another document opens it; the first one opens what DuckWriter was started with.
const opening = windowOpening();
const [first] = opening === undefined ? await startupFiles() : [opening];
await app.newDocument(false);
if (first?.startsWith("template:")) await app.newFromTemplate(first.slice(9) as TemplateKind);
else if (first?.startsWith("mytemplate:")) await app.newFromMyTemplate(first.slice(11));
else if (first) await app.openPath(first);
if (localStorage.getItem("duckwriter.debug") === "1") {
  Object.assign(window, { __dw: app });
}
if (opening === undefined) {
  void app.offerRecovery();
  // Give the document a moment to open before going online. Not while testing.
  if (!navigator.webdriver && localStorage.getItem("duckwriter.debug") !== "1") {
    setTimeout(() => void offerUpdate(document.querySelector(".toolbar")!, { beforeInstall: () => app.confirmDiscard() }), 3000);
  }
}
