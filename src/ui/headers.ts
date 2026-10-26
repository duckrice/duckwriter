// Headers and footers on every page. Double-click one (or use the toolbar
// button) to edit it, as in Word; the page number and page count fill in on
// each page by themselves.

import { DOMSerializer } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import { h } from "../dom";
import type { DocContext, Story } from "../docx/model";
import { ensureStyles } from "../docx/builtins";
import { createState, createView } from "../editor/editor";
import { pageBox, pageOffset, type PageLayout } from "../editor/pagination";
import { schema } from "../editor/schema";

export interface HeaderHost {
  readonly ctx: DocContext;
  readonly body: EditorView;
  /** Called when a header or footer changed, or editing started or stopped. */
  storiesChanged(): void;
}

type Kind = Story["kind"];

const PX = 4 / 3;

export class HeadersAndFooters {
  /** Sits above the text, holding the header and footer areas of every page. */
  readonly layer = h("div", { class: "hf-layer" });
  private bar: HTMLElement;
  private firstPageBox = h("input", { type: "checkbox" }) as HTMLInputElement;
  private evenBox = h("input", { type: "checkbox" }) as HTMLInputElement;
  private barLabel = h("span", { class: "hf-bar-label" });
  private pageCount = 1;
  private editing: { kind: Kind; page: number; story: Story; view: EditorView; isNew: boolean } | null = null;

  constructor(
    private host: HeaderHost,
    barParent: HTMLElement,
  ) {
    const button = (label: string, title: string, onClick: () => void) => {
      const b = h("button", { type: "button", class: "btn small", title }, label);
      b.addEventListener("mousedown", (event) => event.preventDefault());
      b.addEventListener("click", onClick);
      return b;
    };
    this.evenBox.addEventListener("change", () => {
      this.host.ctx.evenAndOdd = this.evenBox.checked;
      const editing = this.editing;
      if (editing) {
        this.stop();
        this.start(editing.kind, editing.page);
      }
      this.host.storiesChanged();
    });
    this.firstPageBox.addEventListener("change", () => {
      this.host.ctx.titlePage = this.firstPageBox.checked;
      const editing = this.editing;
      if (editing) {
        this.stop();
        this.start(editing.kind, editing.page);
      }
      this.host.storiesChanged();
    });
    this.bar = h(
      "div",
      { class: "hf-bar", hidden: true },
      this.barLabel,
      button("Page number", "Insert the page number", () => this.insertField(" PAGE ", "1")),
      button("Number of pages", "Insert the number of pages", () => this.insertField(" NUMPAGES ", String(this.pageCount))),
      h("label", { class: "hf-check" }, this.firstPageBox, "Different first page"),
      h("label", { class: "hf-check" }, this.evenBox, "Different odd & even pages"),
      button("Close", "Back to the text (Esc)", () => this.stop()),
    );
    barParent.prepend(this.bar);
    this.layer.addEventListener("dblclick", (event) => {
      const area = (event.target as HTMLElement).closest<HTMLElement>(".hf");
      if (!area || area.classList.contains("editing")) return;
      // The same double-click must not reach the page, where it means "back to the text".
      event.stopPropagation();
      this.start(area.dataset.kind as Kind, Number(area.dataset.page));
    });
  }

  /** The editor of the header or footer being edited, if any. */
  get view(): EditorView | null {
    return this.editing?.view ?? null;
  }

  /** The story shown on a page: the first page may have its own; an earlier section may have its own. */
  private storyFor(kind: Kind, page: number): Story | undefined {
    const { stories, titlePage, sectionStories } = this.host.ctx;
    let s = this.sections.length - 1;
    while (s > 0 && this.sections[s].firstPage > page) s--;
    if (s >= 0 && s < this.sections.length - 1) {
      // An earlier section: its own references, or else the section before's (as Word does).
      const xml = this.sections[s].xml;
      const first = /<w:titlePg[ />]/.test(xml) && !/<w:titlePg[^>]*w:val="(0|false)"/.test(xml) && page === this.sections[s].firstPage;
      const type = first ? "first" : this.even(page) ? "even" : "default";
      for (let j = s; j >= 0; j--) {
        // The section's references of this kind: <w:headerReference w:type="default" r:id="rId8"/>.
        const tags = this.sections[j].xml.split(`<w:${kind}Reference`).slice(1).map((t) => t.slice(0, t.indexOf(">")));
        const tag = tags.find((t) => (/w:type="([^"]*)"/.exec(t)?.[1] ?? "default") === type);
        const rid = tag && /r:id="([^"]*)"/.exec(tag)?.[1];
        if (rid) return sectionStories?.get(rid) ?? stories.find((x) => x.rid === rid);
      }
      return undefined;
    }
    const type = titlePage && page === 0 ? "first" : this.even(page) ? "even" : "default";
    return stories.find((s) => s.kind === kind && s.type === type);
  }

  /** Whether a page gets the even-page header and footer: by the number it shows, as in Word. */
  private even(page: number): boolean {
    if (!this.host.ctx.evenAndOdd) return false;
    const shown = Number(this.labels[page]);
    return (Number.isFinite(shown) ? shown : page + 1) % 2 === 0;
  }

  /** Draws the headers and footers of `count` pages. */
  render(count: number, labels: string[] = [], sections: { firstPage: number; xml: string }[] = [], layout: PageLayout | null = this.layout) {
    this.pageCount = count;
    this.labels = labels;
    this.sections = sections;
    this.layout = layout;
    const main = this.host.ctx.page;
    const areas = [...this.layer.querySelectorAll<HTMLElement>(".hf")];
    const wanted = count * 2;
    while (areas.length < wanted) {
      const kind: Kind = areas.length % 2 ? "footer" : "header";
      const area = h("div", { class: `hf ${kind}`, title: `Double-click to edit the ${kind}` });
      area.dataset.kind = kind;
      this.layer.append(area);
      areas.push(area);
    }
    while (areas.length > wanted) areas.pop()!.remove();
    areas.forEach((area, i) => {
      const page = Math.floor(i / 2);
      const kind = area.dataset.kind as Kind;
      area.dataset.page = String(page);
      // Each page's own size and margins (a section may have its own).
      const { top, setup: p } = pageBox(layout, page, main);
      area.style.left = `${pageOffset(p, main) + p.marginLeft}pt`;
      area.style.width = `${p.width - p.marginLeft - p.marginRight}pt`;
      if (kind === "header") {
        area.style.top = `${top + p.headerDistance * PX}px`;
        area.style.minHeight = `${Math.max(12, p.marginTop - p.headerDistance)}pt`;
      } else {
        // Its bottom sits the footer distance above the bottom of the page; it grows upwards.
        area.style.top = `${top + (p.height - p.footerDistance) * PX}px`;
        area.style.minHeight = `${Math.max(12, p.marginBottom - p.footerDistance)}pt`;
      }
      const editing = this.editing && this.editing.kind === kind && this.editing.page === page;
      if (!editing) this.fill(area, this.storyFor(kind, page), page);
    });
  }

  /** What each area shows, so that unchanged ones aren't drawn again (hundreds of pages). */
  private shown = new WeakMap<HTMLElement, { doc: unknown; page: number; count: number; label?: string }>();
  /** Each page's number as shown (sections may restart it or show i, ii, iii). */
  private labels: string[] = [];
  /** The last layout, for each page's size. */
  private layout: PageLayout | null = null;
  /** The sections, from the last layout. */
  private sections: { firstPage: number; xml: string }[] = [];

  /** Shows a story's content with this page's number in it. */
  private fill(area: HTMLElement, story: Story | undefined, page: number) {
    const was = this.shown.get(area);
    if (was && !area.classList.contains("editing") && was.doc === story?.doc && was.page === page && was.count === this.pageCount && was.label === this.labels[page]) return;
    this.shown.set(area, { doc: story?.doc, page, count: this.pageCount, label: this.labels[page] });
    area.classList.remove("editing");
    const content = h("div", { class: "hf-doc" });
    if (story) content.append(DOMSerializer.fromSchema(schema).serializeFragment(story.doc.content));
    for (const field of content.querySelectorAll<HTMLElement>(".field")) {
      if (field.dataset.instr === "PAGE") field.textContent = this.labels[page] ?? String(page + 1);
      else if (field.dataset.instr === "NUMPAGES") field.textContent = String(this.pageCount);
    }
    area.classList.toggle("empty", !story || !story.doc.textContent.trim() && !content.querySelector(".field"));
    area.replaceChildren(content);
  }

  /** Starts editing the header or footer of a page. */
  start(kind: Kind, page: number) {
    this.stop();
    const ctx = this.host.ctx;
    let story = this.storyFor(kind, page);
    const isNew = !story;
    if (!story) {
      const style = kind === "header" ? "Header" : "Footer";
      ensureStyles(ctx, style);
      story = {
        kind,
        type: ctx.titlePage && page === 0 ? "first" : this.even(page) ? "even" : "default",
        path: null,
        rid: null,
        rootStart: null,
        prefixes: new Set(),
        rels: new Map(),
        doc: schema.nodes.doc.create(null, schema.nodes.paragraph.create({ style: ctx.styles.byId.has(style) ? style : null })),
        saved: null,
      };
      ctx.stories.push(story);
    }
    const area = this.layer.querySelector<HTMLElement>(`.hf.${kind}[data-page="${page}"]`);
    if (!area) return;
    area.classList.add("editing");
    area.classList.remove("empty");
    area.replaceChildren();
    const s = story;
    const view = createView(
      area,
      createState(s.doc, () => this.host.ctx),
      {
        onChange: (v, docChanged) => {
          if (docChanged) s.doc = v.state.doc;
          this.host.storiesChanged();
        },
      },
      "hf-doc",
    );
    view.dom.addEventListener("keydown", (event) => {
      if ((event as KeyboardEvent).key === "Escape") {
        event.preventDefault();
        this.stop();
        this.host.body.focus();
      }
    });
    this.editing = { kind, page, story: s, view, isNew };
    this.layer.parentElement?.classList.add("editing-hf");
    this.firstPageBox.checked = ctx.titlePage;
    this.evenBox.checked = !!ctx.evenAndOdd;
    this.barLabel.textContent = `${kind === "header" ? "Header" : "Footer"}${s.type === "first" ? " (first page)" : ""}`;
    this.bar.hidden = false;
    view.focus();
    this.host.storiesChanged();
  }

  /** Goes back to editing the text. */
  stop() {
    const editing = this.editing;
    if (!editing) return;
    this.editing = null;
    editing.story.doc = editing.view.state.doc;
    editing.view.destroy();
    // A header that was added but left empty isn't kept.
    if (editing.isNew && !editing.story.doc.textContent.trim() && !hasField(editing.story)) {
      const stories = this.host.ctx.stories;
      stories.splice(stories.indexOf(editing.story), 1);
    }
    this.layer.parentElement?.classList.remove("editing-hf");
    this.bar.hidden = true;
    this.render(this.pageCount, this.labels, this.sections, this.layout);
    this.host.storiesChanged();
  }

  private insertField(instr: string, result: string) {
    const view = this.editing?.view;
    if (!view) return;
    view.dispatch(view.state.tr.replaceSelectionWith(schema.nodes.field.create({ instr, result }), true).scrollIntoView());
    view.focus();
  }
}

function hasField(story: Story): boolean {
  let found = false;
  story.doc.descendants((n) => {
    if (n.type === schema.nodes.field) found = true;
  });
  return found;
}
