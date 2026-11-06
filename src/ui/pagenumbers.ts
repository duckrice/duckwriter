// Page numbers, all in one window: where they go on the page, every page with the number it
// will show, and each section of the document: how it starts, how its pages are numbered
// (i, ii, iii in the front matter, 1, 2, 3 from the first chapter), where the count starts,
// and whether its first page (a title page) has none. Section breaks are added and taken out here too.
// Every change is made at once, so the pages and the strip of page numbers show it straight away.

import type { EditorView } from "prosemirror-view";
import { h } from "../dom";
import { insertSectionBreak, removeSectionBreak, sectionsOf, setSectionNumbering, setSectionStart, type PageNumbering, type SectionStart } from "../editor/commands";
import type { PageLayout } from "../editor/pagination";
import { schema } from "../editor/schema";
import type { DocContext } from "../docx/model";

export type NumbersAt = "none" | "bottom-center" | "bottom-right" | "top-right";

export interface PageNumbersHost {
  body: EditorView;
  ctx: DocContext;
  layout(): PageLayout | null;
  /** Puts page numbers in the header or footer ("none" takes them out). */
  placeNumbers(where: NumbersAt): void;
  /** After a change to the sections: marks the document changed. */
  sectionsChanged(): void;
  goTo(pos: number): void;
}

const FORMATS: [PageNumbering["format"], string][] = [
  ["decimal", "1, 2, 3"],
  ["lowerRoman", "i, ii, iii"],
  ["upperRoman", "I, II, III"],
  ["lowerLetter", "a, b, c"],
  ["upperLetter", "A, B, C"],
];
const STARTS: [SectionStart, string][] = [
  ["nextPage", "On a new page"],
  ["continuous", "On the same page (no new page)"],
  ["oddPage", "On the next odd page"],
  ["evenPage", "On the next even page"],
];
const PLACES: [NumbersAt, string][] = [
  ["none", "None"],
  ["bottom-center", "Bottom, centred: Page 2 of 9"],
  ["bottom-right", "Bottom, right: 2"],
  ["top-right", "Top, right: 2"],
];

/** Redraws the window, if it is open (after the pages are laid out again). */
let refresh: (() => void) | null = null;
export const refreshPageNumbers = () => refresh?.();

/** Where the page numbers are now, as far as the header and footer say. */
export function numbersAt(ctx: DocContext): NumbersAt {
  let found: NumbersAt = "none";
  for (const story of ctx.stories) {
    if (story.type !== "default") continue;
    story.doc.descendants((n) => {
      if (n.type !== schema.nodes.field) return;
      if (/^\s*NUMPAGES\b/.test(n.attrs.instr) && story.kind === "footer") found = "bottom-center";
      else if (/^\s*PAGE\b/.test(n.attrs.instr) && found === "none") found = story.kind === "header" ? "top-right" : "bottom-right";
    });
  }
  return found;
}

export function pageNumbersWindow(host: PageNumbersHost) {
  const { body, ctx } = host;
  const content = h("div", { class: "pn-content" });
  const dialog = h(
    "dialog",
    { class: "dialog pn-dialog", "aria-label": "Page numbers" },
    h("h2", {}, "Page numbers"),
    content,
    h(
      "div",
      { class: "dialog-footer" },
      h("span", { class: "dialog-hint pn-footer-hint" }, "Changes are made at once; Ctrl+Z undoes them."),
      h("button", { type: "button", class: "btn primary", onclick: () => dialog.close() }, "Close"),
    ),
  );

  const run = (command: (s: typeof body.state, d: typeof body.dispatch, v: EditorView) => boolean) => {
    command(body.state, body.dispatch, body);
    host.sectionsChanged();
    draw();
  };

  function draw() {
    const sections = sectionsOf(body.state.doc, ctx);
    const layout = host.layout();
    const at = numbersAt(ctx);
    const scroll = content.scrollTop;

    // Where the numbers go.
    const places = h(
      "div",
      { class: "pn-places", role: "radiogroup", "aria-label": "Where the page numbers go" },
      ...PLACES.map(([value, label]) => {
        const input = h("input", { type: "radio", name: "pn-place", value, checked: at === value }) as HTMLInputElement;
        input.addEventListener("change", () => {
          host.placeNumbers(value);
          draw();
        });
        return h("label", { class: "check-field" }, input, ` ${label}`);
      }),
    );

    // Every page with its number, coloured by section.
    const labels = layout?.labels ?? [];
    const firsts = (layout?.sections ?? []).map((s) => s.firstPage);
    const sectionOfPage = (k: number) => Math.max(0, firsts.filter((f) => f <= k).length - 1);
    const strip = h(
      "div",
      { class: "pn-strip", "aria-label": "Each page and the number it shows" },
      ...labels.map((label, k) => {
        const s = sectionOfPage(k);
        const blank = at === "none" || (firsts[s] === k && sections[s]?.numbering.firstBlank);
        return h(
          "button",
          {
            type: "button",
            class: `pn-page pn-s${s % 4}${blank ? " blank" : ""}`,
            title: `Page ${k + 1}, section ${s + 1}: ${blank ? "no number" : `shows ${label}`}`,
            onclick: () => goToPage(k),
          },
          h("span", { class: "pn-page-number" }, blank ? "–" : label),
        );
      }),
    );

    const rows = sections.map((section, i) => sectionRow(section, i, sections.length, layout));
    const add = h("button", { type: "button", class: "btn small" }, "Add a section break where the cursor is");
    add.addEventListener("click", () => {
      if (!insertSectionBreak(ctx, "nextPage")(body.state)) {
        add.textContent = "Put the cursor in the text first (not in a table), then try again";
        return;
      }
      run((s, d) => insertSectionBreak(ctx, "nextPage")(s, d));
    });

    content.replaceChildren(
      h("h3", {}, "Where they go"),
      places,
      h("h3", {}, "Your pages"),
      labels.length ? strip : h("p", { class: "dialog-hint" }, "The pages are being laid out…"),
      at === "none" ? h("p", { class: "dialog-hint" }, "No page numbers yet: choose where they go above.") : "",
      h("h3", {}, sections.length > 1 ? `Sections (${sections.length})` : "Sections"),
      h(
        "p",
        { class: "dialog-hint" },
        "A section break splits the document into parts that number their pages on their own: i, ii, iii for the front matter, then 1, 2, 3 from the first chapter. Click in the text where the new numbering should start, then add a section break.",
      ),
      ...rows,
      h("div", { class: "pn-add" }, add),
    );
    content.scrollTop = scroll;
  }

  function sectionRow(section: ReturnType<typeof sectionsOf>[number], i: number, count: number, layout: PageLayout | null) {
    const n = section.numbering;
    const first = layout?.sections?.[i]?.firstPage;
    const last = i + 1 < (layout?.sections?.length ?? 0) ? layout!.sections![i + 1].firstPage - 1 : (layout?.count ?? 0) - 1;
    const pages = first === undefined ? "" : first >= last ? `page ${first + 1}` : `pages ${first + 1}–${last + 1}`;
    const set = (change: Partial<PageNumbering>) => run((s, d, v) => setSectionNumbering(ctx, section.end, { ...n, ...change })(s, d, v));

    const format = h("select", { class: "tb-select", "aria-label": `Section ${i + 1}: number style` }, ...FORMATS.map(([value, label]) => h("option", { value }, label))) as HTMLSelectElement;
    format.value = n.format;
    format.addEventListener("change", () => set({ format: format.value as PageNumbering["format"] }));

    const name = `pn-from-${i}`;
    const goOn = h("input", { type: "radio", name, checked: n.start === null }) as HTMLInputElement;
    const restart = h("input", { type: "radio", name, checked: n.start !== null }) as HTMLInputElement;
    const start = h("input", { class: "tb-input dialog-input number", type: "number", min: "0", value: String(n.start ?? 1), "aria-label": `Section ${i + 1}: start at` }) as HTMLInputElement;
    goOn.addEventListener("change", () => set({ start: null }));
    restart.addEventListener("change", () => set({ start: Math.max(0, Math.round(Number(start.value) || 0)) }));
    start.addEventListener("change", () => set({ start: Math.max(0, Math.round(Number(start.value) || 0)) }));

    const firstBlank = h("input", { type: "checkbox", checked: n.firstBlank }) as HTMLInputElement;
    firstBlank.addEventListener("change", () => set({ firstBlank: firstBlank.checked }));

    const begins = h("select", { class: "tb-select", "aria-label": `Section ${i + 1}: starts` }, ...STARTS.map(([value, label]) => h("option", { value }, label))) as HTMLSelectElement;
    begins.value = section.start;
    begins.addEventListener("change", () => run((s, d, v) => setSectionStart(ctx, section.end, begins.value as SectionStart)(s, d, v)));

    const go = h("button", { type: "button", class: "btn small subtle" }, "Go there");
    go.addEventListener("click", () => {
      dialog.close();
      host.goTo(section.from);
    });
    const actions = h("div", { class: "pn-actions" }, go);
    if (i + 1 < count) {
      const remove = h("button", { type: "button", class: "btn small subtle" }, "Remove the break after it");
      remove.title = "Joins this section with the next one, which keeps its own setup";
      remove.addEventListener("click", () => run((s, d, v) => removeSectionBreak(section.end)(s, d, v)));
      actions.append(remove);
    }

    return h(
      "fieldset",
      { class: `pn-section pn-s${i % 4}` },
      h("legend", {}, `Section ${i + 1}`, pages ? ` · ${pages}` : "", section.preview ? ` · “${section.preview}${section.preview.length >= 60 ? "…" : ""}”` : " · (empty)"),
      i > 0 ? h("label", { class: "pn-field" }, "Starts ", begins) : "",
      h("label", { class: "pn-field" }, "Numbers ", format),
      h("label", { class: "check-field" }, goOn, i > 0 ? " Go on from the section before" : " Start at 1"),
      h("label", { class: "check-field" }, restart, " Start at ", start),
      h("label", { class: "check-field" }, firstBlank, " No number on its first page (a title page)"),
      actions,
    );
  }

  function goToPage(k: number) {
    const layout = host.layout();
    if (!layout) return;
    let target = -1;
    for (const [pos, page] of layout.pageOfBlock) if (page === k && (target < 0 || pos < target)) target = pos;
    if (target < 0) return;
    dialog.close();
    host.goTo(target);
  }

  refresh = draw;
  dialog.addEventListener("close", () => {
    refresh = null;
    dialog.remove();
    body.focus();
  });
  document.body.append(dialog);
  draw();
  dialog.showModal();
}
