// Sources and citations: a list of the document's sources to add to, edit,
// import (BibTeX, or a DOI looked up online); citing them in the text
// (Alt+Shift+C); and the reference list (Alt+Shift+B).

import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { CITATION_STYLES, SOURCE_TYPES, emptySource, makeTag, parseBibtex, type CitationStyle, type Source, type SourceType } from "../docx/bibliography";
import { ensureStyles } from "../docx/builtins";
import type { DocContext } from "../docx/model";
import { h } from "../dom";
import { citationInstr, citationsKey, formatCitation, citedOrder } from "../editor/citations";
import { schema } from "../editor/schema";

/** Tells the editor the sources changed, so citations and the reference list follow. */
function refresh(view: EditorView, ctx: DocContext) {
  ctx.bibliography.version++;
  view.dispatch(view.state.tr.setMeta(citationsKey, true));
  // Others writing along get the sources too.
  view.dom.dispatchEvent(new CustomEvent("duckwriter-sources", { bubbles: true }));
}

const label = (s: Source) => {
  const who = s.authors.length ? `${s.authors[0].last}${s.authors.length > 1 ? (s.authors.length > 2 ? " et al." : ` & ${s.authors[1].last}`) : ""}` : "";
  return [who, s.year && `(${s.year})`, s.title].filter(Boolean).join(" ");
};

function modal(title: string, ...content: (HTMLElement | null)[]): HTMLDialogElement {
  const dialog = h("dialog", { class: "dialog wide" }, h("h2", {}, title), ...content) as HTMLDialogElement;
  dialog.addEventListener("close", () => dialog.remove());
  document.body.append(dialog);
  dialog.showModal();
  return dialog;
}

// ---- looking a DOI up --------------------------------------------------------------------

const CROSSREF_TYPES: Record<string, SourceType> = {
  "journal-article": "JournalArticle",
  book: "Book",
  monograph: "Book",
  "book-chapter": "BookSection",
  "proceedings-article": "ConferenceProceedings",
  report: "Report",
};

/** A source's details from its DOI, through Crossref (online). */
export async function lookUpDoi(doi: string): Promise<Source> {
  const id = doi.trim().replace(/^https?:\/\/(dx\.)?doi\.org\//i, "").replace(/^doi:\s*/i, "");
  const response = await fetch(`https://api.crossref.org/works/${encodeURIComponent(id)}`);
  if (!response.ok) throw new Error(response.status === 404 ? "No source has that DOI." : `The lookup failed (${response.status}).`);
  const m = (await response.json()).message;
  const s = emptySource();
  s.type = CROSSREF_TYPES[m.type] ?? "Misc";
  s.title = (m.title?.[0] ?? "").replace(/\s+/g, " ");
  s.authors = (m.author ?? []).map((a: { family?: string; given?: string; name?: string }) => ({ last: a.family ?? a.name ?? "", first: a.given ?? "" }));
  s.year = String(m.issued?.["date-parts"]?.[0]?.[0] ?? m.published?.["date-parts"]?.[0]?.[0] ?? "");
  s.container = m["container-title"]?.[0] ?? "";
  s.volume = m.volume ?? "";
  s.issue = m.issue ?? "";
  s.pages = (m.page ?? "").replace(/-/g, "–");
  s.publisher = m.publisher ?? "";
  s.doi = m.DOI ?? id;
  return s;
}

// ---- the list of sources ------------------------------------------------------------------

/** Shows the sources, to add, edit, import or remove, and the citation style. */
export function manageSources(view: EditorView, ctx: DocContext, afterwards?: () => void) {
  const bib = ctx.bibliography;
  const list = h("div", { class: "source-list", role: "listbox", "aria-label": "Sources" });
  const style = h("select", { class: "tb-select", "aria-label": "Citation style" }, ...CITATION_STYLES.map((s) => h("option", { value: s.style }, s.label))) as HTMLSelectElement;
  style.value = bib.style;
  style.addEventListener("change", () => {
    bib.style = style.value as CitationStyle;
    refresh(view, ctx);
  });

  const render = () => {
    const cited = new Set(citedOrder(view.state.doc));
    list.replaceChildren(
      ...(bib.sources.length
        ? bib.sources.map((s) => {
            const item = h("div", { class: "source-item" }, h("span", { class: "source-tag" }, s.tag), h("span", { class: "source-label" }, label(s) || "(no details)"), cited.has(s.tag) ? h("span", { class: "source-cited", title: "Cited in the text" }, "cited") : null);
            const edit = h("button", { type: "button", class: "btn small" }, "Edit");
            edit.addEventListener("click", () => editSource(s, false));
            const remove = h("button", { type: "button", class: "btn small", title: cited.has(s.tag) ? "Cited in the text: its citations would show a question mark" : "Remove" }, "Remove");
            remove.addEventListener("click", () => {
              bib.sources.splice(bib.sources.indexOf(s), 1);
              refresh(view, ctx);
              render();
            });
            item.append(edit, remove);
            return item;
          })
        : [h("div", { class: "nav-empty" }, "No sources yet. Add one, import BibTeX from your reference manager or Google Scholar, or look one up by its DOI.")]),
    );
  };

  /** The form for one source. */
  const editSource = (source: Source, isNew: boolean) => {
    const s = { ...source, authors: source.authors.map((a) => ({ ...a })) };
    const field = (name: keyof Source, text: string, wide = false) => {
      const input = h("input", { class: `tb-input dialog-input${wide ? " wide" : ""}`, value: String(s[name] ?? "") }) as HTMLInputElement;
      input.addEventListener("input", () => ((s as unknown as Record<string, string>)[name] = input.value));
      return h("label", {}, text, input);
    };
    const type = h("select", { class: "tb-select" }, ...SOURCE_TYPES.map((t) => h("option", { value: t.type }, t.label))) as HTMLSelectElement;
    type.value = s.type;
    type.addEventListener("change", () => (s.type = type.value as SourceType));
    const authors = h("textarea", { class: "tb-input dialog-input wide", rows: 3, placeholder: "One per line: Last name, First names" }) as HTMLTextAreaElement;
    authors.value = s.authors.map((a) => [a.last, a.first].filter(Boolean).join(", ")).join("\n");
    authors.addEventListener("input", () => {
      s.authors = authors.value
        .split("\n")
        .map((line) => line.trim())
        .filter(Boolean)
        .map((line) => {
          const [last, first] = line.includes(",") ? line.split(",", 2) : [line.split(" ").pop() ?? "", line.split(" ").slice(0, -1).join(" ")];
          return { last: last.trim(), first: (first ?? "").trim() };
        });
    });
    const form = modal(
      isNew ? "New source" : "Edit source",
      h(
        "div",
        { class: "dialog-form source-form" },
        h("label", {}, "Kind", type),
        h("label", {}, "Authors", authors),
        field("title", "Title", true),
        field("year", "Year"),
        field("container", "Journal / book / conference", true),
        field("volume", "Volume"),
        field("issue", "Issue"),
        field("pages", "Pages"),
        field("publisher", "Publisher"),
        field("city", "City"),
        field("doi", "DOI", true),
        field("url", "Web address", true),
      ),
      h(
        "div",
        { class: "dialog-footer" },
        h("button", { type: "button", class: "btn", onclick: () => form.close() }, "Cancel"),
        h(
          "button",
          {
            type: "button",
            class: "btn primary",
            onclick: () => {
              if (!s.tag || isNew) s.tag = makeTag(s, bib.sources);
              s.xml = null;
              if (isNew) bib.sources.push(s);
              else bib.sources[bib.sources.indexOf(source)] = s;
              form.close();
              refresh(view, ctx);
              render();
            },
          },
          isNew ? "Add" : "Save",
        ),
      ),
    );
    (form.querySelector("textarea") as HTMLElement).focus();
  };

  const add = h("button", { type: "button", class: "btn" }, "Add a source…");
  add.addEventListener("click", () => editSource(emptySource(), true));

  const importBox = h("textarea", { class: "tb-input dialog-input wide", rows: 3, placeholder: "Paste BibTeX here (@article{…})" }) as HTMLTextAreaElement;
  const importButton = h("button", { type: "button", class: "btn" }, "Import BibTeX");
  const status = h("p", { class: "dialog-error", role: "status" });
  importButton.addEventListener("click", () => {
    const found = parseBibtex(importBox.value);
    if (!found.length) {
      status.textContent = "No BibTeX entries found there.";
      return;
    }
    for (const s of found) {
      s.tag = makeTag(s, [...bib.sources, ...found.filter((f) => f !== s && f.tag)]);
      bib.sources.push(s);
    }
    importBox.value = "";
    status.textContent = "";
    refresh(view, ctx);
    render();
  });

  const doi = h("input", { class: "tb-input dialog-input", placeholder: "10.1000/xyz123", "aria-label": "DOI" }) as HTMLInputElement;
  const lookUp = h("button", { type: "button", class: "btn" }, "Look up DOI");
  lookUp.addEventListener("click", async () => {
    if (!doi.value.trim()) return doi.focus();
    lookUp.disabled = true;
    status.textContent = "";
    try {
      const s = await lookUpDoi(doi.value);
      s.tag = makeTag(s, bib.sources);
      bib.sources.push(s);
      doi.value = "";
      refresh(view, ctx);
      render();
    } catch (error) {
      status.textContent = error instanceof TypeError ? "Couldn't reach Crossref (are you online?)." : String((error as Error).message ?? error);
    } finally {
      lookUp.disabled = false;
    }
  });
  doi.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      lookUp.click();
    }
  });

  render();
  const dialog = modal(
    "Sources",
    h("div", { class: "dialog-form" }, h("label", {}, "Citation style", style)),
    list,
    h("div", { class: "source-actions" }, add, h("span", { class: "dialog-spacer" }), doi, lookUp),
    h("div", { class: "source-actions" }, importBox, importButton),
    status,
    h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn primary", onclick: () => dialog.close() }, "Done")),
  );
  dialog.addEventListener("close", () => {
    view.focus();
    afterwards?.();
  });
  add.focus();
}

// ---- citing ----------------------------------------------------------------------------------

/** Asks which source(s) to cite, and puts the citation at the cursor (Alt+Shift+C). */
export function insertCitation(view: EditorView, ctx: DocContext) {
  const bib = ctx.bibliography;
  if (!bib.sources.length) return manageSources(view, ctx, () => bib.sources.length && insertCitation(view, ctx));
  const search = h("input", { class: "tb-input dialog-input wide", placeholder: "Search: author, title, year", "aria-label": "Search the sources" }) as HTMLInputElement;
  const page = h("input", { class: "tb-input dialog-input number", placeholder: "e.g. 23", "aria-label": "Page" }) as HTMLInputElement;
  const list = h("div", { class: "source-list", role: "listbox", "aria-multiselectable": "true" });
  const chosen = new Set<string>();
  const preview = h("p", { class: "dialog-hint" });
  let current = 0;
  const visible = () => bib.sources.filter((s) => `${s.tag} ${label(s)}`.toLowerCase().includes(search.value.toLowerCase()));
  const render = () => {
    const shown = visible();
    current = Math.max(0, Math.min(current, shown.length - 1));
    list.replaceChildren(
      ...shown.map((s, i) => {
        const box = h("input", { type: "checkbox", tabindex: "-1" }) as HTMLInputElement;
        box.checked = chosen.has(s.tag);
        const item = h("div", { class: `source-item pick${i === current ? " current" : ""}`, role: "option", "aria-selected": String(chosen.has(s.tag)) }, box, h("span", { class: "source-label" }, label(s)));
        item.addEventListener("click", () => {
          if (chosen.has(s.tag)) chosen.delete(s.tag);
          else chosen.add(s.tag);
          current = i;
          render();
        });
        return item;
      }),
    );
    list.querySelector(".current")?.scrollIntoView({ block: "nearest" });
    const tags = picked();
    preview.textContent = tags.length ? `Shows as ${formatCitation({ tags, page: page.value.trim() }, bib, [...new Set([...citedOrder(view.state.doc), ...tags])])}` : "";
  };
  const picked = () => (chosen.size ? bib.sources.filter((s) => chosen.has(s.tag)).map((s) => s.tag) : visible()[current] ? [visible()[current].tag] : []);
  const ok = () => {
    const tags = picked();
    if (!tags.length) return;
    dialog.close();
    const field = schema.nodes.field.create({ instr: citationInstr({ tags, page: page.value.trim() }), result: "" });
    const tr = view.state.tr.replaceSelectionWith(field, true);
    // Straight after a word, a citation gets a space before it: "water (Smith, 2020)".
    const at = tr.mapping.map(view.state.selection.from, -1);
    const before = tr.doc.textBetween(Math.max(tr.doc.resolve(at).start(), at - 1), at);
    if (before && !/\s/.test(before)) tr.insertText(" ", at);
    view.dispatch(tr.setMeta(citationsKey, true).scrollIntoView());
    view.focus();
  };
  search.addEventListener("input", () => {
    current = 0;
    render();
  });
  page.addEventListener("input", render);
  const keys = (event: KeyboardEvent) => {
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      current += event.key === "ArrowDown" ? 1 : -1;
      render();
    } else if (event.key === " " && event.target === search && event.ctrlKey) {
      event.preventDefault();
      const s = visible()[current];
      if (s) chosen.has(s.tag) ? chosen.delete(s.tag) : chosen.add(s.tag);
      render();
    } else if (event.key === "Enter") {
      event.preventDefault();
      ok();
    }
  };
  search.addEventListener("keydown", keys);
  page.addEventListener("keydown", keys);
  const manage = h("button", { type: "button", class: "btn" }, "Sources…");
  manage.addEventListener("click", () => {
    dialog.close();
    manageSources(view, ctx, () => insertCitation(view, ctx));
  });
  const dialog = modal(
    "Cite",
    search,
    list,
    h("div", { class: "dialog-form" }, h("label", {}, "Page (optional)", page)),
    preview,
    h("p", { class: "dialog-hint" }, "Arrows choose, Enter cites. Tick several (click, or Ctrl+Space) to cite them together."),
    h(
      "div",
      { class: "dialog-footer" },
      manage,
      h("span", { class: "dialog-spacer" }),
      h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"),
      h("button", { type: "button", class: "btn primary", onclick: ok }, "Cite"),
    ),
  );
  render();
  search.focus();
}

/** Puts the reference list after the paragraph the cursor is in, or goes to the one there is (Alt+Shift+B). */
export function insertBibliography(view: EditorView, ctx: DocContext) {
  let existing = -1;
  view.state.doc.forEach((node, pos) => {
    if (node.type === schema.nodes.bibliography && existing < 0) existing = pos;
  });
  if (existing >= 0) {
    view.dispatch(view.state.tr.setSelection(TextSelection.near(view.state.doc.resolve(existing))).scrollIntoView());
    view.focus();
    return;
  }
  ensureStyles(ctx, "Bibliography");
  const { $from } = view.state.selection;
  const at = $from.depth ? $from.after(1) : view.state.doc.content.size;
  const tr = view.state.tr.insert(at, schema.nodes.bibliography.create());
  // Text can go on after it.
  if (at + 1 >= tr.doc.content.size) tr.insert(tr.doc.content.size, schema.nodes.paragraph.create());
  view.dispatch(tr.setMeta(citationsKey, true).scrollIntoView());
  view.focus();
}
