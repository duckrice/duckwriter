// Real pages. The text is laid out as one column; after every change it is
// measured, and gaps are put where a page ends so the text continues at the
// top of the next page. Behind the text, page sheets are drawn at the same
// spacing. Paragraphs split between lines without leaving a single line
// behind or alone at the top (widow and orphan control), headings stay with
// what follows them, and page breaks start a new page. A section in columns
// has its blocks at the column's width; when a column is full, the next block
// is pulled up to the top of the next column (a negative margin under the
// block before) and moved across.

import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, type EditorState } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { TableMap } from "prosemirror-tables";
import { resolvedPara, type DocContext, type PageSetup, type ParaProps } from "../docx/model";
import { pageFromSectPr } from "../docx/read";
import { schema } from "./schema";

/** Space between pages on screen, in pixels; none when printing. */
export let PAGE_GAP = 24;

export function setPageGap(px: number) {
  PAGE_GAP = px;
}

/**
 * On paper, for a PDF with pages of several sizes: each page keeps its own size, in a slot this
 * tall (pixels) on paper big enough for any of them; the PDF's pages are cropped to size afterwards.
 */
let paperSlot: number | null = null;
export function setPaperSlot(px: number | null) {
  paperSlot = px;
}
const PX = 4 / 3;

export interface PageLayout {
  /** Number of pages. */
  count: number;
  /** The footnotes at the bottom of each page, by id. */
  notes: string[][];
  /** The page (from 0) each top-level paragraph starts on, by its position. */
  pageOfBlock: Map<number, number>;
  /** Each page's number as headers and footers show it (sections can restart it, or count i, ii, iii). */
  labels?: string[];
  /** The sections: the page each starts on, and its properties (w:sectPr). */
  sections?: { firstPage: number; type?: string; xml: string }[];
  /** Each page: where it starts (pixels from the top of the first) and its size and margins. */
  pages?: { top: number; setup: PageSetup }[];
}

export interface PaginationHooks {
  context(): DocContext;
  /** Called after every layout. */
  onLayout(layout: PageLayout): void;
  /** How tall a footnote is at the bottom of a page, in pixels. */
  noteHeight?(id: string): number;
  /** Measures these footnotes ahead, so that noteHeight changes nothing on the page. */
  measureNotes?(ids: string[]): void;
}

/** The line above the footnotes, with the space around it, in pixels. */
export const NOTE_SEPARATOR = 16;

interface Gap {
  pos: number;
  height: number;
  /** Between two lines of a paragraph, rather than between paragraphs. */
  inline: boolean;
  /** Between two rows of a table with this many columns. */
  columns?: number;
  /** The table's header rows, repeated at the top of the next page (copies, not editable). */
  header?: HTMLElement[];
}

export const paginationKey = new PluginKey<DecorationSet>("pagination");
const schedulers = new WeakMap<EditorView, () => void>();
/** How many layouts ran, and how long they took in all. */
const layouts = ((globalThis as { __layoutStats?: { runs: number; ms: number } }).__layoutStats = { runs: 0, ms: 0 });

export function paginationPlugin(hooks: PaginationHooks) {
  return new Plugin<DecorationSet>({
    key: paginationKey,
    state: {
      init: () => DecorationSet.empty,
      apply(tr, old) {
        const meta = tr.getMeta(paginationKey) as DecorationSet | undefined;
        return meta ?? old.map(tr.mapping, tr.doc);
      },
    },
    props: { decorations: (state) => paginationKey.getState(state) },
    view(view) {
      let scheduled = false;
      let laidOut: Node | null = null;
      let took = 0;
      let waiting: ReturnType<typeof setTimeout> | undefined;
      const run = () => {
        scheduled = false;
        clearTimeout(waiting);
        waiting = undefined;
        if (view.isDestroyed) return;
        const started = performance.now();
        layout(view, hooks);
        took = performance.now() - started;
        laidOut = view.state.doc;
        // For scripts/profile.mjs.
        layouts.runs++;
        layouts.ms += took;
      };
      const schedule = () => {
        if (scheduled) return;
        scheduled = true;
        // Before the next paint, so the text never shows without its page breaks.
        queueMicrotask(run);
      };
      /**
       * After typing. In a long document, where laying out the pages takes a
       * while, that waits until typing pauses; until then the page breaks stay
       * where they were, moved along with the text.
       */
      const changed = () => {
        if (took < 30) return schedule();
        if (scheduled && !waiting) return;
        scheduled = true;
        clearTimeout(waiting);
        waiting = setTimeout(run, 250);
      };
      schedulers.set(view, () => {
        clearTimeout(waiting);
        waiting = undefined;
        scheduled = false;
        schedule();
      });
      schedule();
      // Fonts may arrive after the first layout and change line breaks.
      void document.fonts?.ready.then(schedule);
      return {
        update(v: EditorView, prev: EditorState) {
          if (v.state.doc !== prev.doc && v.state.doc !== laidOut) changed();
        },
        destroy() {
          clearTimeout(waiting);
        },
      };
    },
  });
}

/** Lays the pages out again, e.g. after the page size changed. */
export function relayout(view: EditorView) {
  schedulers.get(view)?.();
}

function gapWidget(gap: Gap): Decoration[] {
  if (gap.columns) {
    // In a table the gap is an empty row, without lines, then the header rows again.
    const header = gap.header ?? [];
    const sign = header.map((r) => r.textContent).join("|");
    return [
      Decoration.widget(
        gap.pos,
        () => {
          const row = document.createElement("tr");
          row.className = "page-gap-row";
          row.contentEditable = "false";
          const cell = document.createElement("td");
          cell.colSpan = gap.columns!;
          cell.style.height = `${gap.height}px`;
          row.append(cell);
          return row;
        },
        { side: -(header.length + 1), key: `gap:r:${Math.round(gap.height * 10)}`, ignoreSelection: true },
      ),
      ...header.map((source, i) =>
        Decoration.widget(
          gap.pos,
          () => {
            const row = source.cloneNode(true) as HTMLElement;
            row.classList.add("repeated-header");
            row.contentEditable = "false";
            row.setAttribute("aria-hidden", "true");
            for (const el of row.querySelectorAll(".selectedCell, .column-resize-handle")) el.classList.remove("selectedCell");
            for (const el of row.querySelectorAll(".column-resize-handle")) el.remove();
            return row;
          },
          { side: -(header.length - i), key: `hdr:${i}:${sign}`, ignoreSelection: true },
        ),
      ),
    ];
  }
  return [Decoration.widget(
    gap.pos,
    () => {
      const el = document.createElement(gap.inline ? "span" : "div");
      el.className = gap.inline ? "page-gap inline" : "page-gap";
      el.style.height = `${gap.height}px`;
      el.contentEditable = "false";
      return el;
    },
    { side: -1, key: `gap:${gap.inline ? "i" : "b"}:${Math.round(gap.height * 10)}`, ignoreSelection: true, marks: [] },
  )];
}

interface Line {
  top: number;
  bottom: number;
  left: number;
}

/** The lines of a paragraph, merged where they overlap (tall equations, big letters). */
function linesOf(el: HTMLElement, originTop: number, scale: number): Line[] {
  const range = document.createRange();
  range.selectNodeContents(el);
  const rects = [...range.getClientRects()].filter((r) => r.height >= 2 && r.width > 0);
  rects.sort((a, b) => a.top - b.top);
  const lines: Line[] = [];
  for (const r of rects) {
    const top = (r.top - originTop) / scale;
    const bottom = (r.bottom - originTop) / scale;
    const last = lines[lines.length - 1];
    if (last && top < last.bottom - 1) {
      last.bottom = Math.max(last.bottom, bottom);
      last.left = Math.min(last.left, r.left);
    } else lines.push({ top, bottom, left: r.left });
  }
  return lines;
}

function layout(view: EditorView, hooks: PaginationHooks) {
  const ctx = hooks.context();

  // Each section's page (size, orientation, margins): the last section's is the document's.
  const sectionOf: number[] = [];
  const setups: PageSetup[] = [];
  // How each section starts (next page, odd page, continuous…): Word keeps that in the section's own
  // properties, which come at its end, so it's known here before the layout reaches the section.
  const starts: string[] = [];
  const columns: (Columns | null)[] = [];
  const startOf = (xml: string) => /<w:type\b[^>]*w:val="(\w+)"/.exec(xml)?.[1] ?? "nextPage";
  {
    let s = 0;
    view.state.doc.forEach((node) => {
      sectionOf.push(s);
      const end = sectionBreak(node);
      if (end) {
        // On paper (PDF, printing) every page has the document's size: one paper size per print.
        setups[s] = ((!onPaper() || paperSlot !== null) && pageFromSectPr(end.xml)) || ctx.page;
        starts[s] = startOf(end.xml);
        columns[s] = columnsOf(end.xml, textWidth(setups[s]));
        s++;
      }
    });
    setups[s] = ctx.page;
    starts[s] = startOf(ctx.sectPr ?? "");
    columns[s] = columnsOf(ctx.sectPr ?? "", textWidth(ctx.page));
  }

  // The pages, made as the text reaches them; a section that starts a page sets that page's size.
  // y is measured from the top of the text on the first page (the document's top margin).
  const pageSetups: PageSetup[] = [];
  const pageTops: number[] = [];
  let setup = setups[0];
  const origin0 = ctx.page.marginTop * PX;
  const ensure = (k: number) => {
    while (pageSetups.length <= k) {
      const n = pageSetups.length;
      pageTops.push(n === 0 ? 0 : pageTops[n - 1] + (paperSlot ?? pageSetups[n - 1].height * PX) + PAGE_GAP);
      pageSetups.push(setup);
    }
  };
  /** From page k on, pages have this setup. */
  const startPages = (k: number, next: PageSetup) => {
    ensure(k);
    pageSetups.length = k + 1;
    pageTops.length = k + 1;
    pageSetups[k] = next;
    setup = next;
  };
  const contentTop = (k: number) => {
    ensure(k);
    return pageTops[k] + pageSetups[k].marginTop * PX - origin0;
  };
  // At least half an inch of text on a page, whatever the margins say.
  const contentBottom = (k: number) => contentTop(k) + Math.max(36, pageSetups[k].height - pageSetups[k].marginTop - pageSetups[k].marginBottom) * PX;
  // A little leeway: far down a long document, a line moved to the top of a
  // page can end up a rounding error above it, and would be moved again forever.
  const pageOf = (y: number) => {
    let last = Math.max(0, pageSetups.length - 1);
    while (contentTop(last + 1) <= y + 0.01) last++;
    // The last page whose text starts at or above y.
    let lo = 0;
    let hi = last;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (contentTop(mid) <= y + 0.01) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  };

  if (hooks.measureNotes) {
    const ids: string[] = [];
    view.state.doc.descendants((node) => {
      if (node.type === schema.nodes.footnote) ids.push(node.attrs.id);
      return node.isBlock;
    });
    if (ids.length) hooks.measureNotes(ids);
  }

  // Measure the text without page breaks, keeping the scroll position steady.
  const dom = view.dom as HTMLElement;
  const sheet = dom.parentElement!;
  sheet.style.minHeight = `${sheet.offsetHeight}px`;
  view.dispatch(view.state.tr.setMeta(paginationKey, DecorationSet.empty).setMeta("addToHistory", false));

  const origin = dom.getBoundingClientRect();
  const scale = dom.offsetHeight ? origin.height / dom.offsetHeight : 1;
  const y = (clientY: number) => (clientY - origin.top) / scale;

  const gaps: Gap[] = [];
  let shift = 0;
  const add = (pos: number, at: number, inline: boolean) => {
    const height = contentTop(pageOf(at) + 1) - at;
    gaps.push({ pos, height, inline });
    shift += height;
  };

  const blocks: { node: Node; pos: number; el: HTMLElement; top: number; bottom: number }[] = [];
  view.state.doc.forEach((node, pos) => {
    const el = view.nodeDOM(pos);
    if (!(el instanceof HTMLElement)) return;
    const r = el.getBoundingClientRect();
    if (r.height === 0 && r.width === 0) return;
    blocks.push({ node, pos, el, top: y(r.top), bottom: y(r.bottom) });
  });
  const props = (node: Node): ParaProps => {
    if (node.type !== schema.nodes.paragraph) return {};
    const style = resolvedPara(ctx.styles, node.attrs.style);
    // Set on the paragraph itself, these are kept in its raw properties and outrank its style's.
    const own = (name: string) => {
      const m = new RegExp(`<w:${name}\\b([^>]*)>`).exec(node.attrs.extra ?? "");
      return m ? !/w:val="(0|false|off)"/.test(m[1]) : undefined;
    };
    return { ...style, keepNext: own("keepNext") ?? style.keepNext, pageBreakBefore: own("pageBreakBefore") ?? style.pageBreakBefore };
  };

  // Footnotes take room at the bottom of the page their mark is on.
  const reserved: number[] = [];
  const notesOn: string[][] = [];
  const noteHeight = (id: string) => hooks.noteHeight?.(id) ?? 0;
  /** The room footnotes take on page k if `ids` are added there. */
  const noteRoom = (k: number, ids: string[]) => {
    const extra = ids.reduce((sum, id) => sum + noteHeight(id), 0);
    const has = (notesOn[k]?.length ?? 0) > 0;
    return (reserved[k] ?? 0) + extra + (!has && ids.length ? NOTE_SEPARATOR : 0);
  };
  const placeNotes = (k: number, ids: string[]) => {
    if (!ids.length) return;
    reserved[k] = noteRoom(k, ids);
    (notesOn[k] ??= []).push(...ids);
  };
  const unplaceNotes = (k: number, ids: string[]) => {
    if (!ids.length || !notesOn[k]) return;
    notesOn[k] = notesOn[k].filter((id) => !ids.includes(id));
    reserved[k] = notesOn[k].length ? notesOn[k].reduce((sum, id) => sum + noteHeight(id), NOTE_SEPARATOR) : 0;
  };
  const fits = (bottom: number, k: number, ids: string[]) => bottom <= contentBottom(k) - noteRoom(k, ids) + 0.5;

  const pageOfBlock = new Map<number, number>();
  // Sections: where each starts (its first page) and its properties (Word keeps them at its end).
  const sections: { firstPage: number; restarts: boolean; type?: string; xml: string }[] = [{ firstPage: 0, restarts: true, xml: "" }];
  let broke: SectionBreak | null = null;
  // Columns: the section they're in, the column being filled, where the columns on this page start,
  // and how far down the longest reaches. Blocks pulled up to a column's top are listed in pulls.
  let colSection = -1;
  let col = 0;
  let colTop = 0;
  let colMax = 0;
  let columnBreakBefore = false;
  const colDecos: Decoration[] = [];
  const pulls: { pos: number; height: number }[] = [];
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    const top = () => b.top + shift;
    // The first section's top margin may be deeper than the document's (the page's own setup).
    if (i === 0 && contentTop(0) - top() > 0.5) {
      const down = contentTop(0) - top();
      gaps.push({ pos: b.pos, height: down, inline: false });
      shift += down;
    }
    // After a section in columns, the text goes on under the longest column.
    if (colSection >= 0 && sectionOf[i] !== colSection) {
      const under = colMax - top();
      if (under > 0.5) {
        gaps.push({ pos: b.pos, height: under, inline: false });
        shift += under;
      }
      colSection = -1;
    }
    const cols = columns[sectionOf[i]];
    let k = pageOf(top());
    const start = broke ? starts[sectionOf[i]] : null;
    const newSection = start !== null && start !== "continuous";
    if (newSection) {
      // A section that starts a page: on the next page (or this one, if it is at its top), with
      // the section's own page size. An odd- or even-page section skips a page number, as Word
      // shows it (the blank page is only printed), so that is done in pageLabels.
      const at = top() > contentTop(k) + 1 ? k + 1 : k;
      startPages(at, setups[sectionOf[i]] ?? setup);
      const down = contentTop(at) - top();
      if (down > 0.5) {
        gaps.push({ pos: b.pos, height: down, inline: false });
        shift += down;
      }
      k = at;
    } else if ((!cols && top() > contentBottom(k) - 0.5) || (props(b.node).pageBreakBefore && top() > contentTop(k) + 1)) {
      // Starting in the space between pages, or asked to start a page.
      add(b.pos, top(), false);
      k = pageOf(top());
    }
    if (broke) sections.push({ firstPage: k, restarts: start !== "continuous", type: start ?? undefined, xml: "" });
    broke = sectionBreak(b.node);
    if (broke) sections[sections.length - 1].xml = broke.xml;

    const breaks: { pos: number; bottom: number }[] = [];
    const notes: { id: string; y: number }[] = [];
    if (b.node.type === schema.nodes.paragraph) {
      b.node.forEach((child, offset) => {
        if (child.type !== schema.nodes.page_break && child.type !== schema.nodes.footnote) return;
        const el = view.nodeDOM(b.pos + 1 + offset);
        if (!(el instanceof HTMLElement)) return;
        const r = el.getBoundingClientRect();
        if (child.type === schema.nodes.page_break) breaks.push({ pos: b.pos + 1 + offset + child.nodeSize, bottom: y(r.bottom) });
        else notes.push({ id: child.attrs.id, y: y((r.top + r.bottom) / 2) });
      });
    }
    const allNotes = notes.map((n) => n.id);
    if (cols) {
      if (colSection !== sectionOf[i] || k !== pageOf(colTop)) {
        // The section's first block, or one a page break sent on: the first column of this page.
        colSection = sectionOf[i];
        col = 0;
        colTop = top();
        colMax = top();
      }
      const full = b.bottom + shift > contentBottom(k) - noteRoom(k, allNotes) + 0.5;
      // shortcut: a paragraph moves to the next column whole; Word splits it between lines.
      if ((full || columnBreakBefore) && top() > colTop + 1) {
        colMax = Math.max(colMax, top());
        const prev = blocks[i - 1];
        if (col + 1 < cols.widths.length && prev) {
          const up = colTop - top();
          colDecos.push(Decoration.node(prev.pos, prev.pos + prev.node.nodeSize, { style: `margin-bottom:${up}px` }));
          pulls.push({ pos: b.pos, height: up });
          shift += up;
          col++;
        } else {
          add(b.pos, top(), false);
          k = pageOf(top());
          col = 0;
          colTop = top();
          colMax = top();
        }
      }
      columnBreakBefore = hasColumnBreak(b.node);
      if (col > 0) {
        let x = 0;
        for (let c = 0; c < col; c++) x += cols.widths[c] + cols.gaps[c];
        // shortcut: the line between columns (w:sep) is kept in the file but not drawn.
        colDecos.push(Decoration.node(b.pos, b.pos + b.node.nodeSize, { style: `position:relative;left:${x}pt` }));
      }
      colMax = Math.max(colMax, b.bottom + shift);
      // Taller than a column, or with page breaks: split between pages as any block is, below.
      if (!breaks.length && b.bottom + shift <= contentBottom(k) + 0.5) {
        placeNotes(k, allNotes);
        continue;
      }
    }
    if (!breaks.length && fits(b.bottom + shift, k, allNotes)) {
      placeNotes(k, allNotes);
      continue;
    }

    const lines = b.node.type === schema.nodes.paragraph ? linesOf(b.el, origin.top, scale) : [];
    const moveWhole = () => {
      if (top() <= contentTop(pageOf(top())) + 1) return false;
      // A heading just before goes along, so it isn't left alone at the bottom of the page.
      const prev = blocks[i - 1];
      const prevTop = prev ? prev.top + shift : 0;
      if (prev && props(prev.node).keepNext && !gaps.some((g) => g.pos === prev.pos) && pageOf(prevTop) === pageOf(top()) && prevTop > contentTop(pageOf(prevTop)) + 1) {
        add(prev.pos, prevTop, false);
      } else add(b.pos, top(), false);
      return true;
    };
    if (b.node.type === schema.nodes.table) {
      // A table continues on the next page between two rows, with its header rows again on top.
      const columns = TableMap.get(b.node).width;
      const header: HTMLElement[] = [];
      let headerHeight = 0;
      b.node.forEach((row, offset, r) => {
        if (header.length !== r || row.firstChild?.type !== schema.nodes.table_header) return;
        const el = view.nodeDOM(b.pos + 1 + offset);
        if (!(el instanceof HTMLElement)) return;
        header.push(el.cloneNode(true) as HTMLElement);
        headerHeight += el.getBoundingClientRect().height / scale;
      });
      b.node.forEach((row, offset, r) => {
        const rowPos = b.pos + 1 + offset;
        const el = view.nodeDOM(rowPos);
        if (!(el instanceof HTMLElement)) return;
        const rect = el.getBoundingClientRect();
        const rowTop = y(rect.top) + shift;
        const rowBottom = y(rect.bottom) + shift;
        const k = pageOf(rowTop);
        if (fits(rowBottom, k, []) || rowTop <= contentTop(k) + 1) return;
        if (r === 0 || r < header.length) {
          if (r === 0) moveWhole();
          return;
        }
        const height = contentTop(k + 1) - rowTop;
        // Repeating the header makes no sense if it would fill the page.
        const repeat = header.length && headerHeight < (contentBottom(k + 1) - contentTop(k + 1)) / 2 ? header : undefined;
        gaps.push({ pos: rowPos, height, inline: false, columns, header: repeat });
        shift += height + (repeat ? headerHeight : 0);
        void row;
      });
      continue;
    }
    if (lines.length < 2 && !breaks.length) {
      moveWhole();
      placeNotes(pageOf(top()), allNotes);
      continue;
    }

    // The first position on a line, by halving: positions further on are never
    // on a line higher up. (posAtCoords searches the whole document when the
    // line is outside the window, which in a long document takes minutes.)
    const lineStart = (line: Line) => {
      const below = (pos: number) => {
        try {
          const c = view.coordsAtPos(pos, 1);
          return y((c.top + c.bottom) / 2) >= line.top;
        } catch {
          return false;
        }
      };
      let lo = b.pos + 1;
      let hi = b.pos + b.node.nodeSize - 1;
      if (!below(hi)) return null;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (below(mid)) hi = mid;
        else lo = mid + 1;
      }
      return lo > b.pos + 1 ? lo : null;
    };
    const notesIn = (line: Line) => notes.filter((n) => n.y >= line.top - 1 && n.y <= line.bottom + 1).map((n) => n.id);
    /** Where each line's notes were placed, to take them back when the line moves on. */
    const placed: { k: number; ids: string[] }[] = [];
    const takeBack = (from: number) => {
      for (let m = from; m < placed.length; m++) if (placed[m]) unplaceNotes(placed[m].k, placed[m].ids);
      placed.length = from;
    };
    let handledBreak = 0;
    let movedWhole = false;
    let n = 0;
    // Each line moves on at most once or twice; more means something is wrong.
    let rounds = lines.length * 4 + 8;
    while (n < lines.length && rounds-- > 0) {
      const line = lines[n];
      // Page breaks above this line send it (and what follows) to a new page.
      while (handledBreak < breaks.length && breaks[handledBreak].bottom <= line.top + 0.5) {
        add(breaks[handledBreak].pos, breaks[handledBreak].bottom + shift, true);
        handledBreak++;
      }
      const lineTop = line.top + shift;
      k = pageOf(lineTop);
      const ids = notesIn(line);
      // It fits, or it is already at the top of a page and can't do better.
      if (fits(line.bottom + shift, k, ids) || lineTop <= contentTop(k) + 1) {
        placeNotes(k, ids);
        placed[n] = { k, ids };
        n++;
        continue;
      }
      // The first line doesn't fit, or only one line would stay behind: the paragraph moves.
      if (n <= 1 && !movedWhole) {
        movedWhole = true;
        takeBack(0);
        if (moveWhole()) {
          n = 0;
          continue;
        }
      }
      // Don't leave the last line alone at the top of the next page.
      const from = n === lines.length - 1 && n >= 3 ? n - 1 : n;
      const pos = lineStart(lines[from]);
      if (pos === null) {
        placeNotes(k, ids);
        placed[n] = { k, ids };
        n++;
        continue;
      }
      takeBack(from);
      add(pos, lines[from].top + shift, true);
      n = from;
    }
    for (; handledBreak < breaks.length; handledBreak++) add(breaks[handledBreak].pos, breaks[handledBreak].bottom + shift, true);
  }

  const last = blocks[blocks.length - 1];
  const end = Math.max(last ? last.bottom + shift : 0, colSection >= 0 ? colMax : 0);
  const count = Math.max(1, pageOf(Math.max(0, end - 0.5)) + 1);
  gaps.sort((a, b) => a.pos - b.pos || Number(a.inline) - Number(b.inline));
  view.dispatch(view.state.tr.setMeta(paginationKey, DecorationSet.create(view.state.doc, [...gaps.flatMap(gapWidget), ...colDecos])).setMeta("addToHistory", false));
  sheet.style.minHeight = "";
  // Every gap before a block moved it down; its final top is its first line's page.
  let moved = 0;
  const gapsBefore = [...gaps, ...pulls].sort((a, b) => a.pos - b.pos);
  let g = 0;
  for (const b of blocks) {
    while (g < gapsBefore.length && gapsBefore[g].pos <= b.pos) moved += gapsBefore[g++].height;
    pageOfBlock.set(b.pos, pageOf(b.top + moved + 1));
  }
  sections[sections.length - 1].xml = ctx.sectPr ?? "";
  hooks.onLayout({ count, notes: Array.from({ length: count }, (_, k) => notesOn[k] ?? []), pageOfBlock, labels: pageLabels(count, sections), sections, pages: Array.from({ length: count }, (_, k) => (ensure(k), { top: pageTops[k], setup: pageSetups[k] })) });
}

/**
 * A section whose text is wider or set further in than the document's (a
 * landscape table): its paragraphs and tables take its own width, by margins
 * on each block. (Indents are padding, so these don't clash with them.)
 */
export function sectionWidthPlugin(context: () => DocContext) {
  let cached: { doc: Node; page: string; set: DecorationSet } | null = null;
  const setups = new Map<string, PageSetup | null>();
  return new Plugin({
    props: {
      decorations(state) {
        const main = context().page;
        const last = context().sectPr ?? "";
        const key = JSON.stringify(main) + onPaper() + paperSlot + last;
        if (cached && cached.doc === state.doc && cached.page === key) return cached.set;
        // Each section's blocks, with the setup found at the section's end.
        const decorations: Decoration[] = [];
        let blocks: { pos: number; end: number }[] = [];
        const section = (xml: string, own: boolean) => {
          if (own && !setups.has(xml)) setups.set(xml, pageFromSectPr(xml));
          // On paper every page has the document's size (see layout).
          const p = own && (!onPaper() || paperSlot !== null) ? setups.get(xml) : null;
          // A page of another width is centred on the document's, as Word shows it.
          const shift = p ? pageOffset(p, main) : 0;
          const left = p ? shift + p.marginLeft - main.marginLeft : 0;
          const right = p ? main.width - main.marginRight - (shift + p.width - p.marginRight) : 0;
          // In columns, each block is a column wide (the layout moves it across).
          // shortcut: unequal columns are all as wide as the first; Word gives each its own width.
          const cols = columnsOf(xml, textWidth(p ?? main));
          const narrower = cols ? textWidth(p ?? main) - cols.widths[0] : 0;
          if (Math.abs(left) > 0.5 || Math.abs(right + narrower) > 0.5) {
            for (const b of blocks) decorations.push(Decoration.node(b.pos, b.end, { style: `margin-left:${left}pt;margin-right:${right + narrower}pt` }));
          }
          blocks = [];
        };
        state.doc.forEach((node, pos) => {
          blocks.push({ pos, end: pos + node.nodeSize });
          const end = sectionBreak(node);
          if (end) section(end.xml, true);
        });
        section(last, false);
        const set = DecorationSet.create(state.doc, decorations);
        cached = { doc: state.doc, page: key, set };
        return set;
      },
    },
  });
}

/** Laying out for paper (PDF, printing), where there are no gaps between pages. */
const onPaper = () => PAGE_GAP === 0;

/** How far a page of this size sits left of the document's pages (points; negative: further left), centred on them. */
export const pageOffset = (setup: PageSetup, main: PageSetup) => (main.width - setup.width) / 2;

/** Page k: where it starts (pixels from the top of the first page) and its size and margins. */
export function pageBox(layout: PageLayout | null | undefined, k: number, page: PageSetup): { top: number; setup: PageSetup } {
  return layout?.pages?.[k] ?? { top: k * (page.height * PX + PAGE_GAP), setup: page };
}

/** The page at a height (pixels from the top of the first page). */
export function pageAt(layout: PageLayout | null | undefined, y: number, page: PageSetup): number {
  const pages = layout?.pages;
  if (!pages?.length) return Math.max(0, Math.floor(y / (page.height * PX + PAGE_GAP)));
  let k = 0;
  while (k + 1 < pages.length && pages[k + 1].top <= y) k++;
  return k;
}

/** A section's columns, in points: each one's width, the space after each but the last, and a line between. */
export interface Columns {
  widths: number[];
  gaps: number[];
  line: boolean;
}

/** The columns a section's properties (w:sectPr) ask for, on text this wide (points); null for one. */
export function columnsOf(xml: string, textWidth: number): Columns | null {
  const m = /<w:cols\b([^>]*?)(?:\/>|>([\s\S]*?)<\/w:cols>)/.exec(xml);
  if (!m) return null;
  const attr = (s: string, name: string) => new RegExp(`w:${name}="([^"]*)"`).exec(s)?.[1];
  const line = /^(1|true|on)$/.test(attr(m[1], "sep") ?? "");
  const own = [...(m[2] ?? "").matchAll(/<w:col\b([^>]*?)\/?>/g)].map((c) => ({ w: Number(attr(c[1], "w") ?? 0) / 20, space: Number(attr(c[1], "space") ?? 0) / 20 }));
  if (attr(m[1], "equalWidth") === "0" && own.length > 1 && own.every((c) => c.w > 0)) return { widths: own.map((c) => c.w), gaps: own.slice(0, -1).map((c) => c.space), line };
  const num = Math.min(45, Math.round(Number(attr(m[1], "num") ?? 1)));
  if (!(num >= 2)) return null;
  // Word's default space is half an inch.
  const space = Number(attr(m[1], "space") ?? 720) / 20;
  const width = (textWidth - space * (num - 1)) / num;
  return width >= 18 ? { widths: Array(num).fill(width), gaps: Array(num - 1).fill(space), line } : null;
}

/** Whether a paragraph has a column break (kept as Word wrote it) in it. */
function hasColumnBreak(node: Node): boolean {
  let found = false;
  node.forEach((child) => {
    if (child.type === schema.nodes.raw_inline && /<w:br\b[^>]*w:type="column"/.test(child.attrs.xml ?? "")) found = true;
  });
  return found;
}

const textWidth = (p: PageSetup) => p.width - p.marginLeft - p.marginRight;

interface SectionBreak {
  type: string;
  /** The section's properties (w:sectPr), as kept in the paragraph that ends it. */
  xml: string;
}

/** A paragraph that ends a section: Word keeps the section's properties in its own. */
export function sectionBreak(node: Node): SectionBreak | null {
  if (node.type !== schema.nodes.paragraph) return null;
  const xml = /<w:sectPr\b[\s\S]*?<\/w:sectPr>|<w:sectPr\b[^>]*\/>/.exec(node.attrs.extra ?? "")?.[0];
  if (!xml) return null;
  return { type: /<w:type\b[^>]*w:val="(\w+)"/.exec(xml)?.[1] ?? "nextPage", xml };
}

const roman = (n: number) => {
  const map: [number, string][] = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
  let out = "";
  for (const [v, s] of map) while (n >= v) (out += s), (n -= v);
  return out;
};
const letters = (n: number) => String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1);

/** A page number as a section's numbering shows it: 3, iii, III, c or C. */
export function formatPageNumber(n: number, format: string): string {
  if (n < 1) return String(n);
  switch (format) {
    case "lowerRoman":
      return roman(n);
    case "upperRoman":
      return roman(n).toUpperCase();
    case "lowerLetter":
      return letters(n);
    case "upperLetter":
      return letters(n).toUpperCase();
    default:
      return String(n);
  }
}

/** Each page's number as shown in headers and footers: numbering restarts or changes format by section. */
export function pageLabels(count: number, sections: { firstPage: number; restarts: boolean; type?: string; xml: string }[]): string[] {
  const labels: string[] = [];
  let n = 0;
  let s = 0;
  let format = "decimal";
  for (let k = 0; k < count; k++) {
    while (s + 1 < sections.length && sections[s + 1].firstPage <= k) s++;
    const section = sections[s];
    const numType = /<w:pgNumType\b[^>]*>/.exec(section.xml)?.[0] ?? "";
    if (k === section.firstPage || k === 0) {
      format = /w:fmt="(\w+)"/.exec(numType)?.[1] ?? "decimal";
      const start = /w:start="(\d+)"/.exec(numType)?.[1];
      n = start && (section.restarts || k === 0) ? Number(start) : n + 1;
      if ((section.type === "oddPage" && n % 2 === 0) || (section.type === "evenPage" && n % 2 === 1)) n++;
    } else n++;
    labels.push(formatPageNumber(n, format));
  }
  return labels;
}
