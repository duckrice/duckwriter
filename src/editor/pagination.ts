// Real pages. The text is laid out as one column; after every change it is
// measured, and gaps are put where a page ends so the text continues at the
// top of the next page. Behind the text, page sheets are drawn at the same
// spacing. Paragraphs split between lines without leaving a single line
// behind or alone at the top (widow and orphan control), headings stay with
// what follows them, and page breaks start a new page.

import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, type EditorState } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { TableMap } from "prosemirror-tables";
import { resolvedPara, type DocContext } from "../docx/model";
import { schema } from "./schema";

/** Space between pages on screen, in pixels; none when printing. */
export let PAGE_GAP = 24;

export function setPageGap(px: number) {
  PAGE_GAP = px;
}
const PX = 4 / 3;

export interface PageLayout {
  /** Number of pages. */
  count: number;
  /** The footnotes at the bottom of each page, by id. */
  notes: string[][];
  /** The page (from 0) each top-level paragraph starts on, by its position. */
  pageOfBlock: Map<number, number>;
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
  const page = ctx.page;
  const contentHeight = (page.height - page.marginTop - page.marginBottom) * PX;
  const pitch = page.height * PX + PAGE_GAP;
  // A little leeway: far down a long document, a line moved to the top of a
  // page can end up a rounding error above it, and would be moved again forever.
  const pageOf = (y: number) => Math.max(0, Math.floor((y + 0.01) / pitch));
  const contentTop = (k: number) => k * pitch;
  const contentBottom = (k: number) => k * pitch + contentHeight;

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
  const props = (node: Node) => (node.type === schema.nodes.paragraph ? resolvedPara(ctx.styles, node.attrs.style) : {});

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
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    const top = () => b.top + shift;
    let k = pageOf(top());
    // Starting in the space between pages, or asked to start a page.
    if (top() > contentBottom(k) - 0.5 || (props(b.node).pageBreakBefore && top() > contentTop(k) + 1)) {
      add(b.pos, top(), false);
      k = pageOf(top());
    }

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
        const repeat = header.length && headerHeight < contentHeight / 2 ? header : undefined;
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
  const end = last ? last.bottom + shift : 0;
  const count = Math.max(1, pageOf(Math.max(0, end - 0.5)) + 1);
  gaps.sort((a, b) => a.pos - b.pos || Number(a.inline) - Number(b.inline));
  view.dispatch(view.state.tr.setMeta(paginationKey, DecorationSet.create(view.state.doc, gaps.flatMap(gapWidget))).setMeta("addToHistory", false));
  sheet.style.minHeight = "";
  // Every gap before a block moved it down; its final top is its first line's page.
  let moved = 0;
  const gapsBefore = [...gaps].sort((a, b) => a.pos - b.pos);
  let g = 0;
  for (const b of blocks) {
    while (g < gapsBefore.length && gapsBefore[g].pos <= b.pos) moved += gapsBefore[g++].height;
    pageOfBlock.set(b.pos, pageOf(b.top + moved + 1));
  }
  hooks.onLayout({ count, notes: Array.from({ length: count }, (_, k) => notesOn[k] ?? []), pageOfBlock });
}
