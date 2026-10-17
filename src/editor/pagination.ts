// Real pages. The text is laid out as one column; after every change it is
// measured, and gaps are put where a page ends so the text continues at the
// top of the next page. Behind the text, page sheets are drawn at the same
// spacing. Paragraphs split between lines without leaving a single line
// behind or alone at the top (widow and orphan control), headings stay with
// what follows them, and page breaks start a new page.

import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, type EditorState } from "prosemirror-state";
import { Decoration, DecorationSet, type EditorView } from "prosemirror-view";
import { resolvedPara, type DocContext } from "../docx/model";
import { schema } from "./schema";

/** Space between pages on screen, in pixels. */
export const PAGE_GAP = 24;
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
}

/** The line above the footnotes, with the space around it, in pixels. */
export const NOTE_SEPARATOR = 16;

interface Gap {
  pos: number;
  height: number;
  /** Between two lines of a paragraph, rather than between paragraphs. */
  inline: boolean;
}

export const paginationKey = new PluginKey<DecorationSet>("pagination");
const schedulers = new WeakMap<EditorView, () => void>();

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
      const schedule = () => {
        if (scheduled) return;
        scheduled = true;
        // Before the next paint, so the text never shows without its page breaks.
        queueMicrotask(() => {
          scheduled = false;
          if (view.isDestroyed) return;
          layout(view, hooks);
          laidOut = view.state.doc;
        });
      };
      schedulers.set(view, schedule);
      schedule();
      // Fonts may arrive after the first layout and change line breaks.
      void document.fonts?.ready.then(schedule);
      return {
        update(v: EditorView, prev: EditorState) {
          if (v.state.doc !== prev.doc && v.state.doc !== laidOut) schedule();
        },
      };
    },
  });
}

/** Lays the pages out again, e.g. after the page size changed. */
export function relayout(view: EditorView) {
  schedulers.get(view)?.();
}

function gapWidget(gap: Gap): Decoration {
  return Decoration.widget(
    gap.pos,
    () => {
      const el = document.createElement(gap.inline ? "span" : "div");
      el.className = gap.inline ? "page-gap inline" : "page-gap";
      el.style.height = `${gap.height}px`;
      el.contentEditable = "false";
      return el;
    },
    { side: -1, key: `gap:${gap.inline ? "i" : "b"}:${Math.round(gap.height * 10)}`, ignoreSelection: true, marks: [] },
  );
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
  const pageOf = (y: number) => Math.max(0, Math.floor(y / pitch));
  const contentTop = (k: number) => k * pitch;
  const contentBottom = (k: number) => k * pitch + contentHeight;

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
    if (lines.length < 2 && !breaks.length) {
      moveWhole();
      placeNotes(pageOf(top()), allNotes);
      continue;
    }

    const lineStart = (line: Line) => {
      const found = view.posAtCoords({ left: line.left + 1, top: origin.top + ((line.top + line.bottom) / 2) * scale });
      if (!found) return null;
      const pos = Math.max(b.pos + 1, Math.min(found.pos, b.pos + b.node.nodeSize - 1));
      return pos > b.pos + 1 ? pos : null;
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
    while (n < lines.length) {
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
  view.dispatch(view.state.tr.setMeta(paginationKey, DecorationSet.create(view.state.doc, gaps.map(gapWidget))).setMeta("addToHistory", false));
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
