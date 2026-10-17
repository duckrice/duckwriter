// Find and replace in the text. Every match is highlighted, the current one
// more strongly; matches may run across bold, italic and other formatting,
// but not across paragraphs or table cells.

import type { Node } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type EditorState, type Transaction } from "prosemirror-state";
import { Decoration, DecorationSet } from "prosemirror-view";

export interface FindOptions {
  caseSensitive: boolean;
  wholeWord: boolean;
}

export interface Match {
  from: number;
  to: number;
}

interface FindState {
  query: string;
  options: FindOptions;
  matches: Match[];
  /** Index of the current match, or -1. */
  current: number;
  decorations: DecorationSet;
}

export const findKey = new PluginKey<FindState>("find");

const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

/** Every place the query occurs in the document. */
export function findMatches(doc: Node, query: string, options: FindOptions): Match[] {
  if (!query) return [];
  let pattern = escapeRegExp(query);
  // A whole word: no letter or digit just before or after.
  if (options.wholeWord) pattern = `(?<![\\p{L}\\p{N}_])${pattern}(?![\\p{L}\\p{N}_])`;
  const re = new RegExp(pattern, options.caseSensitive ? "gu" : "giu");
  const out: Match[] = [];
  doc.descendants((block, blockPos) => {
    if (!block.isTextblock) return true;
    // The paragraph's text, with each character's place in the document.
    let text = "";
    const positions: number[] = [];
    block.forEach((child, offset) => {
      const start = blockPos + 1 + offset;
      if (child.isText) {
        for (let i = 0; i < child.text!.length; i++) positions.push(start + i);
        text += child.text;
      } else {
        // Pictures, equations, fields: not text, and a match can't run across them.
        positions.push(start);
        text += "￼";
      }
    });
    for (const m of text.matchAll(re)) {
      if (!m[0].length) continue;
      const from = positions[m.index!];
      const to = positions[m.index! + m[0].length - 1] + 1;
      out.push({ from, to });
    }
    return false;
  });
  return out;
}

function decorate(doc: Node, matches: Match[], current: number): DecorationSet {
  return DecorationSet.create(
    doc,
    matches.map((m, i) => Decoration.inline(m.from, m.to, { class: i === current ? "find-match current" : "find-match" })),
  );
}

/** The match at or after `pos`, wrapping round to the first. */
function matchFrom(matches: Match[], pos: number): number {
  if (!matches.length) return -1;
  const i = matches.findIndex((m) => m.from >= pos);
  return i < 0 ? 0 : i;
}

export function findPlugin() {
  return new Plugin<FindState>({
    key: findKey,
    state: {
      init: () => ({ query: "", options: { caseSensitive: false, wholeWord: false }, matches: [], current: -1, decorations: DecorationSet.empty }),
      apply(tr, old, _oldState, state) {
        const meta = tr.getMeta(findKey) as Partial<FindState> | undefined;
        if (!meta && !tr.docChanged) return old;
        const query = meta?.query ?? old.query;
        const options = meta?.options ?? old.options;
        const matches = findMatches(state.doc, query, options);
        let current = meta?.current ?? old.current;
        if (meta?.query !== undefined || meta?.options) current = matchFrom(matches, state.selection.from);
        else if (tr.docChanged) current = matchFrom(matches, state.selection.from);
        if (current >= matches.length) current = matches.length ? 0 : -1;
        return { query, options, matches, current, decorations: decorate(state.doc, matches, current) };
      },
    },
    props: { decorations: (state) => findKey.getState(state)?.decorations },
  });
}

export function findState(state: EditorState): FindState | undefined {
  return findKey.getState(state);
}

/** Looks for something (or stops looking, with ""). */
export function search(state: EditorState, query: string, options: FindOptions): Transaction {
  return state.tr.setMeta(findKey, { query, options });
}

/** Goes to the next match (or the previous one, with -1) and selects it. */
export function step(state: EditorState, direction: 1 | -1): Transaction | null {
  const find = findState(state);
  if (!find || !find.matches.length) return null;
  const n = find.matches.length;
  const sel = state.selection;
  let current: number;
  // On a match: the one after (or before) it. Elsewhere: the first one after (or before) the cursor.
  const on = find.matches.findIndex((m) => m.from === sel.from && m.to === sel.to);
  if (on >= 0) current = (on + direction + n) % n;
  else if (direction > 0) {
    const after = find.matches.findIndex((m) => m.from >= sel.from);
    current = after < 0 ? 0 : after;
  } else {
    let before = -1;
    find.matches.forEach((m, i) => {
      if (m.to <= sel.from) before = i;
    });
    current = before < 0 ? n - 1 : before;
  }
  const match = find.matches[current];
  return state.tr.setSelection(TextSelection.create(state.doc, match.from, match.to)).setMeta(findKey, { current }).scrollIntoView();
}

/** Replaces the current match (keeping the formatting of its first letter) and goes to the next. */
export function replaceCurrent(state: EditorState, replacement: string): Transaction | null {
  const find = findState(state);
  if (!find || find.current < 0) return step(state, 1);
  const match = find.matches[find.current];
  const sel = state.selection;
  // Only once the match has been shown (selected), as in Word.
  if (sel.from !== match.from || sel.to !== match.to) return step(state, 1);
  const marks = state.doc.resolve(match.from + 1).marks();
  const tr = state.tr.replaceWith(match.from, match.to, replacement ? state.schema.text(replacement, marks) : []);
  const after = tr.mapping.map(match.to);
  tr.setSelection(TextSelection.create(tr.doc, after));
  return tr;
}

/** Replaces every match; returns the transaction and how many were replaced. */
export function replaceAll(state: EditorState, replacement: string): { tr: Transaction; count: number } | null {
  const find = findState(state);
  if (!find || !find.matches.length) return null;
  const tr = state.tr;
  // From the end, so earlier positions stay right.
  for (const match of [...find.matches].reverse()) {
    const marks = state.doc.resolve(match.from + 1).marks();
    tr.replaceWith(match.from, match.to, replacement ? state.schema.text(replacement, marks) : []);
  }
  return { tr, count: find.matches.length };
}
