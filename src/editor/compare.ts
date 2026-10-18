// Comparing two versions of a document: the result is the newer one, with
// what changed shown as tracked changes (inserted text underlined, deleted
// text struck through), to go through and accept or reject as usual.
// Paragraphs are matched first; paragraphs that changed are compared word by
// word.

import { Fragment, type Mark, type Node } from "prosemirror-model";
import { schema } from "./schema";

export interface CompareOptions {
  /** Whose changes they are shown as. */
  author: string;
  /** A new change id (Word wants them unique). */
  newId(): string;
}

/** The longest common subsequence of two lists, as pairs of indexes. */
function lcs<T>(a: T[], b: T[], same: (x: T, y: T) => boolean): [number, number][] {
  const n = a.length;
  const m = b.length;
  // Long lists: only the first and last stretches that agree, the middle as changed.
  if (n * m > 4_000_000) {
    const out: [number, number][] = [];
    let i = 0;
    while (i < n && i < m && same(a[i], b[i])) out.push([i, i++]);
    let j = 0;
    const tail: [number, number][] = [];
    while (j < n - i && j < m - i && same(a[n - 1 - j], b[m - 1 - j])) tail.unshift([n - 1 - j, m - 1 - j++]);
    return [...out, ...tail];
  }
  const table = Array.from({ length: n + 1 }, () => new Uint32Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) table[i][j] = same(a[i], b[j]) ? table[i + 1][j + 1] + 1 : Math.max(table[i + 1][j], table[i][j + 1]);
  const out: [number, number][] = [];
  let i = 0;
  let j = 0;
  while (i < n && j < m) {
    if (same(a[i], b[j])) {
      out.push([i, j]);
      i++;
      j++;
    } else if (table[i + 1][j] >= table[i][j + 1]) i++;
    else j++;
  }
  return out;
}

/** A paragraph's text with each character's place in it ("￼" for a picture, an equation…). */
function textOf(para: Node): { text: string; offsets: number[] } {
  let text = "";
  const offsets: number[] = [];
  para.forEach((child, offset) => {
    if (child.isText) {
      for (let i = 0; i < child.text!.length; i++) offsets.push(offset + i);
      text += child.text;
    } else {
      offsets.push(offset);
      text += "￼";
    }
  });
  offsets.push(para.content.size);
  return { text, offsets };
}

/** Words, spaces and punctuation, each a token, with where they start. */
function tokens(text: string): { t: string; at: number }[] {
  const out: { t: string; at: number }[] = [];
  const re = /[\p{L}\p{N}_'’]+|\s+|./gu;
  let m: RegExpExecArray | null;
  while ((m = re.exec(text))) out.push({ t: m[0], at: m.index });
  return out;
}

/** How alike two paragraphs are, from 0 to 1 (shared words). */
function likeness(a: string, b: string): number {
  const wa = a.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [];
  const wb = b.toLowerCase().match(/[\p{L}\p{N}]+/gu) ?? [];
  if (!wa.length && !wb.length) return 1;
  const counts = new Map<string, number>();
  for (const w of wa) counts.set(w, (counts.get(w) ?? 0) + 1);
  let shared = 0;
  for (const w of wb) {
    const c = counts.get(w);
    if (c) {
      shared++;
      counts.set(w, c - 1);
    }
  }
  return (2 * shared) / (wa.length + wb.length);
}

function marked(content: Fragment, mark: Mark, keepNonText: boolean): Node[] {
  const out: Node[] = [];
  content.forEach((n) => {
    if (!n.isText && !keepNonText && n.type !== schema.nodes.hard_break) return;
    out.push(n.mark(mark.addToSet(n.marks)));
  });
  return out;
}

export function compareDocs(original: Node, revised: Node, options: CompareOptions): Node {
  const date = new Date().toISOString().replace(/\.\d+Z$/, "Z");
  const insertion = () => schema.marks.insertion.create({ id: options.newId(), author: options.author, date });
  const deletion = () => schema.marks.deletion.create({ id: options.newId(), author: options.author, date });
  const paraChange = (kind: "ins" | "del") => ({ kind, id: options.newId(), author: options.author, date });

  const a = [...original.content.content];
  const b = [...revised.content.content];
  const key = (n: Node) => (n.type === schema.nodes.paragraph ? `p:${n.textContent}` : `${n.type.name}:${n.textContent}`);
  const anchors = lcs(a, b, (x, y) => key(x) === key(y));
  const out: Node[] = [];

  /** A paragraph that is only in the old version. */
  const removed = (p: Node) => {
    if (p.type !== schema.nodes.paragraph) return;
    out.push(p.type.create({ ...p.attrs, change: paraChange("del") }, marked(p.content, deletion(), false), p.marks));
  };
  /** A paragraph (or table…) that is only in the new version. */
  const added = (p: Node) => {
    if (p.type !== schema.nodes.paragraph) return out.push(p);
    out.push(p.type.create({ ...p.attrs, change: paraChange("ins") }, marked(p.content, insertion(), true), p.marks));
  };
  /** A paragraph in both, changed: word by word. */
  const changed = (old: Node, now: Node) => {
    const ot = textOf(old);
    const nt = textOf(now);
    const ow = tokens(ot.text);
    const nw = tokens(nt.text);
    const pairs = lcs(ow, nw, (x, y) => x.t === y.t);
    const content: Node[] = [];
    const cutOld = (from: number, to: number) => old.content.cut(ot.offsets[from], ot.offsets[to]);
    const cutNew = (from: number, to: number) => now.content.cut(nt.offsets[from], nt.offsets[to]);
    let i = 0;
    let j = 0;
    const flush = (toI: number, toJ: number) => {
      if (toI > i) content.push(...marked(cutOld(ow[i].at, toI < ow.length ? ow[toI].at : ot.text.length), deletion(), false));
      if (toJ > j) content.push(...marked(cutNew(nw[j].at, toJ < nw.length ? nw[toJ].at : nt.text.length), insertion(), true));
    };
    for (const [pi, pj] of [...pairs, [ow.length, nw.length] as [number, number]]) {
      flush(pi, pj);
      if (pi < ow.length && pj < nw.length) cutNew(nw[pj].at, pj + 1 < nw.length ? nw[pj + 1].at : nt.text.length).forEach((n) => content.push(n));
      i = pi + 1;
      j = pj + 1;
    }
    out.push(now.type.create(now.attrs, Fragment.from(content).size ? joinText(content) : null, now.marks));
  };

  let ai = 0;
  let bi = 0;
  for (const [pa, pb] of [...anchors, [a.length, b.length] as [number, number]]) {
    // Between two matches: pair up paragraphs that are much alike, the rest are removed or added.
    const olds = a.slice(ai, pa);
    const news = b.slice(bi, pb);
    let k = 0;
    for (const now of news) {
      const at = olds.findIndex((o, idx) => idx >= k && o.type === schema.nodes.paragraph && now.type === schema.nodes.paragraph && likeness(o.textContent, now.textContent) >= 0.5);
      if (at < 0) {
        added(now);
        continue;
      }
      olds.slice(k, at).forEach(removed);
      changed(olds[at], now);
      k = at + 1;
    }
    olds.slice(k).forEach(removed);
    if (pa < a.length && pb < b.length) out.push(b[pb]);
    ai = pa + 1;
    bi = pb + 1;
  }
  return schema.nodes.doc.create(null, out.length ? out : schema.nodes.paragraph.create());
}

/** Neighbouring text with the same formatting as one piece. */
function joinText(nodes: Node[]): Node[] {
  const out: Node[] = [];
  for (const n of nodes) {
    const last = out[out.length - 1];
    if (last?.isText && n.isText && last.sameMarkup(n)) out[out.length - 1] = schema.text(last.text! + n.text!, n.marks);
    else out.push(n);
  }
  return out;
}

/** How many changes a comparison found (insertions and deletions, by reader's units). */
export function countChanges(doc: Node): number {
  let count = 0;
  let last: string | null = null;
  doc.descendants((n) => {
    if (n.isTextblock) {
      last = null;
      return true;
    }
    if (!n.isInline) return true;
    const kind = schema.marks.insertion.isInSet(n.marks) ? "ins" : schema.marks.deletion.isInSet(n.marks) ? "del" : null;
    if (kind && kind !== last) count++;
    last = kind;
    return false;
  });
  return count;
}
