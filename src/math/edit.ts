// Writing an equation straight into the text, as in Word: the shorthand
// builds up into fractions, scripts and roots when you press space (or an
// operator), and each part becomes a box the cursor can go into.
//
//   I_inn␣   → I with "inn" below          x^2+   → x² +
//   a/b␣     → a over b                    (a+b)/c␣ → a+b over c
//   x_␣      → x with an empty box below, the cursor in it
//   \alpha␣  → α     \frac␣ → a fraction with two boxes     sin␣ → sin, upright
//
// The equation is kept as LaTeX in the document (ast.ts), so files, Word and
// everything else see the same thing as before.

import { ACCENTS, FUNCTIONS, LIMITS, NARY, SYMBOLS, parseMath, toLatex, type MathNode } from "./ast";

export type Kind = "frac" | "sqrt" | "root" | "sub" | "sup" | "subsup" | "delim" | "accent" | "matrix";
export type Box = { k: "box"; kind: Kind; slots: Row[]; open?: string; close?: string; chr?: string; cols?: number };
export type Item =
  /** One typed character, still plain: a letter, digit or operator. */
  | { k: "ch"; v: string }
  /** A word drawn upright (sin, lim) or ordinary text ("…"), stepped over as one. */
  | { k: "sym"; v: string; text?: boolean }
  | Box;
export type Row = Item[];

/** Where the cursor is: [item, slot, item, slot, …] down from the top, then the place in that row. */
export interface Caret {
  path: number[];
  index: number;
}

const ch = (v: string): Item => ({ k: "ch", v });
/** A space typed in an equation: a medium space, \; in the LaTeX. */
export const SPACE = " ";
const box = (kind: Kind, slots: Row[], extra: Partial<Box> = {}): Box => ({ k: "box", kind, slots, ...extra });
const isCh = (item: Item | undefined, test: string | RegExp): boolean =>
  !!item && item.k === "ch" && (typeof test === "string" ? test.includes(item.v) : test.test(item.v));
const ALNUM = /^[\p{L}\p{N}.]$/u;
const LETTER = /^[a-zA-Z]$/;
/** Signs that belong to what follows ^ or _ (x^-1, Na^+), not operators of their own. */
const SIGNS = "+-±∓";
const MATRICES: Record<string, [string, string]> = { matrix: ["", ""], pmatrix: ["(", ")"], bmatrix: ["[", "]"], vmatrix: ["|", "|"], Bmatrix: ["{", "}"] };
const CLOSING: Record<string, string> = { ")": "(", "]": "[" };
export const BIG = new Set(Object.values(NARY));

export function rowAt(root: Row, path: number[]): Row {
  let r = root;
  for (let i = 0; i < path.length; i += 2) r = (r[path[i]] as Box).slots[path[i + 1]];
  return r;
}

const slotCount = (item: Item) => (item.k === "box" ? item.slots.length : 0);

export class MathEdit {
  caret: Caret;
  /** The box just before the cursor is marked; the next Backspace removes it. */
  marked = false;

  constructor(public root: Row = []) {
    this.caret = { path: [], index: root.length };
  }

  get row(): Row {
    return rowAt(this.root, this.caret.path);
  }

  isEmpty() {
    return this.root.length === 0;
  }

  // ---- typing --------------------------------------------------------------------

  type(text: string) {
    for (const c of text.replace(/\r?\n/g, " ")) this.typeOne(c);
  }

  private typeOne(c: string) {
    if (this.marked) this.backspace();
    const row = this.row;
    const before = row[this.caret.index - 1];
    if (this.inQuote()) {
      if (c === '"') this.closeQuote();
      else this.insert(ch(c));
      return;
    }
    // Straight after ^ _ / everything is part of what comes next: x^-1, x^(a+b).
    const afterOperator = isCh(before, "_^/");
    if (c === " ") {
      // Space builds up what was typed; with nothing to build it is a space, which also keeps things apart (a b/c: only b/c).
      if (!this.build(true)) this.insert(ch(SPACE));
      return;
    }
    if (c === "'") c = "′";
    if (!afterOperator) {
      if (c === "\\" || c === "_" || c === "^" || c === "/") this.convertCommand(false);
      else if (!LETTER.test(c) && !ALNUM.test(c) && c !== "′") this.build(false);
      else if (!LETTER.test(c)) this.convertCommand(false);
    }
    this.insert(ch(c));
    if (CLOSING[c]) this.matchBracket(c);
  }

  private insert(item: Item) {
    this.row.splice(this.caret.index, 0, item);
    this.caret.index++;
  }

  /** Inserts a whole equation (pasted LaTeX) at the cursor. */
  insertRow(items: Row) {
    this.row.splice(this.caret.index, 0, ...items);
    this.caret.index += items.length;
  }

  private inQuote(): boolean {
    return this.row.slice(0, this.caret.index).filter((i) => isCh(i, '"')).length % 2 === 1;
  }

  private closeQuote() {
    const row = this.row;
    let j = this.caret.index - 1;
    while (!isCh(row[j], '"')) j--;
    const text = row.slice(j + 1, this.caret.index).map((i) => (i.k === "sym" ? i.v : i.k === "ch" ? i.v : "")).join("");
    row.splice(j, this.caret.index - j, { k: "sym", v: text, text: true });
    this.caret.index = j + 1;
  }

  /** ) or ] closes the nearest open bracket before it into brackets that grow with what's inside. */
  private matchBracket(close: string) {
    const row = this.row;
    const end = this.caret.index - 1;
    let j = end - 1;
    while (j >= 0 && !isCh(row[j], CLOSING[close])) j--;
    if (j < 0) return;
    const body = row.slice(j + 1, end);
    row.splice(j, end - j + 1, box("delim", [body], { open: CLOSING[close], close }));
    this.caret.index = j + 1;
  }

  /**
   * Builds up what was typed just before the cursor. `bySpace` is true for
   * the space bar: it may make empty boxes (\frac, x_) and put the cursor in one.
   */
  build(bySpace: boolean): boolean {
    if (this.convertCommand(bySpace)) return true;
    if (this.convertWord(bySpace)) return true;
    return this.buildTail();
  }

  /** \name just before the cursor. */
  private convertCommand(bySpace: boolean): boolean {
    const row = this.row;
    let j = this.caret.index;
    while (j > 0 && isCh(row[j - 1], LETTER)) j--;
    if (j === 0 || !isCh(row[j - 1], "\\") || j === this.caret.index) return false;
    const name = row.slice(j, this.caret.index).map((i) => (i as { v: string }).v).join("");
    const start = j - 1;
    const replace = (item: Item, into?: number[]) => {
      row.splice(start, this.caret.index - start, item);
      this.caret.index = start + 1;
      if (into) this.caret = { path: [...this.caret.path, start, ...into.slice(0, 1)], index: 0 };
      return true;
    };
    if (name in SYMBOLS && name.length > 1) return replace(ch(SYMBOLS[name] || " "));
    if (FUNCTIONS.has(name)) return replace({ k: "sym", v: name });
    if (NARY[name]) return bySpace ? replace(box("subsup", [[ch(NARY[name])], [], []]), [1]) : replace(ch(NARY[name]));
    if (LIMITS.has(name)) return bySpace ? replace(box("sub", [[{ k: "sym", v: name }], []]), [1]) : replace({ k: "sym", v: name });
    if (!bySpace) return false;
    if (name === "frac") return replace(box("frac", [[], []]), [0]);
    if (name === "sqrt") return replace(box("sqrt", [[]]), [0]);
    if (name === "root") return replace(box("root", [[], []]), [0]);
    if (ACCENTS[name]) return replace(box("accent", [[]], { chr: ACCENTS[name] }), [0]);
    if (MATRICES[name]) {
      const [open, close] = MATRICES[name];
      return replace(box("matrix", [[], [], [], []], { open, close, cols: 2 }), [0]);
    }
    if (name === "text") {
      row.splice(start, this.caret.index - start, ch('"'));
      this.caret.index = start + 1;
      return true;
    }
    return false;
  }

  /** sin, cos, lim … typed without a backslash, as Word knows them. */
  private convertWord(bySpace: boolean): boolean {
    const row = this.row;
    let j = this.caret.index;
    while (j > 0 && isCh(row[j - 1], LETTER)) j--;
    if (j === this.caret.index || isCh(row[j - 1], "\\_^")) return false;
    const name = row.slice(j, this.caret.index).map((i) => (i as { v: string }).v).join("");
    if (!FUNCTIONS.has(name) && !LIMITS.has(name)) return false;
    const fn: Item = { k: "sym", v: name };
    if (LIMITS.has(name) && bySpace) {
      row.splice(j, name.length, box("sub", [[fn], []]));
      this.caret = { path: [...this.caret.path, j, 1], index: 0 };
    } else {
      row.splice(j, name.length, fn);
      this.caret.index = j + 1;
    }
    return true;
  }

  /** x_i, x^2, a/b … from the last operator up to the cursor. */
  private buildTail(): boolean {
    const row = this.row;
    const end = this.caret.index;
    let start = end;
    while (start > 0) {
      const item = row[start - 1];
      const boundary =
        isOperator(item) || (item.k === "ch" && !ALNUM.test(item.v) && !"_^/′".includes(item.v) && !(SIGNS.includes(item.v) && isCh(row[start - 2], "_^/")));
      if (boundary) break;
      start--;
    }
    const segment = row.slice(start, end);
    if (!segment.some((i) => isCh(i, "_^/"))) return false;
    const built = buildSegment(segment);
    row.splice(start, end - start, ...built);
    const empty = firstEmpty(built);
    if (empty) this.caret = { path: [...this.caret.path, start + empty[0], ...empty.slice(1)], index: 0 };
    else this.caret.index = start + built.length;
    return true;
  }

  /** Builds up everything still typed out, when leaving the equation. */
  finish() {
    const visit = (path: number[]) => {
      const r = rowAt(this.root, path);
      this.caret = { path, index: r.length };
      this.build(true);
      r.forEach((item, i) => {
        if (item.k === "box") item.slots.forEach((_, s) => visit([...path, i, s]));
      });
    };
    visit([]);
    this.caret = { path: [], index: this.root.length };
    this.marked = false;
  }

  // ---- moving ----------------------------------------------------------------------

  /** One place left; false at the start of the equation (the cursor leaves it). */
  left(): boolean {
    this.marked = false;
    const { path, index } = this.caret;
    const row = this.row;
    if (index > 0) {
      const item = row[index - 1];
      if (item.k === "box") {
        const s = item.slots.length - 1;
        this.caret = { path: [...path, index - 1, s], index: item.slots[s].length };
      } else this.caret.index--;
      return true;
    }
    if (!path.length) return false;
    const [p, s] = path.slice(-2);
    const up = path.slice(0, -2);
    if (s > 0) this.caret = { path: [...up, p, s - 1], index: rowAt(this.root, [...up, p, s - 1]).length };
    else this.caret = { path: up, index: p };
    return true;
  }

  right(): boolean {
    this.marked = false;
    const { path, index } = this.caret;
    const row = this.row;
    if (index < row.length) {
      const item = row[index];
      if (item.k === "box") this.caret = { path: [...path, index, 0], index: 0 };
      else this.caret.index++;
      return true;
    }
    if (!path.length) return false;
    const [p, s] = path.slice(-2);
    const up = path.slice(0, -2);
    if (s + 1 < slotCount(rowAt(this.root, up)[p])) this.caret = { path: [...up, p, s + 1], index: 0 };
    else this.caret = { path: up, index: p + 1 };
    return true;
  }

  /** Up or down between a fraction's top and bottom, a script and its base, matrix rows. */
  vertical(dir: -1 | 1): boolean {
    this.marked = false;
    for (let d = this.caret.path.length - 2; d >= 0; d -= 2) {
      const up = this.caret.path.slice(0, d);
      const [p, s] = this.caret.path.slice(d, d + 2);
      const b = rowAt(this.root, up)[p] as Box;
      let to = -1;
      if (b.kind === "frac") to = dir < 0 ? (s === 1 ? 0 : -1) : s === 0 ? 1 : -1;
      else if (b.kind === "subsup") to = dir < 0 ? (s !== 2 ? 2 : -1) : s !== 1 ? 1 : -1;
      else if (b.kind === "sub") to = dir < 0 && s === 1 ? 0 : dir > 0 && s === 0 ? 1 : -1;
      else if (b.kind === "sup") to = dir < 0 && s === 0 ? 1 : dir > 0 && s === 1 ? 0 : -1;
      else if (b.kind === "matrix") {
        const n = s + dir * (b.cols ?? 1);
        to = n >= 0 && n < b.slots.length ? n : -1;
      }
      if (to >= 0) {
        const target = b.slots[to];
        this.caret = { path: [...up, p, to], index: Math.min(this.caret.index, target.length) };
        return true;
      }
    }
    return false;
  }

  home() {
    this.marked = false;
    if (this.caret.index > 0) this.caret.index = 0;
    else this.caret = { path: [], index: 0 };
  }

  end() {
    this.marked = false;
    if (this.caret.index < this.row.length) this.caret.index = this.row.length;
    else this.caret = { path: [], index: this.root.length };
  }

  /** Tab in a matrix goes to the next cell; elsewhere false. */
  nextCell(): boolean {
    const m = this.enclosingMatrix();
    if (!m) return false;
    const { up, p, s, b } = m;
    if (s + 1 < b.slots.length) this.caret = { path: [...up, p, s + 1], index: 0 };
    else this.caret = { path: up, index: p + 1 };
    return true;
  }

  /** Enter in a matrix adds a row under the cursor's; elsewhere false. */
  addMatrixRow(): boolean {
    const m = this.enclosingMatrix();
    if (!m) return false;
    const { up, p, s, b } = m;
    const cols = b.cols ?? 1;
    const at = (Math.floor(s / cols) + 1) * cols;
    b.slots.splice(at, 0, ...Array.from({ length: cols }, () => []));
    this.caret = { path: [...up, p, at], index: 0 };
    return true;
  }

  private enclosingMatrix() {
    for (let d = this.caret.path.length - 2; d >= 0; d -= 2) {
      const up = this.caret.path.slice(0, d);
      const [p, s] = this.caret.path.slice(d, d + 2);
      const b = rowAt(this.root, up)[p] as Box;
      if (b.kind === "matrix") return { up, p, s, b };
    }
    return null;
  }

  // ---- deleting --------------------------------------------------------------------

  /** As in Word: a box before the cursor is first marked, then removed; at the start of a box it comes apart. */
  backspace() {
    const { path, index } = this.caret;
    const row = this.row;
    if (this.marked) {
      row.splice(index - 1, 1);
      this.caret.index--;
      this.marked = false;
      return;
    }
    if (index > 0) {
      if (row[index - 1].k === "box") this.marked = true;
      else {
        row.splice(index - 1, 1);
        this.caret.index--;
      }
      return;
    }
    if (!path.length) return;
    const [p, s] = path.slice(-2);
    const up = path.slice(0, -2);
    const parent = rowAt(this.root, up);
    const b = parent[p] as Box;
    const before = b.slots.slice(0, s).reduce((n, slot) => n + slot.length, 0);
    parent.splice(p, 1, ...b.slots.flat());
    this.caret = { path: up, index: p + before };
  }

  del() {
    this.marked = false;
    const row = this.row;
    const item = row[this.caret.index];
    if (!item) return;
    if (item.k === "box") this.right();
    else row.splice(this.caret.index, 1);
  }

  // ---- saving ------------------------------------------------------------------------

  latex(): string {
    return toLatex(toNode(this.root));
  }
}

/** sin, lim, ∑ (with or without limits): they work on what follows, so a fraction after them starts after them. */
function isOperator(item: Item): boolean {
  if (item.k === "sym") return !item.text;
  if (item.k === "ch") return BIG.has(item.v);
  if (item.kind !== "sub" && item.kind !== "sup" && item.kind !== "subsup") return false;
  const base = item.slots[0];
  return base.length === 1 && isOperator(base[0]);
}

/** Builds x_i, x^2 and a/b in a stretch of typed text with no other operators in it. */
function buildSegment(segment: Item[]): Row {
  // Fractions bind loosest: a_1/b^2 is a₁ over b².
  const parts: Item[][] = [[]];
  for (const item of segment) {
    if (isCh(item, "/")) parts.push([]);
    else parts[parts.length - 1].push(item);
  }
  let out = scripts(parts[0]);
  for (const part of parts.slice(1)) out = [box("frac", [unwrap(out), unwrap(scripts(part))])];
  return out;
}

/** (a+b) as a whole operand loses its brackets, as in Word: (a+b)/c, x^(n+1). */
function unwrap(r: Row): Row {
  const only = r.length === 1 ? r[0] : null;
  return only && only.k === "box" && only.kind === "delim" && only.open === "(" ? only.slots[0] : r;
}

function scripts(part: Item[]): Row {
  const out: Row = [];
  for (let i = 0; i < part.length; i++) {
    const item = part[i];
    if (!isCh(item, "_^")) {
      out.push(item);
      continue;
    }
    const isSub = (item as { v: string }).v === "_";
    // What follows: an optional sign, then a run of letters and digits (all of "inn"), or one thing.
    const operand: Row = [];
    let j = i + 1;
    if (isCh(part[j], SIGNS)) operand.push(part[j++]);
    while (j < part.length && isCh(part[j], ALNUM)) operand.push(part[j++]);
    if (operand.length === 0 || (operand.length === 1 && isCh(operand[0], SIGNS))) {
      if (j < part.length && !isCh(part[j], "_^")) operand.push(part[j++]);
    }
    // e^-x^2: a second ^ goes on top of the first (and _ under _), as in Word.
    if (isCh(part[j], isSub ? "_" : "^")) {
      operand.splice(0, operand.length, ...scripts(part.slice(i + 1)));
      j = part.length;
    }
    i = j - 1;
    const slot = unwrap(operand);
    // x_i^2: the second script joins the first.
    const last = out[out.length - 1];
    if (last && last.k === "box" && ((last.kind === "sub" && !isSub) || (last.kind === "sup" && isSub))) {
      const [base, first] = last.slots;
      out[out.length - 1] = box("subsup", isSub ? [base, slot, first] : [base, first, slot]);
      continue;
    }
    if (last && last.k === "box" && last.kind === "subsup") {
      last.slots[isSub ? 1 : 2].push(...slot);
      continue;
    }
    // The base: a number (10^5), otherwise the one thing just before.
    let base: Row = [];
    if (isCh(last, /^[0-9.]$/)) {
      let k = out.length;
      while (k > 0 && isCh(out[k - 1], /^[0-9.]$/)) k--;
      base = out.splice(k);
    } else if (last) base = [out.pop()!];
    out.push(box(isSub ? "sub" : "sup", [base, slot]));
  }
  return out;
}

/** The first empty box (depth first), as a path relative to `r`. */
function firstEmpty(r: Row): number[] | null {
  for (let i = 0; i < r.length; i++) {
    const item = r[i];
    if (item.k !== "box") continue;
    for (let s = 0; s < item.slots.length; s++) {
      if (item.slots[s].length === 0) return [i, s];
      const inner = firstEmpty(item.slots[s]);
      if (inner) return [i, s, ...inner];
    }
  }
  return null;
}

// ---- to and from the document's LaTeX ---------------------------------------------------

export function fromLatex(latex: string): Row {
  if (!latex.trim()) return [];
  try {
    return fromNode(parseMath(latex));
  } catch {
    return [...latex].map(ch);
  }
}

export function fromNode(n: MathNode): Row {
  switch (n.t) {
    case "row":
      return n.items.flatMap(fromNode);
    case "id":
    case "num":
    case "op":
      return [...n.v].map((c) => ch(c === "−" ? "-" : c === "∗" ? "*" : c));
    case "fn":
      return [{ k: "sym", v: n.v }];
    case "text":
      return [{ k: "sym", v: n.v, text: true }];
    case "frac":
      return [box("frac", [fromNode(n.num), fromNode(n.den)])];
    case "sqrt":
      return [n.index ? box("root", [fromNode(n.index), fromNode(n.body)]) : box("sqrt", [fromNode(n.body)])];
    case "scripts":
      return [scriptBox(fromNode(n.base), n.sub && fromNode(n.sub), n.sup && fromNode(n.sup))];
    case "nary": {
      const op = ch(n.op);
      const head = n.sub || n.sup ? scriptBox([op], n.sub && fromNode(n.sub), n.sup && fromNode(n.sup)) : op;
      return [head, ...fromNode(n.body)];
    }
    case "limit":
      return n.under.t === "row" && n.under.items.length === 0 ? [{ k: "sym", v: n.name }] : [box("sub", [[{ k: "sym", v: n.name }], fromNode(n.under)])];
    case "delim":
      return [box("delim", [fromNode(n.body)], { open: n.open, close: n.close })];
    case "accent":
      return [box("accent", [fromNode(n.body)], { chr: n.chr })];
    case "matrix": {
      const cols = Math.max(1, ...n.rows.map((r) => r.length));
      return [box("matrix", n.rows.flatMap((r) => Array.from({ length: cols }, (_, i) => (r[i] ? fromNode(r[i]) : []))), { open: n.open, close: n.close, cols })];
    }
  }
}

function scriptBox(base: Row, sub?: Row, sup?: Row): Box {
  if (sub && sup) return box("subsup", [base, sub, sup]);
  return sub ? box("sub", [base, sub]) : box("sup", [base, sup ?? []]);
}

export function toNode(r: Row): MathNode {
  const items: MathNode[] = [];
  for (let i = 0; i < r.length; i++) {
    const item = r[i];
    if (item.k === "ch") {
      if (/^[0-9]$/.test(item.v)) {
        let v = item.v;
        while (isCh(r[i + 1], /^[0-9]$/) || (isCh(r[i + 1], ".") && isCh(r[i + 2], /^[0-9]$/))) v += (r[++i] as { v: string }).v;
        items.push({ t: "num", v });
      } else if (/^\p{L}$/u.test(item.v)) items.push({ t: "id", v: item.v });
      else items.push({ t: "op", v: item.v === "-" ? "−" : item.v === "*" ? "∗" : item.v });
    } else if (item.k === "sym") items.push(item.text ? { t: "text", v: item.v } : { t: "fn", v: item.v });
    else items.push(boxNode(item));
  }
  return items.length === 1 ? items[0] : { t: "row", items };
}

function boxNode(b: Box): MathNode {
  const [a, c, d] = b.slots.map(toNode);
  switch (b.kind) {
    case "frac":
      return { t: "frac", num: a, den: c };
    case "sqrt":
      return { t: "sqrt", body: a };
    case "root":
      return { t: "sqrt", index: a, body: c };
    case "sub":
      return { t: "scripts", base: a, sub: c };
    case "sup":
      return { t: "scripts", base: a, sup: c };
    case "subsup":
      return { t: "scripts", base: a, sub: c, sup: d };
    case "delim":
      return { t: "delim", open: b.open ?? "(", close: b.close ?? ")", body: a };
    case "accent":
      return { t: "accent", chr: b.chr ?? "̂", body: a };
    case "matrix": {
      const cols = b.cols ?? 1;
      const rows: MathNode[][] = [];
      b.slots.forEach((slot, i) => {
        if (i % cols === 0) rows.push([]);
        rows[rows.length - 1].push(toNode(slot));
      });
      return { t: "matrix", rows, open: b.open ?? "", close: b.close ?? "" };
    }
  }
}
