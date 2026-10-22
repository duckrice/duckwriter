// Formatting commands. Formatting comes from three places, as in Word: the
// document defaults, the paragraph's style, and direct formatting (marks).
// Commands look at the result of all three, so "Bold" in a bold heading turns
// bold off rather than doing nothing.

import type { Mark, MarkType, Node } from "prosemirror-model";
import { TextSelection, type Command, type EditorState, type Transaction } from "prosemirror-state";
import { resolvedPara, resolvedRun, styleChain, type DocContext, type ParaProps, type RunProps } from "../docx/model";
import { listOf, listPluginKey, selectedParagraphs } from "./lists";
import { schema } from "./schema";
import { markAttr, paraMarks, updateParaMarks } from "./paramark";

/** Formatting chosen with the cursor in an empty line is kept on that line's paragraph mark, as in Word. */
function rememberInEmptyLine(tr: Transaction, state: EditorState, marks: readonly Mark[]): Transaction {
  const { $from } = state.selection;
  if ($from.parent.type === schema.nodes.paragraph && !$from.parent.content.size) tr.setNodeMarkup($from.before(), undefined, { ...$from.parent.attrs, mark: markAttr(marks) });
  return tr;
}

/** The character formatting a piece of text ends up with. */
export function effectiveRun(ctx: DocContext, para: Node | null, marks: readonly Mark[]): RunProps {
  const run: RunProps = { ...resolvedRun(ctx.styles, para?.attrs.style) };
  const charStyle = marks.find((m) => m.type === schema.marks.charStyle);
  if (charStyle) for (const s of styleChain(ctx.styles, charStyle.attrs.id)) Object.assign(run, s.run);
  for (const mark of marks) {
    const a = mark.attrs;
    switch (mark.type.name) {
      case "bold":
        run.bold = a.on;
        break;
      case "italic":
        run.italic = a.on;
        break;
      case "underline":
        run.underline = a.on;
        break;
      case "strike":
        run.strike = a.on;
        break;
      case "font":
        run.font = a.family;
        break;
      case "size":
        run.size = a.pt;
        break;
      case "color":
        run.color = a.color;
        break;
      case "highlight":
        run.highlight = a.color;
        break;
      case "vertAlign":
        run.vertAlign = a.value;
        break;
    }
  }
  return run;
}

/** The paragraph formatting a paragraph ends up with. */
export function effectivePara(ctx: DocContext, para: Node): ParaProps {
  const props = resolvedPara(ctx.styles, para.attrs.style);
  for (const key of ["align", "spaceBefore", "spaceAfter", "lineSpacing", "lineHeight", "indentLeft", "indentRight", "indentFirst"] as const) {
    if (para.attrs[key] !== null) (props as Record<string, unknown>)[key] = para.attrs[key];
  }
  return props;
}

/** The formatting at the cursor, or at the start of the selection. */
export function formattingAt(ctx: DocContext, state: EditorState): { run: RunProps; para: ParaProps; paragraph: Node | null } {
  const { $from, empty, from, to } = state.selection;
  let paragraph = $from.parent.type === schema.nodes.paragraph ? $from.parent : null;
  let marks: readonly Mark[] = state.storedMarks ?? $from.marks();
  if (!empty) {
    // The first letter of the selection, not whatever precedes it (and with
    // everything selected, the first letter of the document).
    let found = false;
    state.doc.nodesBetween(from, to, (node, _pos, parent) => {
      if (found) return false;
      if (node.isInline && parent) {
        marks = node.marks;
        paragraph = parent.type === schema.nodes.paragraph ? parent : paragraph;
        found = true;
        return false;
      }
      return true;
    });
  }
  return { run: effectiveRun(ctx, paragraph, marks), para: paragraph ? effectivePara(ctx, paragraph) : {}, paragraph };
}

/** Text in the selection, each piece with its paragraph. */
function textRanges(state: EditorState): { node: Node; para: Node; from: number; to: number }[] {
  const out: { node: Node; para: Node; from: number; to: number }[] = [];
  const { from, to } = state.selection;
  state.doc.nodesBetween(from, to, (node, pos, parent) => {
    if (node.isInline && parent) out.push({ node, para: parent, from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) });
  });
  return out;
}

type Toggle = "bold" | "italic" | "underline" | "strike";

/** Bold, italic, underline or strike-through on or off, Word style. */
export function toggleFormat(ctx: DocContext, name: Toggle): Command {
  return (state, dispatch) => {
    const type = schema.marks[name];
    const { empty, $from } = state.selection;
    if (empty) {
      const current = formattingAt(ctx, state).run[name] ?? false;
      if (dispatch) {
        const paraWants = resolvedRun(ctx.styles, $from.parent.attrs.style)[name] ?? false;
        let marks = type.removeFromSet(state.storedMarks ?? $from.marks());
        if (!current !== paraWants) marks = type.create({ on: !current }).addToSet(marks);
        dispatch(rememberInEmptyLine(state.tr.setStoredMarks(marks), state, marks));
      }
      return true;
    }
    const ranges = textRanges(state).filter((r) => r.node.isText);
    if (!ranges.length) return false;
    // On, unless everything selected already has it.
    const turnOn = !ranges.every((r) => effectiveRun(ctx, r.para, r.node.marks)[name]);
    if (dispatch) {
      const tr = state.tr;
      for (const r of ranges) {
        tr.removeMark(r.from, r.to, type);
        const fromStyle = effectiveRun(ctx, r.para, r.node.marks.filter((m) => m.type !== type))[name] ?? false;
        if (fromStyle !== turnOn) tr.addMark(r.from, r.to, type.create({ on: turnOn }));
      }
      updateParaMarks(tr, state.selection.from, state.selection.to, (marks, para) => {
        const rest = type.removeFromSet(marks);
        return (resolvedRun(ctx.styles, para.attrs.style)[name] ?? false) === turnOn ? rest : type.create({ on: turnOn }).addToSet(rest);
      });
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/** Sets (or with null removes) a mark with attributes, like a font or a colour, on the selection. */
export function setMark(type: MarkType, attrs: Record<string, unknown> | null): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const { empty, $from, from, to } = state.selection;
    if (empty) {
      let marks = type.removeFromSet(state.storedMarks ?? $from.marks());
      if (attrs) marks = type.create(attrs).addToSet(marks);
      dispatch(rememberInEmptyLine(state.tr.setStoredMarks(marks), state, marks));
      return true;
    }
    const tr = state.tr.removeMark(from, to, type);
    if (attrs) tr.addMark(from, to, type.create(attrs));
    // Paragraphs selected to their end take it on their paragraph mark too, so their empty lines and new lines keep it.
    updateParaMarks(tr, from, to, (marks) => (attrs ? type.create(attrs).addToSet(type.removeFromSet(marks)) : type.removeFromSet(marks)));
    dispatch(tr.scrollIntoView());
    return true;
  };
}

/** Superscript or subscript on or off. */
export function toggleVertAlign(value: "superscript" | "subscript"): Command {
  return (state, dispatch) => {
    const { $from } = state.selection;
    const marks = state.selection.empty ? (state.storedMarks ?? $from.marks()) : ($from.nodeAfter?.marks ?? []);
    const current = marks.find((m) => m.type === schema.marks.vertAlign)?.attrs.value;
    return setMark(schema.marks.vertAlign, current === value ? null : { value })(state, dispatch);
  };
}

/** Makes text bigger or smaller by a step on Word's size ladder. */
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
export function stepSize(ctx: DocContext, direction: 1 | -1): Command {
  return (state, dispatch) => {
    const current = formattingAt(ctx, state).run.size ?? 11;
    const next = direction > 0 ? (SIZES.find((s) => s > current) ?? current + 10) : ([...SIZES].reverse().find((s) => s < current) ?? Math.max(1, current - 1));
    return setMark(schema.marks.size, { pt: next })(state, dispatch);
  };
}

/** Makes text a number of points bigger (or smaller), as Ctrl+] and Ctrl+[ do in Word. */
export function stepSizeBy(ctx: DocContext, points: number): Command {
  return (state, dispatch) => {
    const current = formattingAt(ctx, state).run.size ?? 11;
    return setMark(schema.marks.size, { pt: Math.min(1638, Math.max(1, Math.round((current + points) * 2) / 2)) })(state, dispatch);
  };
}

/** Changes attributes of every selected paragraph. */
export function setParagraphs(change: (attrs: Node["attrs"], node: Node) => Partial<Node["attrs"]>): Command {
  return (state, dispatch) => {
    const paras = selectedParagraphs(state);
    if (!paras.length) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const { node, pos } of paras) tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...change(node.attrs, node) });
      dispatch(tr.setMeta(listPluginKey, true).scrollIntoView());
    }
    return true;
  };
}

/** Applies a paragraph style; direct paragraph formatting is cleared, as in Word. */
export function setStyle(ctx: DocContext, id: string | null): Command {
  const style = id === ctx.styles.defaultParagraph ? null : id;
  return setParagraphs((attrs) => ({
    style,
    align: null,
    spaceBefore: null,
    spaceAfter: null,
    lineSpacing: null,
    lineHeight: null,
    lineRule: null,
    indentLeft: attrs.numId !== null ? attrs.indentLeft : null,
    indentFirst: attrs.numId !== null ? attrs.indentFirst : null,
  }));
}

export function setAlign(align: "left" | "center" | "right" | "justify"): Command {
  return setParagraphs(() => ({ align }));
}

/**
 * Sorts the paragraphs the selection touches by their text, A to Z (or Z to A), as Word's Sort:
 * numbers in order of size, letters as the language orders them (æ, ø and å last in Norwegian).
 */
export function sortParagraphs(descending = false): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    // The paragraphs' parent: the document (also for Select all), or a table cell.
    const depth = Math.max(0, Math.min($from.depth, $to.depth) - 1);
    if ($from.node(depth) !== $to.node(depth)) return false;
    const parent = $from.node(depth);
    const first = $from.index(depth);
    const last = Math.min($to.index(depth), parent.childCount - 1);
    if (last <= first) return false;
    const blocks: import("prosemirror-model").Node[] = [];
    let start = depth === 0 ? 0 : $from.start(depth);
    for (let i = 0; i < first; i++) start += parent.child(i).nodeSize;
    let end = start;
    for (let i = first; i <= last; i++) {
      blocks.push(parent.child(i));
      end += parent.child(i).nodeSize;
    }
    if (blocks.some((b) => !b.isTextblock)) return false;
    const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });
    const sorted = [...blocks].sort((a, b) => collator.compare(a.textContent, b.textContent) * (descending ? -1 : 1));
    if (dispatch) dispatch(state.tr.replaceWith(start, end, sorted).scrollIntoView());
    return true;
  };
}

export function setLineSpacing(multiple: number): Command {
  return setParagraphs(() => ({ lineSpacing: multiple, lineHeight: null, lineRule: null }));
}

/** Indents paragraphs by half an inch more (+1) or less (-1). */
export function stepIndent(ctx: DocContext, direction: 1 | -1): Command {
  return setParagraphs((_attrs, node) => {
    const current = effectivePara(ctx, node).indentLeft ?? 0;
    return { indentLeft: Math.max(0, Math.round((current + 36 * direction) / 36) * 36) };
  });
}

/** Removes character formatting from the selection and resets its paragraphs to the default style. */
export function clearFormatting(ctx: DocContext): Command {
  return (state, dispatch) => {
    if (!dispatch) return true;
    const { from, to, empty } = state.selection;
    const tr = state.tr;
    if (!empty) {
      for (const type of Object.values(schema.marks)) if (type !== schema.marks.link) tr.removeMark(from, to, type);
      updateParaMarks(tr, from, to, () => []);
    } else rememberInEmptyLine(tr.setStoredMarks([]), state, []);
    const reset = setStyle(ctx, null);
    reset(state.apply(tr), (next) => {
      for (const step of next.steps) tr.step(step);
    });
    dispatch(tr);
    return true;
  };
}

/**
 * Enter. A new paragraph after a heading gets the style that follows it
 * (normally Normal); otherwise it keeps the paragraph's formatting, and the
 * text formatting carries on. Enter in an empty list item ends the list.
 */
/** Formatting without a link, or Word's blue underlined style for links. */
export function withoutLink(marks: readonly Mark[]): readonly Mark[] {
  return marks.filter((m) => m.type !== schema.marks.link && !(m.type === schema.marks.charStyle && m.attrs.id === "Hyperlink"));
}

export function newParagraph(ctx: DocContext): Command {
  return (state, dispatch) => {
    const { $from, $to } = state.selection;
    const para = $from.parent;
    if (para.type !== schema.nodes.paragraph || $from.parent !== $to.parent) return false;
    if (para.content.size === 0 && listOf(ctx, para)) {
      const fromStyle = resolvedPara(ctx.styles, para.attrs.style).numId;
      const style = para.attrs.style === "ListParagraph" ? null : para.attrs.style;
      if (dispatch) dispatch(state.tr.setNodeMarkup($from.before(), undefined, { ...para.attrs, style, numId: fromStyle ? "0" : null, ilvl: 0, indentLeft: null, indentFirst: null }).setMeta(listPluginKey, true));
      return true;
    }
    if (!dispatch) return true;
    // A new paragraph goes on in the same formatting, but not as a link.
    const marks = withoutLink(state.storedMarks ?? ($from.parentOffset > 0 ? $from.marks() : (para.firstChild?.marks ?? paraMarks(para.attrs))));
    const atEnd = $to.parentOffset === para.content.size;
    const style = para.attrs.style ?? ctx.styles.defaultParagraph;
    const next = style ? ctx.styles.byId.get(style)?.next : undefined;
    let attrs = para.attrs;
    // A list carries on to the next item whatever the style says comes next.
    if (atEnd && next && next !== style && !listOf(ctx, para)) {
      attrs = { ...schema.nodes.paragraph.create().attrs, style: next === ctx.styles.defaultParagraph ? null : next };
    }
    const carried = atEnd && attrs !== para.attrs ? [] : marks;
    const tr = state.tr.deleteSelection();
    tr.split(tr.mapping.map($from.pos), 1, [{ type: schema.nodes.paragraph, attrs: { ...attrs, mark: markAttr(carried) } }]);
    tr.setStoredMarks(carried);
    dispatch(tr.setMeta(listPluginKey, true).scrollIntoView());
    return true;
  };
}

/** Backspace at the start of a list item takes it out of the list first, as in Word. */
export function leaveListAtStart(ctx: DocContext): Command {
  return (state, dispatch) => {
    const { $from, empty } = state.selection;
    if (!empty || $from.parentOffset !== 0 || $from.parent.type !== schema.nodes.paragraph) return false;
    const para = $from.parent;
    if (!listOf(ctx, para)) return false;
    if (dispatch) {
      const fromStyle = resolvedPara(ctx.styles, para.attrs.style).numId;
      dispatch(state.tr.setNodeMarkup($from.before(), undefined, { ...para.attrs, numId: fromStyle ? "0" : null, ilvl: 0 }).setMeta(listPluginKey, true));
    }
    return true;
  };
}

export function insertPageBreak(): Command {
  return (state, dispatch) => {
    if (dispatch) {
      const tr = state.tr.replaceSelectionWith(schema.nodes.page_break.create(), true);
      // Continue on a new paragraph after the break, like Word.
      const pos = tr.selection.from;
      tr.split(pos);
      dispatch(tr.setSelection(TextSelection.create(tr.doc, pos + 2)).scrollIntoView());
    }
    return true;
  };
}

export function insertTab(): Command {
  return (state, dispatch) => {
    if (dispatch) dispatch(state.tr.insertText("\t").scrollIntoView());
    return true;
  };
}

export type { Transaction };

/** lower case → UPPER CASE → Title Case → lower case, for the selection or the word at the cursor (Alt+Shift+K). */
export function changeCase(): Command {
  return (state, dispatch) => {
    let { from, to } = state.selection;
    if (from === to) {
      const { $from } = state.selection;
      const text = $from.parent.textBetween(0, $from.parent.content.size, undefined, "\uFFFC");
      let s = $from.parentOffset;
      let e = s;
      while (s > 0 && /[\p{L}\p{N}'’-]/u.test(text[s - 1])) s--;
      while (e < text.length && /[\p{L}\p{N}'’-]/u.test(text[e])) e++;
      from = $from.start() + s;
      to = $from.start() + e;
    }
    // The selected text pieces, in order, so words running across formatting are one word.
    const pieces: { a: number; b: number; text: string; marks: readonly import("prosemirror-model").Mark[] }[] = [];
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (!node.isText) return true;
      const a = Math.max(from, pos);
      const b = Math.min(to, pos + node.nodeSize);
      pieces.push({ a, b, text: node.text!.slice(a - pos, b - pos), marks: node.marks });
      return false;
    });
    const text = pieces.map((p) => p.text).join("");
    if (!text.trim()) return false;
    const changed =
      text === text.toLowerCase() ? text.toUpperCase() : text === text.toUpperCase() ? text.toLowerCase().replace(/(^|[^\p{L}\p{N}'’])(\p{L})/gu, (_m, a: string, b: string) => a + b.toUpperCase()) : text.toLowerCase();
    // Letters that change length (ß → SS) can't be matched up piece by piece: then the whole is left alone.
    if (changed.length !== text.length) return false;
    if (dispatch) {
      const tr = state.tr;
      let at = 0;
      for (const p of pieces) {
        const next = changed.slice(at, at + p.text.length);
        at += p.text.length;
        if (next !== p.text) tr.replaceWith(p.a, p.b, state.schema.text(next, p.marks));
      }
      tr.setSelection(TextSelection.create(tr.doc, from, to));
      dispatch(tr.scrollIntoView());
    }
    return true;
  };
}

/**
 * Ctrl+Up and Ctrl+Down, as in Word: to the start of this paragraph (or the
 * one before, if already there), or to the start of the next; with Shift, the
 * selection grows to there.
 */
export function moveByParagraph(direction: -1 | 1, extend = false): Command {
  return (state, dispatch) => {
    const { head, anchor } = state.selection;
    const starts: number[] = [];
    state.doc.descendants((node, pos) => {
      if (!node.isTextblock) return true;
      starts.push(pos + 1);
      return false;
    });
    if (!starts.length) return false;
    let target: number | undefined;
    if (direction < 0) target = [...starts].reverse().find((s) => s < head);
    else target = starts.find((s) => s > head);
    // At the last paragraph, Ctrl+Down goes to its end.
    if (target === undefined) {
      if (direction < 0) return false;
      const last = starts[starts.length - 1];
      target = last + state.doc.resolve(last).parent.content.size;
      if (target === head) return false;
    }
    if (dispatch) {
      const selection = extend ? TextSelection.create(state.doc, anchor, target) : TextSelection.create(state.doc, target);
      dispatch(state.tr.setSelection(selection).scrollIntoView());
    }
    return true;
  };
}
