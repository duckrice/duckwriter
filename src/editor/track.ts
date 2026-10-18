// Track changes: while it is on, typed text is marked as inserted and deleted
// text stays, struck through, until the change is accepted, as in Word.
// Deleting your own tracked insertion simply removes it.

import { Fragment, Slice, type Mark, type Node } from "prosemirror-model";
import { Plugin, PluginKey, TextSelection, type Transaction } from "prosemirror-state";
import { AddMarkStep, Mapping, RemoveMarkStep, ReplaceAroundStep, ReplaceStep, type Step } from "prosemirror-transform";
import { paraPropsXml, runPropsXml } from "../docx/write";
import { newAnnotationId, type DocContext } from "../docx/model";
import { schema } from "./schema";

export const trackKey = new PluginKey("track");

/** Tells the tracking to leave a change alone (accepting changes, adding comments…). */
export const untracked = (tr: Transaction) => tr.setMeta("noTrack", true);

const nowIso = () => new Date().toISOString().replace(/\.\d+Z$/, "Z");

/** The marks that are formatting, whose changes are tracked as formatting changes. */
export const FORMATTING = new Set(["bold", "italic", "underline", "strike", "font", "size", "color", "highlight", "vertAlign", "charStyle"]);

export interface TrackOptions {
  on(): boolean;
  author(): string;
  context(): DocContext;
}

/** A mark for new text: the one next to it if that is the same person's, so typing makes one change. */
function changeMark(type: "insertion" | "deletion", neighbour: readonly Mark[] | undefined, opts: TrackOptions): Mark {
  const markType = schema.marks[type];
  const near = neighbour && markType.isInSet(neighbour);
  if (near && near.attrs.author === opts.author()) return near;
  return markType.create({ id: newAnnotationId(opts.context()), author: opts.author(), date: nowIso() });
}

const isOwnInsertion = (node: Node, author: string) => {
  const mark = schema.marks.insertion.isInSet(node.marks);
  return !!mark && mark.attrs.author === author && !schema.marks.deletion.isInSet(node.marks);
};

/** Whether a deleted piece is only text in paragraphs (not table rows or the like, which go untracked). */
function onlyText(fragment: Fragment): boolean {
  let ok = true;
  fragment.forEach((node) => {
    if (node.isInline) return;
    if (node.type !== schema.nodes.paragraph) ok = false;
  });
  return ok;
}

/** The deleted content again, marked as deleted; the user's own insertions left out (really deleted). */
function markDeleted(fragment: Fragment, mark: () => Mark, author: string): Fragment {
  const out: Node[] = [];
  fragment.forEach((node) => {
    if (node.isInline) {
      if (isOwnInsertion(node, author)) return;
      // A comment's anchors stay as they are.
      if (node.type === schema.nodes.comment) return void out.push(node);
      out.push(schema.marks.deletion.isInSet(node.marks) ? node : node.mark(mark().addToSet(node.marks)));
    } else out.push(node.copy(markDeleted(node.content, mark, author)));
  });
  return Fragment.from(out);
}

function hasText(fragment: Fragment): boolean {
  let found = false;
  fragment.descendants((node) => {
    if (node.isInline) found = true;
    return !found;
  });
  return found;
}

/** Word's mark on a table row inserted or deleted while changes were tracked. */
function rowMark(kind: "ins" | "del", opts: TrackOptions): string {
  const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
  return `<w:${kind} w:id="${newAnnotationId(opts.context())}" w:author="${esc(opts.author())}" w:date="${nowIso()}"/>`;
}

export function trackPlugin(opts: TrackOptions) {
  return new Plugin({
    key: trackKey,
    appendTransaction(trs, oldState, newState) {
      if (!opts.on()) return null;
      const user = trs.filter((tr) => tr.docChanged);
      if (!user.length || user.some((tr) => tr.getMeta("noTrack") || tr.getMeta("history$") || tr.getMeta("appendedTransaction") || tr.getMeta("addToHistory") === false)) return null;
      // Every step, with the document before it and the mapping from after it to now.
      const steps: { step: Step; before: Node; rest: Mapping }[] = [];
      for (const tr of user) {
        tr.steps.forEach((step, i) => steps.push({ step, before: tr.docs[i], rest: null as unknown as Mapping }));
      }
      const all = user.flatMap((tr) => tr.mapping.maps);
      steps.forEach((s, i) => (s.rest = new Mapping(all.slice(i + 1))));

      const out = newState.tr;
      const author = opts.author();
      let selection: number | null = null;
      for (const { step, before, rest } of steps) {
        // A paragraph's formatting (alignment, spacing, style…): what it was is kept, as Word's pPrChange.
        if (step instanceof ReplaceAroundStep) {
          const { from } = step as unknown as { from: number };
          const old = before.nodeAt(from);
          const at = out.mapping.map(rest.map(from, 1), 1);
          const now = out.doc.nodeAt(at);
          if (old?.type !== schema.nodes.paragraph || now?.type !== schema.nodes.paragraph || now.attrs.propsChange) continue;
          if (now.attrs.change?.kind === "ins" && now.attrs.change.author === author) continue;
          const fields = ["style", "align", "spaceBefore", "spaceAfter", "lineSpacing", "lineHeight", "lineRule", "indentLeft", "indentRight", "indentFirst", "numId", "ilvl"];
          if (fields.every((f) => old.attrs[f] === now.attrs[f])) continue;
          out.setNodeMarkup(at, null, { ...now.attrs, propsChange: { id: newAnnotationId(opts.context()), author, date: nowIso(), before: paraPropsXml(old.attrs, opts.context()) } });
          continue;
        }
        if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) {
          // Formatting: what it was before is kept, so the change can be rejected.
          const { from, to, mark } = step as unknown as { from: number; to: number; mark: Mark };
          if (!FORMATTING.has(mark.type.name)) continue;
          let change: Mark | null = null;
          before.nodesBetween(from, to, (node, pos) => {
            if (!node.isInline) return true;
            if (schema.marks.formatChange.isInSet(node.marks) || isOwnInsertion(node, author)) return false;
            const a = out.mapping.map(rest.map(Math.max(pos, from)));
            const b = out.mapping.map(rest.map(Math.min(pos + node.nodeSize, to)));
            if (b <= a) return false;
            change ??= schema.marks.formatChange.create({ id: newAnnotationId(opts.context()), author, date: nowIso(), before: "" });
            out.addMark(a, b, schema.marks.formatChange.create({ ...change.attrs, before: runPropsXml(node.marks, opts.context()) }));
            return false;
          });
          continue;
        }
        if (!(step instanceof ReplaceStep)) continue;
        const { from, to } = step as unknown as { from: number; to: number };
        const slice = (step as unknown as { slice: Slice }).slice;
        // What was put in: marked as inserted, with any new paragraph marks.
        if (slice.size) {
          const a = out.mapping.map(rest.map(from, -1), -1);
          const b = out.mapping.map(rest.map(from + slice.size, 1), 1);
          const left = a > 0 ? out.doc.resolve(a).nodeBefore?.marks : undefined;
          const mark = changeMark("insertion", left, opts);
          out.removeMark(a, b, schema.marks.deletion);
          out.addMark(a, b, mark);
          out.doc.nodesBetween(a, b, (node, pos) => {
            if (node.type === schema.nodes.table_row) {
              if (pos >= a && pos + node.nodeSize <= b && !/<w:(ins|del)\b/.test(node.attrs.trPr)) out.setNodeMarkup(pos, null, { ...node.attrs, trPr: node.attrs.trPr + rowMark("ins", opts) });
              return true;
            }
            if (!node.isTextblock) return true;
            const end = pos + node.nodeSize - 1;
            if (end >= a && end < b && !node.attrs.change && !node.attrs.propsChange) out.setNodeMarkup(pos, null, { ...node.attrs, change: { kind: "ins", id: newAnnotationId(opts.context()), author, date: nowIso() } });
            return false;
          });
        }
        // What was taken out: put back, marked as deleted.
        if (to > from) {
          const removed = before.slice(from, to);
          // Whole table rows: they stay, marked as deleted (unless they were your own new rows).
          if (removed.openStart === 0 && removed.content.childCount && removed.content.content.every((n) => n.type === schema.nodes.table_row)) {
            const rows = removed.content.content.filter((row) => !new RegExp(`<w:ins\\b[^>]*w:author="${author.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"`).test(row.attrs.trPr));
            if (!rows.length) continue;
            let deletion: Mark | null = null;
            const mark = () => (deletion ??= changeMark("deletion", undefined, opts));
            const kept = rows.map((row) => row.type.create({ ...row.attrs, trPr: /<w:del\b/.test(row.attrs.trPr) ? row.attrs.trPr : row.attrs.trPr + rowMark("del", opts) }, markDeleted(row.content, mark, author)));
            const at = out.mapping.map(rest.map(from, -1), -1);
            out.replace(at, at, new Slice(Fragment.from(kept), 0, 0));
            continue;
          }
          if (!onlyText(removed.content)) continue;
          // Joining a paragraph you added yourself to the next one just joins them.
          const $from = before.resolve(from);
          const ownParagraph = $from.parent.attrs.change?.kind === "ins" && $from.parent.attrs.change.author === author;
          let deletion: Mark | null = null;
          const mark = () => (deletion ??= changeMark("deletion", $from.nodeBefore?.marks, opts));
          const kept = markDeleted(removed.content, mark, author);
          if (!hasText(kept) && (removed.openStart === 0 || ownParagraph)) {
            // The joined paragraph ends with the mark of the second one, which wasn't new.
            const $joined = out.doc.resolve(out.mapping.map(rest.map(from)));
            if (ownParagraph && $joined.parent.attrs.change) out.setNodeMarkup($joined.before(), null, { ...$joined.parent.attrs, change: null });
            continue;
          }
          const at = out.mapping.map(rest.map(from, -1), -1);
          const size = out.doc.content.size;
          out.replace(at, at, new Slice(kept, removed.openStart, removed.openEnd));
          const added = out.doc.content.size - size;
          // Paragraph marks inside what was deleted are deleted too.
          out.doc.nodesBetween(at, at + added, (node, pos) => {
            if (!node.isTextblock) return true;
            const end = pos + node.nodeSize - 1;
            if (end >= at && end < at + added && node.attrs.change?.kind !== "del") {
              out.setNodeMarkup(pos, null, { ...node.attrs, change: { kind: "del", id: newAnnotationId(opts.context()), author, date: nowIso() } });
            }
            return false;
          });
          // Backspace leaves the cursor before what it deleted, Delete after it.
          if (!slice.size && oldState.selection.empty) selection = oldState.selection.head >= to ? at : at + added;
        }
      }
      if (selection !== null) out.setSelection(TextSelection.near(out.doc.resolve(Math.min(selection, out.doc.content.size))));
      return out.docChanged ? out.setMeta("noTrack", true) : null;
    },
  });
}
