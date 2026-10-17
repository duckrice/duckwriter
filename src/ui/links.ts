// Links (Ctrl+K): to a web page, an e-mail address, or a heading in the
// document. Ctrl+click follows them, as in Word.

import type { Mark } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { h } from "../dom";
import { ensureStyles } from "../docx/builtins";
import { schema } from "../editor/schema";
import { bookmarkFor, referenceTargets } from "./references";

export interface LinkRange {
  from: number;
  to: number;
  mark: Mark;
}

/** The whole link the cursor (or the start of the selection) is in. */
export function linkAt(state: EditorState): LinkRange | null {
  const { $from, from } = state.selection;
  const parent = $from.parent;
  if (!parent.isTextblock) return null;
  const start = $from.start();
  let found: LinkRange | null = null;
  let run: LinkRange | null = null;
  parent.forEach((child, offset) => {
    const pos = start + offset;
    const mark = schema.marks.link.isInSet(child.marks);
    if (mark && run && run.to === pos && run.mark.eq(mark)) run.to = pos + child.nodeSize;
    else run = mark ? { from: pos, to: pos + child.nodeSize, mark } : null;
    if (run && run.from <= from && run.to >= from && !found) found = run;
    // A later piece of the same link extends what was found.
    if (found && run && run.from === found.from) found.to = run.to;
  });
  return found;
}

/** A typed address as a link: "duck.no" is a web page, "kari@duck.no" an e-mail address. */
export function normalizeAddress(text: string): string {
  const t = text.trim();
  if (!t) return "";
  if (/^[^\s@/:]+@[^\s@]+\.[^\s@]+$/.test(t)) return `mailto:${t}`;
  if (/^[a-z][a-z0-9+.-]*:/i.test(t)) return t;
  return `https://${t}`;
}

/** Shows the link dialog: insert a link, or change or remove the one at the cursor. */
export function editLink(view: EditorView, ctx: DocContext) {
  const state = view.state;
  const existing = linkAt(state);
  const { from, to, empty } = state.selection;
  const range = existing ?? (empty ? null : { from, to });
  const shown = range ? state.doc.textBetween(range.from, range.to, " ") : "";
  const targets = referenceTargets(ctx, state.doc).filter((t) => t.kind === "heading");

  const text = h("input", { class: "tb-input dialog-input", placeholder: "The text in the document", value: shown }) as HTMLInputElement;
  const address = h("input", { class: "tb-input dialog-input", placeholder: "https://… or name@example.com", value: existing?.mark.attrs.href.replace(/^mailto:/, "") ?? "" }) as HTMLInputElement;
  const place = h(
    "select",
    { class: "tb-select" },
    h("option", { value: "" }, "—"),
    ...targets.map((t, i) => h("option", { value: String(i) }, t.label.length > 60 ? `${t.label.slice(0, 60)}…` : t.label)),
  ) as HTMLSelectElement;
  place.addEventListener("change", () => {
    if (place.value) address.value = "";
  });
  address.addEventListener("input", () => {
    if (address.value) place.value = "";
  });

  const apply = () => {
    const target = place.value ? targets[Number(place.value)] : null;
    const href = target ? "" : normalizeAddress(address.value);
    if (!target && !href) return address.focus();
    const tr = view.state.tr;
    const anchor = target ? bookmarkFor(tr, target) : null;
    const mark = schema.marks.link.create({ href, anchor, rid: null });
    // Blue and underlined, with Word's own style for links.
    ensureStyles(ctx, "Hyperlink");
    const style = schema.marks.charStyle.create({ id: "Hyperlink" });
    const label = text.value || (target ? target.label : address.value.trim());
    if (range) {
      const a = tr.mapping.map(range.from);
      const b = tr.mapping.map(range.to);
      if (label !== shown) {
        const marks = tr.doc.resolve(a + 1).marks().filter((m) => m.type !== schema.marks.link);
        tr.replaceWith(a, b, schema.text(label, style.addToSet(mark.addToSet(marks))));
      } else {
        tr.removeMark(a, b, schema.marks.link);
        tr.addMark(a, b, mark);
        if (!existing) tr.addMark(a, b, style);
      }
    } else {
      const at = tr.mapping.map(from);
      const marks = (view.state.storedMarks ?? view.state.selection.$from.marks()).filter((m) => m.type !== schema.marks.link);
      tr.insert(at, schema.text(label, style.addToSet(mark.addToSet(marks))));
    }
    dialog.close();
    view.dispatch(tr.scrollIntoView());
  };
  const remove = () => {
    dialog.close();
    if (!existing) return;
    const tr = view.state.tr.removeMark(existing.from, existing.to, schema.marks.link);
    // Not blue and underlined any more either.
    const style = schema.marks.charStyle.create({ id: "Hyperlink" });
    if (tr.doc.rangeHasMark(existing.from, existing.to, style)) tr.removeMark(existing.from, existing.to, style);
    view.dispatch(tr);
  };

  const dialog = h(
    "dialog",
    { class: "dialog" },
    h("h2", {}, existing ? "Edit link" : "Insert link"),
    h(
      "div",
      { class: "dialog-form" },
      h("label", {}, "Text", text),
      h("label", {}, "Address", address),
      targets.length ? h("label", {}, "Or a heading", place) : null,
    ),
    h(
      "div",
      { class: "dialog-footer" },
      existing ? h("button", { type: "button", class: "btn", onclick: remove }, "Remove link") : null,
      h("span", { class: "dialog-spacer" }),
      h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"),
      h("button", { type: "button", class: "btn primary", onclick: apply }, "OK"),
    ),
  );
  for (const input of [text, address]) {
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        apply();
      }
    });
  }
  dialog.addEventListener("close", () => {
    dialog.remove();
    view.focus();
  });
  document.body.append(dialog);
  dialog.showModal();
  (shown ? address : text).focus();
}
