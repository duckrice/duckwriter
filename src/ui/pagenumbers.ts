// Page number format for a section, as Word's: i, ii, iii in the front matter and 1, 2, 3 from the
// first chapter, starting again at a number, and no number on a title page.

import { h } from "../dom";
import type { PageNumbering } from "../editor/commands";

const FORMATS: [PageNumbering["format"], string][] = [
  ["decimal", "1, 2, 3"],
  ["lowerRoman", "i, ii, iii"],
  ["upperRoman", "I, II, III"],
  ["lowerLetter", "a, b, c"],
  ["upperLetter", "A, B, C"],
];

/** Asks how this section's pages are numbered; `apply` gets the answer on OK. */
export function pageNumberFormat(current: PageNumbering, sections: number, apply: (n: PageNumbering) => void) {
  const format = h("select", { class: "tb-select" }, ...FORMATS.map(([value, label]) => h("option", { value }, label))) as HTMLSelectElement;
  format.value = current.format;
  const goOn = h("input", { type: "radio", name: "pn-start", checked: current.start === null }) as HTMLInputElement;
  const restart = h("input", { type: "radio", name: "pn-start", checked: current.start !== null }) as HTMLInputElement;
  const start = h("input", { class: "tb-input dialog-input number", type: "number", min: "0", value: String(current.start ?? 1), "aria-label": "Start at" }) as HTMLInputElement;
  start.addEventListener("focus", () => (restart.checked = true));
  const firstBlank = h("input", { type: "checkbox", checked: current.firstBlank }) as HTMLInputElement;
  const ok = () => {
    dialog.close();
    apply({ format: format.value as PageNumbering["format"], start: restart.checked ? Math.max(0, Math.round(Number(start.value) || 0)) : null, firstBlank: firstBlank.checked });
  };
  const dialog = h(
    "dialog",
    { class: "dialog" },
    h("h2", {}, "Page number format"),
    h(
      "p",
      { class: "dialog-hint" },
      sections > 1
        ? "For the section the cursor is in. Each section can have its own: i, ii, iii in the front matter, then 1, 2, 3 from the first chapter."
        : "For the whole document. For roman numbers first and 1, 2, 3 later, put a section break where the numbering changes (Layout → Layout options → Section break: next page), then set each section here.",
    ),
    h("div", { class: "dialog-form" }, h("label", {}, "Numbers", format)),
    h("label", { class: "check-field" }, goOn, sections > 1 ? " Go on from the section before" : " Start at 1"),
    h("label", { class: "check-field" }, restart, " Start at ", start),
    h("label", { class: "check-field" }, firstBlank, " No page number on the section's first page (a title page)"),
    h(
      "div",
      { class: "dialog-footer" },
      h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"),
      h("button", { type: "button", class: "btn primary", onclick: ok }, "OK"),
    ),
  );
  dialog.addEventListener("close", () => dialog.remove());
  document.body.append(dialog);
  dialog.showModal();
  format.focus();
}
