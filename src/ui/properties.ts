// The document's properties (Alt+Shift+P): title, subject, author, keywords
// and description, as Word shows them under File → Info, with who saved it
// last and when.

import type { DocContext } from "../docx/model";
import { h } from "../dom";

/** Asks for the properties; `changed` is called when any was edited. */
export function showProperties(ctx: DocContext, changed: () => void, done?: () => void) {
  const p = ctx.properties;
  const field = (label: string, key: "title" | "subject" | "creator" | "keywords" | "description", hint = "") => {
    const input = (key === "description" ? h("textarea", { class: "tb-input dialog-input wide", rows: 3 }) : h("input", { class: "tb-input dialog-input wide", placeholder: hint })) as HTMLInputElement | HTMLTextAreaElement;
    input.value = p[key];
    input.dataset.key = key;
    return h("label", {}, label, input);
  };
  const when = (iso: string) => {
    const d = iso ? new Date(iso) : null;
    return d && !Number.isNaN(d.getTime()) ? d.toLocaleString([], { dateStyle: "medium", timeStyle: "short" }) : "–";
  };
  const form = h(
    "div",
    { class: "dialog-form" },
    field("Title", "title"),
    field("Subject", "subject"),
    field("Author", "creator", "Your name"),
    field("Keywords", "keywords", "Separated by commas or semicolons"),
    field("Description", "description"),
  );
  const facts = h(
    "table",
    { class: "count-table" },
    h(
      "tbody",
      {},
      h("tr", {}, h("th", {}, "Last saved by"), h("td", {}, p.anonymous ? "(names left out)" : p.lastModifiedBy || "–")),
      h("tr", {}, h("th", {}, "Created"), h("td", {}, when(p.created))),
      h("tr", {}, h("th", {}, "Last saved"), h("td", {}, when(p.modified))),
      h("tr", {}, h("th", {}, "Revision"), h("td", {}, p.revision || "–")),
    ),
  );
  const save = () => {
    let edited = false;
    for (const input of form.querySelectorAll<HTMLInputElement | HTMLTextAreaElement>("[data-key]")) {
      const key = input.dataset.key as "title" | "subject" | "creator" | "keywords" | "description";
      const value = input.value.trim();
      if (p[key] !== value) {
        p[key] = value;
        edited = true;
        // Naming an author again: names are kept from now on.
        if (key === "creator" && value) p.anonymous = false;
      }
    }
    if (edited) changed();
  };
  const dialog = h(
    "dialog",
    { class: "dialog wide" },
    h("h2", {}, "Document properties"),
    form,
    facts,
    h("p", { class: "dialog-hint" }, "Saved in the document, as Word keeps them (File → Info in Word)."),
    h(
      "div",
      { class: "dialog-footer" },
      h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"),
      h(
        "button",
        {
          type: "button",
          class: "btn primary",
          onclick: () => {
            save();
            dialog.close();
          },
        },
        "OK",
      ),
    ),
  ) as HTMLDialogElement;
  dialog.addEventListener("close", () => {
    dialog.remove();
    done?.();
  });
  document.body.append(dialog);
  dialog.showModal();
  form.querySelector<HTMLElement>("input")?.focus();
}
