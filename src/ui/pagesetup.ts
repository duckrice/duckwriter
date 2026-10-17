// Page setup: paper size, orientation and margins, in centimetres.

import type { PageSetup } from "../docx/model";
import { h } from "../dom";

const CM = 72 / 2.54;

/** Paper sizes, portrait, in points. */
export const PAPERS: { name: string; width: number; height: number }[] = [
  { name: "A4", width: 595.3, height: 841.9 },
  { name: "A5", width: 419.55, height: 595.3 },
  { name: "A3", width: 841.9, height: 1190.55 },
  { name: "Letter", width: 612, height: 792 },
  { name: "Legal", width: 612, height: 1008 },
];

const near = (a: number, b: number) => Math.abs(a - b) < 3;

/** The paper a page is printed on, whichever way round it is. */
export function paperOf(page: { width: number; height: number }): { name: string; landscape: boolean } {
  const landscape = page.width > page.height;
  const short = Math.min(page.width, page.height);
  const long = Math.max(page.width, page.height);
  const paper = PAPERS.find((p) => near(p.width, short) && near(p.height, long));
  return { name: paper?.name ?? "Custom", landscape };
}

const cm = (points: number) => (Math.round((points / CM) * 100) / 100).toString();

/** Shows the page setup dialog; `apply` gets the new page when OK is chosen. */
export function pageSetup(page: PageSetup, apply: (page: PageSetup) => void) {
  const current = paperOf(page);
  const paper = h("select", { class: "tb-select" }, ...PAPERS.map((p) => h("option", { value: p.name }, `${p.name} (${cm(p.width)} × ${cm(p.height)} cm)`)), h("option", { value: "Custom" }, "As it is")) as HTMLSelectElement;
  paper.value = current.name;
  const orientation = h("select", { class: "tb-select" }, h("option", { value: "portrait" }, "Portrait (upright)"), h("option", { value: "landscape" }, "Landscape (sideways)")) as HTMLSelectElement;
  orientation.value = current.landscape ? "landscape" : "portrait";
  const box = (points: number, label: string) => h("input", { class: "tb-input dialog-input number", type: "number", min: "0", max: "10", step: "0.1", value: cm(points), "aria-label": `${label} margin, cm` }) as HTMLInputElement;
  const top = box(page.marginTop, "Top");
  const bottom = box(page.marginBottom, "Bottom");
  const left = box(page.marginLeft, "Left");
  const right = box(page.marginRight, "Right");
  const presets = h(
    "div",
    { class: "margin-presets" },
    ...[
      { name: "Normal", t: 2.54, s: 2.54 },
      { name: "Narrow", t: 1.27, s: 1.27 },
      { name: "Wide", t: 2.54, s: 5.08 },
    ].map((p) =>
      h("button", {
        type: "button",
        class: "btn small",
        onclick: () => {
          top.value = bottom.value = String(p.t);
          left.value = right.value = String(p.s);
        },
      }, p.name),
    ),
  );

  const ok = () => {
    const chosen = PAPERS.find((p) => p.name === paper.value);
    const short = chosen ? chosen.width : Math.min(page.width, page.height);
    const long = chosen ? chosen.height : Math.max(page.width, page.height);
    const landscape = orientation.value === "landscape";
    const value = (input: HTMLInputElement, fallback: number) => {
      const n = Number(input.value.replace(",", "."));
      return Number.isFinite(n) && n >= 0 ? n * CM : fallback;
    };
    const next: PageSetup = {
      ...page,
      width: landscape ? long : short,
      height: landscape ? short : long,
      marginTop: value(top, page.marginTop),
      marginBottom: value(bottom, page.marginBottom),
      marginLeft: value(left, page.marginLeft),
      marginRight: value(right, page.marginRight),
    };
    // Some room must be left for the text.
    if (next.width - next.marginLeft - next.marginRight < 2 * CM || next.height - next.marginTop - next.marginBottom < 2 * CM) {
      error.textContent = "The margins leave no room for the text.";
      return;
    }
    dialog.close();
    apply(next);
  };
  const error = h("p", { class: "dialog-error", role: "alert" });
  const dialog = h(
    "dialog",
    { class: "dialog" },
    h("h2", {}, "Page setup"),
    h(
      "div",
      { class: "dialog-form" },
      h("label", {}, "Paper", paper),
      h("label", {}, "Orientation", orientation),
      h("label", {}, "Top (cm)", top),
      h("label", {}, "Bottom (cm)", bottom),
      h("label", {}, "Left (cm)", left),
      h("label", {}, "Right (cm)", right),
      h("label", {}, "Margins", presets),
    ),
    error,
    h(
      "div",
      { class: "dialog-footer" },
      h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"),
      h("button", { type: "button", class: "btn primary", onclick: ok }, "OK"),
    ),
  );
  dialog.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && (event.target as HTMLElement).tagName === "INPUT") {
      event.preventDefault();
      ok();
    }
  });
  dialog.addEventListener("close", () => dialog.remove());
  document.body.append(dialog);
  dialog.showModal();
  paper.focus();
}
