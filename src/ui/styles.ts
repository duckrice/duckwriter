// Changing a style: every paragraph in it changes at once (all the Heading 1s
// in another font, say), and Word sees the change too. "Modify style…" or
// "Update the style to match" the paragraph the cursor is in (Alt+Shift+Y).

import type { EditorView } from "prosemirror-view";
import { resolvedPara, resolvedRun, type DocContext, type ParaProps, type RunProps } from "../docx/model";
import { h } from "../dom";
import { formattingAt } from "../editor/commands";
import { schema } from "../editor/schema";

/** Changes a style's formatting; it is written into the document's styles when saved. */
export function changeStyle(ctx: DocContext, id: string, run: RunProps, para: ParaProps) {
  const style = ctx.styles.byId.get(id);
  if (!style) return;
  style.run = { ...style.run, ...run };
  style.para = { ...style.para, ...para };
  const before = ctx.styleChanges.get(id) ?? { run: {}, para: {} };
  ctx.styleChanges.set(id, { run: { ...before.run, ...run }, para: { ...before.para, ...para } });
  ctx.styleVersion = (ctx.styleVersion ?? 0) + 1;
}

/** The style of the paragraph the cursor is in. */
function currentStyle(view: EditorView, ctx: DocContext): string | null {
  const { paragraph } = formattingAt(ctx, view.state);
  return paragraph?.attrs.style ?? ctx.styles.defaultParagraph;
}

/** The style takes the formatting of the paragraph the cursor is in, which then has none of its own. */
export function updateStyleToMatch(view: EditorView, ctx: DocContext): string | null {
  const id = currentStyle(view, ctx);
  if (!id) return null;
  const { run, para, paragraph } = formattingAt(ctx, view.state);
  const pick = <T extends object>(o: T, keys: (keyof T)[]) => Object.fromEntries(keys.filter((k) => o[k] !== undefined).map((k) => [k, o[k]])) as Partial<T>;
  changeStyle(ctx, id, pick(run, ["font", "size", "bold", "italic", "underline", "color"]), pick(para, ["align", "spaceBefore", "spaceAfter", "lineSpacing"]));
  // The paragraph's own formatting is the style's now.
  if (paragraph) {
    const tr = view.state.tr;
    view.state.doc.descendants((node, pos) => {
      if (node !== paragraph) return true;
      tr.setNodeMarkup(pos, null, { ...node.attrs, align: null, spaceBefore: null, spaceAfter: null, lineSpacing: null });
      for (const name of ["font", "size", "bold", "italic", "underline", "color"] as const) tr.removeMark(pos, pos + node.nodeSize, schema.marks[name]);
      return false;
    });
    view.dispatch(tr);
  }
  return id;
}

/** The dialog for a style's font, size, emphasis, colour, alignment and spacing. */
export function modifyStyle(view: EditorView, ctx: DocContext, onDone: (id: string) => void) {
  const id = currentStyle(view, ctx);
  const style = id ? ctx.styles.byId.get(id) : undefined;
  if (!id || !style) return;
  const run = resolvedRun(ctx.styles, id);
  const para = resolvedPara(ctx.styles, id);
  const input = (attrs: Record<string, string>, value: string) => {
    const el = h("input", { class: "tb-input dialog-input", ...attrs }) as HTMLInputElement;
    el.value = value;
    return el;
  };
  const font = input({ list: "dw-fonts", "aria-label": "Font" }, run.font ?? "");
  const size = input({ type: "number", min: "1", max: "400", step: "0.5", class: "tb-input dialog-input number", "aria-label": "Size" }, String(run.size ?? 11));
  const box = (on: boolean | undefined) => {
    const el = h("input", { type: "checkbox" }) as HTMLInputElement;
    el.checked = !!on;
    return el;
  };
  const bold = box(run.bold);
  const italic = box(run.italic);
  const underline = box(run.underline);
  const color = input({ type: "color", "aria-label": "Colour" }, run.color && /^#[0-9a-f]{6}$/i.test(run.color) ? run.color : "#000000");
  const align = h("select", { class: "tb-select" }, ...["left", "center", "right", "justify"].map((a) => h("option", { value: a }, a === "center" ? "Centre" : a[0].toUpperCase() + a.slice(1)))) as HTMLSelectElement;
  align.value = para.align ?? "left";
  const before = input({ type: "number", min: "0", max: "200", step: "1", class: "tb-input dialog-input number", "aria-label": "Space before" }, String(Math.round(para.spaceBefore ?? 0)));
  const after = input({ type: "number", min: "0", max: "200", step: "1", class: "tb-input dialog-input number", "aria-label": "Space after" }, String(Math.round(para.spaceAfter ?? 0)));
  const lines = h("select", { class: "tb-select" }, ...[1, 1.15, 1.5, 2].map((n) => h("option", { value: String(n) }, String(n)))) as HTMLSelectElement;
  lines.value = String([1, 1.15, 1.5, 2].find((n) => Math.abs(n - (para.lineSpacing ?? 1)) < 0.05) ?? 1);

  const ok = () => {
    changeStyle(
      ctx,
      id,
      { font: font.value.trim() || undefined, size: Number(size.value) || run.size, bold: bold.checked, italic: italic.checked, underline: underline.checked, color: color.value === "#000000" && !run.color ? undefined : color.value },
      { align: align.value as ParaProps["align"], spaceBefore: Number(before.value) || 0, spaceAfter: Number(after.value) || 0, lineSpacing: Number(lines.value) },
    );
    dialog.close();
    onDone(id);
  };
  const dialog = h(
    "dialog",
    { class: "dialog" },
    h("h2", {}, `Style: ${style.name}`),
    h("p", { class: "dialog-hint" }, "Every paragraph in this style changes. Word sees the change too."),
    h(
      "div",
      { class: "dialog-form" },
      h("label", {}, "Font", font),
      h("label", {}, "Size (pt)", size),
      h("label", {}, "Emphasis", h("span", { class: "style-checks" }, h("label", {}, bold, "Bold"), h("label", {}, italic, "Italic"), h("label", {}, underline, "Underline"))),
      h("label", {}, "Colour", color),
      h("label", {}, "Alignment", align),
      h("label", {}, "Space before (pt)", before),
      h("label", {}, "Space after (pt)", after),
      h("label", {}, "Line spacing", lines),
    ),
    h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "Cancel"), h("button", { type: "button", class: "btn primary", onclick: ok }, "OK")),
  ) as HTMLDialogElement;
  dialog.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && (event.target as HTMLElement).tagName === "INPUT" && (event.target as HTMLInputElement).type !== "checkbox") {
      event.preventDefault();
      ok();
    }
  });
  dialog.addEventListener("close", () => {
    dialog.remove();
    view.focus();
  });
  document.body.append(dialog);
  dialog.showModal();
  font.focus();
}
