// Turns a document's styles into CSS, so text formatted by a style (a
// heading, a quote, the document's default font) looks the way the style says.

import { HIGHLIGHTS, resolvedPara, resolvedRun, styleChain, type DocContext, type ParaProps, type RunProps } from "../docx/model";
import { cssFontFamily, lineHeightFor } from "./schema";

const pt = (value: number) => `${Math.round(value * 100) / 100}pt`;

function runCss(run: RunProps): string[] {
  const css: string[] = [];
  if (run.font) css.push(`font-family:${cssFontFamily(run.font)}`);
  if (run.size !== undefined) css.push(`font-size:${pt(run.size)}`);
  if (run.bold !== undefined) css.push(`font-weight:${run.bold ? "bold" : "normal"}`);
  if (run.italic !== undefined) css.push(`font-style:${run.italic ? "italic" : "normal"}`);
  if (run.color) css.push(`color:${run.color}`);
  if (run.highlight) css.push(`background-color:${HIGHLIGHTS[run.highlight] ?? run.highlight}`);
  if (run.vertAlign) css.push(run.vertAlign === "superscript" ? "vertical-align:super;font-size:smaller" : "vertical-align:sub;font-size:smaller");
  const lines = [run.underline && "underline", run.strike && "line-through"].filter(Boolean);
  if (run.underline !== undefined || run.strike !== undefined) css.push(`text-decoration-line:${lines.length ? lines.join(" ") : "none"}`);
  return css;
}

/** A line height, rounded up to the line grid when the document has one. */
function onGrid(lineHeight: string, grid: number | null): string {
  if (!grid) return lineHeight;
  const length = /^[\d.]+$/.test(lineHeight) ? `calc(${lineHeight}em)` : lineHeight;
  return `round(up, ${length}, ${grid}pt)`;
}

function paraCss(para: ParaProps, grid: number | null = null): string[] {
  const css: string[] = [];
  css.push(`text-align:${para.align ?? "left"}`);
  // On a line grid the space between paragraphs is whole grid steps too, so every line stays on the grid.
  const space = (points: number) => (grid && points > 0 ? `round(up, ${pt(points)}, ${grid}pt)` : pt(points));
  css.push(`padding-top:${space(para.spaceBefore ?? 0)}`);
  css.push(`padding-bottom:${space(para.spaceAfter ?? 0)}`);
  css.push(`padding-left:${pt(para.indentLeft ?? 0)}`);
  css.push(`padding-right:${pt(para.indentRight ?? 0)}`);
  css.push(`text-indent:${pt(para.indentFirst ?? 0)}`);
  if (para.lineSpacing !== undefined) css.push(`line-height:${onGrid(lineHeightFor(para.lineSpacing), grid)}`);
  else if (para.lineHeight !== undefined) css.push(para.lineRule === "atLeast" ? `line-height:${onGrid(`max(${pt(para.lineHeight)}, ${lineHeightFor(1)}em)`, grid)}` : `line-height:${pt(para.lineHeight)}`);
  else css.push(`line-height:${onGrid(lineHeightFor(1), grid)}`);
  return css;
}

type RunPropsBorders = import("../docx/model").TableBorders | undefined;

const escape = (id: string) => id.replace(/["\\]/g, "\\$&");

/** CSS for the editor, scoped to `scope` (the editable area). */
export function stylesheetFor(ctx: DocContext, scope: string): string {
  const { styles } = ctx;
  const rules: string[] = [];
  const base = resolvedRun(styles, null);
  rules.push(`${scope} { ${runCss(base).join(";")} }`);
  const grid = ctx.page.lineGrid;
  rules.push(`${scope} .para:not([data-style]) { ${[...runCss(resolvedRun(styles, null)), ...paraCss(resolvedPara(styles, null), grid)].join(";")} }`);
  for (const style of styles.byId.values()) {
    if (style.type === "paragraph") {
      const css = [...runCss(resolvedRun(styles, style.id)), ...paraCss(resolvedPara(styles, style.id), grid)];
      const selector = `${scope} .para[data-style="${escape(style.id)}"]`;
      rules.push(`${selector} { ${css.join(";")} }`);
      if (resolvedPara(styles, style.id).contextual) {
        rules.push(`${selector}:has(+ .para[data-style="${escape(style.id)}"]) { padding-bottom: 0 }`);
        rules.push(`${selector} + ${selector} { padding-top: 0 }`);
      }
    } else if (style.type === "table") {
      // A table style draws (or leaves out) the lines, and sets the spacing of the paragraphs in it.
      const chain = styleChain(styles, style.id);
      const borders = chain.reduce<RunPropsBorders>((b, s) => (s.tableBorders !== undefined ? s.tableBorders : b), undefined);
      const selector = `${scope} table[data-style="${escape(style.id)}"]`;
      if (borders) rules.push(`${selector}:not(.no-lines) > tbody > tr > :is(td, th) { border: ${Math.max(0.5, borders.size)}pt solid ${borders.color} }`);
      else if (borders === null) rules.push(`${selector}:not(.lines) > tbody > tr > :is(td, th) { border-color: transparent }`);
      const para = Object.assign({}, ...chain.map((s) => s.para)) as ParaProps;
      const css: string[] = [];
      if (para.spaceAfter !== undefined) css.push(`padding-bottom:${pt(para.spaceAfter)}`);
      if (para.spaceBefore !== undefined) css.push(`padding-top:${pt(para.spaceBefore)}`);
      if (para.lineSpacing !== undefined) css.push(`line-height:${lineHeightFor(para.lineSpacing)}`);
      if (css.length) rules.push(`${selector} .para:not([data-style]) { ${css.join(";")} }`);
    } else if (style.type === "character") {
      const run: RunProps = Object.assign({}, ...styleChain(styles, style.id).map((s) => s.run));
      const css = runCss(run);
      if (css.length) rules.push(`${scope} [data-cstyle="${escape(style.id)}"] { ${css.join(";")} }`);
    }
  }
  return rules.join("\n");
}
