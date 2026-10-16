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
  const lines = [run.underline && "underline", run.strike && "line-through"].filter(Boolean);
  if (run.underline !== undefined || run.strike !== undefined) css.push(`text-decoration-line:${lines.length ? lines.join(" ") : "none"}`);
  return css;
}

function paraCss(para: ParaProps): string[] {
  const css: string[] = [];
  css.push(`text-align:${para.align ?? "left"}`);
  css.push(`padding-top:${pt(para.spaceBefore ?? 0)}`);
  css.push(`padding-bottom:${pt(para.spaceAfter ?? 0)}`);
  css.push(`padding-left:${pt(para.indentLeft ?? 0)}`);
  css.push(`padding-right:${pt(para.indentRight ?? 0)}`);
  css.push(`text-indent:${pt(para.indentFirst ?? 0)}`);
  if (para.lineSpacing !== undefined) css.push(`line-height:${lineHeightFor(para.lineSpacing)}`);
  else if (para.lineHeight !== undefined) css.push(para.lineRule === "atLeast" ? `line-height:max(${pt(para.lineHeight)}, ${lineHeightFor(1)})` : `line-height:${pt(para.lineHeight)}`);
  else css.push(`line-height:${lineHeightFor(1)}`);
  return css;
}

const escape = (id: string) => id.replace(/["\\]/g, "\\$&");

/** CSS for the editor, scoped to `scope` (the editable area). */
export function stylesheetFor(ctx: DocContext, scope: string): string {
  const { styles } = ctx;
  const rules: string[] = [];
  const base = resolvedRun(styles, null);
  rules.push(`${scope} { ${runCss(base).join(";")} }`);
  rules.push(`${scope} .para:not([data-style]) { ${[...runCss(resolvedRun(styles, null)), ...paraCss(resolvedPara(styles, null))].join(";")} }`);
  for (const style of styles.byId.values()) {
    if (style.type === "paragraph") {
      const css = [...runCss(resolvedRun(styles, style.id)), ...paraCss(resolvedPara(styles, style.id))];
      const selector = `${scope} .para[data-style="${escape(style.id)}"]`;
      rules.push(`${selector} { ${css.join(";")} }`);
      if (resolvedPara(styles, style.id).contextual) {
        rules.push(`${selector}:has(+ .para[data-style="${escape(style.id)}"]) { padding-bottom: 0 }`);
        rules.push(`${selector} + ${selector} { padding-top: 0 }`);
      }
    } else if (style.type === "character") {
      const run: RunProps = Object.assign({}, ...styleChain(styles, style.id).map((s) => s.run));
      const css = runCss(run);
      if (css.length) rules.push(`${scope} [data-cstyle="${escape(style.id)}"] { ${css.join(";")} }`);
    }
  }
  return rules.join("\n");
}
