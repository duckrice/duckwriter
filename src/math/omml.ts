// Word's equation format (Office Math, "OMML") to and from equation trees.

import { escapeAttr, escapeText } from "../docx/xml";
import { FUNCTIONS, LIMITS, isEmpty, row, type MathNode } from "./ast";

export const MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";

// ---- writing --------------------------------------------------------------------------

/** Word draws equations in Cambria Math; runs say so, as Word's own do. */
const FONT = '<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>';

/** A run of equation text: a variable (italic), a name like sin (upright), or ordinary words (normal text). */
const run = (text: string, kind: "math" | "upright" | "text" = "math") =>
  `<m:r>${kind === "upright" ? '<m:rPr><m:sty m:val="p"/></m:rPr>' : kind === "text" ? "<m:rPr><m:nor/></m:rPr>" : ""}${FONT}<m:t xml:space="preserve">${escapeText(text)}</m:t></m:r>`;

const arg = (tag: string, node: MathNode | undefined) => (node && !isEmpty(node) ? `<m:${tag}>${omml(node)}</m:${tag}>` : `<m:${tag}/>`);

function omml(node: MathNode): string {
  switch (node.t) {
    case "row": {
      // Neighbouring letters, digits and signs share one run, as in Word's own equations.
      let out = "";
      let text = "";
      for (const item of node.items) {
        if (item.t === "id" || item.t === "num" || item.t === "op") text += item.v;
        else {
          if (text) out += run(text);
          text = "";
          out += omml(item);
        }
      }
      return out + (text ? run(text) : "");
    }
    case "id":
    case "num":
    case "op":
      return run(node.v);
    case "fn":
      return run(node.v, "upright");
    case "text":
      return run(node.v, "text");
    case "frac":
      return `<m:f>${arg("num", node.num)}${arg("den", node.den)}</m:f>`;
    case "sqrt":
      return node.index
        ? `<m:rad>${arg("deg", node.index)}${arg("e", node.body)}</m:rad>`
        : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${arg("e", node.body)}</m:rad>`;
    case "scripts":
      if (node.sub && node.sup) return `<m:sSubSup>${arg("e", node.base)}${arg("sub", node.sub)}${arg("sup", node.sup)}</m:sSubSup>`;
      if (node.sub) return `<m:sSub>${arg("e", node.base)}${arg("sub", node.sub)}</m:sSub>`;
      return `<m:sSup>${arg("e", node.base)}${arg("sup", node.sup)}</m:sSup>`;
    case "nary": {
      const props = [
        node.op === "∫" ? "" : `<m:chr m:val="${node.op}"/>`,
        `<m:limLoc m:val="${"∫∬∭∮".includes(node.op) ? "subSup" : "undOvr"}"/>`,
        node.sub ? "" : '<m:subHide m:val="1"/>',
        node.sup ? "" : '<m:supHide m:val="1"/>',
      ].join("");
      return `<m:nary><m:naryPr>${props}</m:naryPr>${arg("sub", node.sub)}${arg("sup", node.sup)}${arg("e", node.body)}</m:nary>`;
    }
    case "limit":
      return isEmpty(node.under)
        ? `<m:func><m:fName>${run(node.name, "upright")}</m:fName><m:e/></m:func>`
        : `<m:limLow><m:e>${run(node.name, "upright")}</m:e>${arg("lim", node.under)}</m:limLow>`;
    case "delim":
      return `<m:d><m:dPr><m:begChr m:val="${escapeAttr(node.open)}"/><m:endChr m:val="${escapeAttr(node.close)}"/></m:dPr>${arg("e", node.body)}</m:d>`;
    case "accent":
      return node.chr === "̅"
        ? `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${arg("e", node.body)}</m:bar>`
        : `<m:acc><m:accPr><m:chr m:val="${node.chr}"/></m:accPr>${arg("e", node.body)}</m:acc>`;
    case "matrix": {
      const m = `<m:m>${node.rows.map((r) => `<m:mr>${r.map((cell) => arg("e", cell)).join("")}</m:mr>`).join("")}</m:m>`;
      return node.open || node.close ? `<m:d><m:dPr><m:begChr m:val="${escapeAttr(node.open)}"/><m:endChr m:val="${escapeAttr(node.close)}"/></m:dPr><m:e>${m}</m:e></m:d>` : m;
    }
  }
}

/** The equation as OMML: an inline <m:oMath>, or a displayed one wrapped in <m:oMathPara>. */
export function toOmml(node: MathNode, display: boolean, declareNamespace: boolean): string {
  const ns = declareNamespace ? ` xmlns:m="${MATH_NS}"` : "";
  const math = `<m:oMath${display ? "" : ns}>${omml(node)}</m:oMath>`;
  return display ? `<m:oMathPara${ns}>${math}</m:oMathPara>` : math;
}

// ---- reading --------------------------------------------------------------------------

const isM = (el: Element, name: string) => el.namespaceURI === MATH_NS && el.localName === name;
const kids = (el: Element) => Array.from(el.children);
const part = (el: Element | undefined, name: string) => el && kids(el).find((c) => isM(c, name));
const val = (el: Element | undefined, name: string) => {
  const p = el && part(el, name);
  return p ? (p.getAttributeNS(MATH_NS, "val") ?? p.getAttribute("m:val")) : null;
};

/** Reads an <m:oMath> or <m:oMathPara> into a tree. */
export function fromOmml(el: Element): MathNode {
  if (isM(el, "oMathPara")) return row(kids(el).filter((c) => isM(c, "oMath")).map(fromOmml));
  return readRow(el);
}

function readRow(el: Element | undefined): MathNode {
  if (!el) return row([]);
  const items: MathNode[] = [];
  for (const c of kids(el)) items.push(...readItem(c));
  return mergeRuns(items);
}

/** Splits text runs into variables, numbers, operators and words. */
function textItems(text: string, upright: boolean): MathNode[] {
  if (upright) return [FUNCTIONS.has(text) || LIMITS.has(text) || /^[a-zA-Z]{2,}$/.test(text) ? { t: "fn", v: text } : { t: "text", v: text }];
  const out: MathNode[] = [];
  for (const m of text.matchAll(/(\d+(?:[.,]\d+)?)|([a-zA-Z]+)|(\s+)|(.)/gsu)) {
    if (m[1]) out.push({ t: "num", v: m[1] });
    else if (m[2]) {
      // Word writes "sin" as an ordinary run; treat known names as functions.
      if (FUNCTIONS.has(m[2])) out.push({ t: "fn", v: m[2] });
      else for (const ch of m[2]) out.push({ t: "id", v: ch });
    } else if (m[3]) continue;
    else out.push(/[\p{L}]/u.test(m[4]) ? { t: "id", v: m[4] } : { t: "op", v: m[4] === "-" ? "−" : m[4] });
  }
  return out;
}

function mergeRuns(items: MathNode[]): MathNode {
  return items.length === 1 ? items[0] : { t: "row", items };
}

function readItem(c: Element): MathNode[] {
  if (c.namespaceURI !== MATH_NS) return [];
  const e = part(c, "e");
  switch (c.localName) {
    case "r": {
      const text = kids(c).filter((k) => isM(k, "t")).map((k) => k.textContent ?? "").join("");
      const rPr = part(c, "rPr");
      if (part(rPr, "nor")) return [{ t: "text", v: text }];
      return textItems(text, val(rPr, "sty") === "p");
    }
    case "f":
      return [{ t: "frac", num: readRow(part(c, "num")), den: readRow(part(c, "den")) }];
    case "rad": {
      const hidden = val(part(c, "radPr"), "degHide");
      const deg = readRow(part(c, "deg"));
      return [{ t: "sqrt", body: readRow(e), index: hidden === "1" || hidden === "on" || isEmpty(deg) ? undefined : deg }];
    }
    case "sSub":
      return [{ t: "scripts", base: readRow(e), sub: readRow(part(c, "sub")) }];
    case "sSup":
      return [{ t: "scripts", base: readRow(e), sup: readRow(part(c, "sup")) }];
    case "sSubSup":
      return [{ t: "scripts", base: readRow(e), sub: readRow(part(c, "sub")), sup: readRow(part(c, "sup")) }];
    case "sPre":
      return [{ t: "scripts", base: { t: "text", v: "" }, sub: readRow(part(c, "sub")), sup: readRow(part(c, "sup")) }, ...[readRow(e)]];
    case "nary": {
      const props = part(c, "naryPr");
      const sub = readRow(part(c, "sub"));
      const sup = readRow(part(c, "sup"));
      return [{ t: "nary", op: val(props, "chr") ?? "∫", sub: isEmpty(sub) ? undefined : sub, sup: isEmpty(sup) ? undefined : sup, body: readRow(e) }];
    }
    case "d": {
      const props = part(c, "dPr");
      const open = val(props, "begChr") ?? "(";
      const close = val(props, "endChr") ?? ")";
      const parts = kids(c).filter((k) => isM(k, "e")).map(readRow);
      const sep = val(props, "sepChr") ?? "|";
      const body = parts.length === 1 ? parts[0] : row(parts.flatMap((p, i) => (i ? [{ t: "op" as const, v: sep }, p] : [p])));
      // A matrix in brackets.
      if (body.t === "matrix" && !body.open && !body.close) return [{ ...body, open, close }];
      return [{ t: "delim", open, close, body }];
    }
    case "m":
      return [{ t: "matrix", open: "", close: "", rows: kids(c).filter((k) => isM(k, "mr")).map((mr) => kids(mr).filter((k) => isM(k, "e")).map(readRow)) }];
    case "acc":
      return [{ t: "accent", chr: val(part(c, "accPr"), "chr") ?? "̂", body: readRow(e) }];
    case "bar":
      return [{ t: "accent", chr: "̅", body: readRow(e) }];
    case "limLow": {
      const base = readRow(e);
      const name = base.t === "fn" ? base.v : base.t === "text" ? base.v : null;
      if (name && LIMITS.has(name)) return [{ t: "limit", name, under: readRow(part(c, "lim")) }];
      return [{ t: "scripts", base, sub: readRow(part(c, "lim")) }];
    }
    case "limUpp":
      return [{ t: "scripts", base: readRow(e), sup: readRow(part(c, "lim")) }];
    case "func": {
      const name = readRow(part(c, "fName"));
      return [name, readRow(e)];
    }
    case "groupChr":
    case "box":
    case "borderBox":
    case "phant":
      return [readRow(e)];
    case "eqArr":
      return [{ t: "matrix", open: "", close: "", rows: kids(c).filter((k) => isM(k, "e")).map((k) => [readRow(k)]) }];
    case "oMath":
      return [readRow(c)];
    default:
      // Properties and anything unknown: keep any text inside.
      return c.localName.endsWith("Pr") ? [] : textItems(c.textContent ?? "", false);
  }
}
