import * as docx from "docx";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { parseMath, toLatex, MathSyntaxError } from "../src/math/ast";
import { toMathML } from "../src/math/mathml";
import { fromOmml, toOmml, MATH_NS } from "../src/math/omml";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { createState } from "../src/editor/editor";
import { schema } from "../src/editor/schema";

const EXAMPLES = [
  "x^2+y^2=r^2",
  "a_{ij}",
  "\\frac{a+b}{c}",
  "\\sqrt{x}+\\sqrt[3]{y}",
  "\\alpha+\\beta\\le\\gamma",
  "\\sum_{i=1}^{n} x_i",
  "\\int_0^1 f(x)dx",
  "\\lim_{x\\to 0}\\frac{\\sin x}{x}=1",
  "\\left( \\frac{1}{2} \\right)",
  "\\vec{F}=m\\vec{a}",
  "\\begin{pmatrix} a & b \\\\ c & d \\end{pmatrix}",
  "E=mc^2",
  "\\Delta G=\\Delta H-T\\Delta S",
  "\\text{rate}=k[A]^2",
];

/** Parses OMML text the way the reader does. */
function omml(xml: string): Element {
  const doc = new DOMParser().parseFromString(
    `<root xmlns:m="${MATH_NS}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${xml}</root>`,
    "application/xml",
  );
  return doc.documentElement.firstElementChild!;
}

describe("the equation shorthand", () => {
  it("understands the usual LaTeX-like notation", () => {
    expect(parseMath("x^2")).toEqual({ t: "scripts", base: { t: "id", v: "x" }, sup: { t: "num", v: "2" } });
    expect(parseMath("\\frac{1}{2}")).toEqual({ t: "frac", num: { t: "num", v: "1" }, den: { t: "num", v: "2" } });
    expect(parseMath("\\alpha")).toEqual({ t: "id", v: "α" });
    const sum = parseMath("\\sum_{i=1}^{n} x_i + 1");
    expect(sum.t === "row" && sum.items[0].t === "nary" && sum.items[0].body).toEqual({ t: "scripts", base: { t: "id", v: "x" }, sub: { t: "id", v: "i" } });
  });

  it("gives a plain message for mistakes", () => {
    expect(() => parseMath("\\frac{1}{")).toThrow(MathSyntaxError);
    expect(() => parseMath("\\notacommand")).toThrow(/Unknown command \\notacommand/);
    expect(() => parseMath("x^")).toThrow(/missing/);
  });

  it("writes the shorthand back the same way", () => {
    for (const example of EXAMPLES) expect(parseMath(toLatex(parseMath(example))), example).toEqual(parseMath(example));
  });

  it("draws equations as MathML", () => {
    const math = toMathML(parseMath("\\frac{a}{b}+x^2"), true);
    expect(math.getAttribute("display")).toBe("block");
    expect(math.querySelector("mfrac")).not.toBeNull();
    expect(math.querySelector("msup mn")?.textContent).toBe("2");
  });
});

describe("Word equations", () => {
  it("survive being written as OMML and read back", () => {
    for (const example of EXAMPLES) {
      const tree = parseMath(example);
      const back = fromOmml(omml(toOmml(tree, false, false)));
      expect(toLatex(back), example).toBe(toLatex(tree));
    }
  });

  it("reads equations from a Word file, and saves them unchanged when not edited", async () => {
    const d = new docx.Document({
      sections: [
        {
          children: [
            new docx.Paragraph({
              children: [
                new docx.TextRun("Energy: "),
                new docx.Math({
                  children: [
                    new docx.MathRun("E=m"),
                    new docx.MathSuperScript({ children: [new docx.MathRun("c")], superScript: [new docx.MathRun("2")] }),
                  ],
                }),
              ],
            }),
            new docx.Paragraph({
              children: [new docx.Math({ children: [new docx.MathFraction({ numerator: [new docx.MathRun("1")], denominator: [new docx.MathRun("2")] })] })],
            }),
          ],
        },
      ],
    });
    const bytes = new Uint8Array(await docx.Packer.toBuffer(d));
    const { doc, context } = await readDocx(bytes);
    const maths: string[] = [];
    doc.descendants((n) => {
      if (n.type.name === "math") maths.push(n.attrs.latex);
    });
    expect(maths).toEqual(["E=mc^2", "\\frac{1}{2}"]);

    const before = await (await JSZip.loadAsync(bytes)).file("word/document.xml")!.async("string");
    const after = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/document.xml")!.async("string");
    const first = before.slice(before.indexOf("<m:oMath>"), before.indexOf("</m:oMath>") + 10);
    expect(after).toContain(first);
  });

  it("saves new and edited equations as real Word equations", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    let state = createState(doc, () => context);
    const eq = schema.nodes.math.create({ latex: "\\sum_{i=1}^{n} i=\\frac{n(n+1)}{2}", display: true });
    state = state.apply(state.tr.insert(1, eq));
    const saved = await JSZip.loadAsync(await writeDocx(state.doc, context));
    const xml = await saved.file("word/document.xml")!.async("string");
    expect(xml).toContain("<m:oMathPara><m:oMath><m:nary>");
    expect(xml).toContain('<m:chr m:val="∑"/>');
    expect(xml).toContain("<m:f><m:num>");
    const again = await readDocx(await saved.generateAsync({ type: "uint8array" }));
    const math = again.doc.firstChild!.firstChild!;
    expect(math.attrs.display).toBe(true);
    expect(parseMath(math.attrs.latex)).toEqual(parseMath("\\sum_{i=1}^{n} i=\\frac{n(n+1)}{2}"));
  });
});
