import { describe, expect, it } from "vitest";
import { symbolChar } from "../src/docx/read";
import { boxShape, horizontalLine, lineShape, schema } from "../src/editor/schema";
import { colorIn, drawShapes, presetPath, readDiagram } from "../src/editor/shapes";

// A cover page's rule as Word writes it: a straight connector, 2.5 pt, black, just under its paragraph.
const LINE = `<w:r><w:drawing><wp:anchor behindDoc="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>-276860</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>100965</wp:posOffset></wp:positionV><wp:extent cx="6311900" cy="0"/><wp:wrapTopAndBottom/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="31750"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

describe("Word's line shapes", () => {
  it("are read as lines: width, thickness, colour and place", () => {
    const line = lineShape(LINE)!;
    expect([Math.round(line.width), line.height, line.thickness, line.color, Math.round(line.x * 10) / 10, Math.round(line.y * 10) / 10]).toEqual([497, 0, 2.5, "#000000", -21.8, 8]);
  });
  it("other shapes and pictures aren't", () => {
    expect(lineShape(LINE.replace('prst="line"', 'prst="rect"'))).toBeNull();
    expect(lineShape(LINE.replace(/<a:ln [\s\S]*<\/a:ln>/, "<a:ln><a:noFill/></a:ln>"))).toBeNull();
    expect(lineShape("<w:r><w:drawing><pic:pic/></w:drawing></w:r>")).toBeNull();
  });
});

// As Word 365 writes them (Insert → Text Box, Insert → Shapes → Oval), trimmed.
const anchor = (name: string, cx: number, cy: number, sp: string) =>
  `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor behindDoc="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>370205</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>-180975</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:wrapNone/><wp:docPr id="1" name="${name}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>${sp}</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent><w:p><w:r><w:t>Fallback copy</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;
const TEXT_BOX = anchor("Text Box 1", 1905000, 762000, `<wps:cNvSpPr txBox="1"/><wps:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="6350"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Text in a box</w:t></w:r></w:p><w:p><w:r><w:t>R &amp; C</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/>`);
const OVAL = anchor("Oval 3", 762000, 762000, `<wps:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr><wps:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style><wps:bodyPr anchor="ctr"/>`);

describe("Word's shapes and text boxes", () => {
  it("a text box: its size, place, outline and lines of text (not the fallback copy)", () => {
    const box = boxShape(TEXT_BOX)!;
    expect(box).toMatchObject({ kind: "rect", width: 150, height: 60, fill: null, stroke: "#000000", thickness: 0.5, text: ["Text in a box", "R & C"], front: true, middle: false });
    expect([Math.round(box.x * 10) / 10, Math.round(box.y * 10) / 10]).toEqual([29.2, -14.2]);
  });
  it("an oval takes the shape style's colours, with its text in the middle", () => {
    expect(boxShape(OVAL)).toMatchObject({ kind: "ellipse", width: 60, height: 60, fill: "#4472C4", stroke: "#172C51", text: [], middle: true });
  });
  it("is drawn, not shown as a chip", () => {
    const node = schema.nodes.raw_inline.create({ xml: TEXT_BOX, label: "Text box: Text in a box" });
    const [tag, attrs] = schema.nodes.raw_inline.spec.toDOM!(node) as [string, Record<string, string>];
    expect([tag, attrs.class]).toEqual(["span", "shape-box front"]);
    expect(attrs.style).toContain("width:150pt;height:60pt");
  });
  it("shapes it can't draw, and the line, aren't boxes", () => {
    expect(boxShape(OVAL.replace('prst="ellipse"', 'prst="cloud"'))).toBeNull();
    expect(boxShape(LINE)).toBeNull();
  });
  it("Word's horizontal line is a line across", () => {
    expect(horizontalLine(`<w:r><w:pict><v:rect id="_x0000_i1025" style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" fillcolor="#a0a0a0" stroked="f"/></w:pict></w:r>`)).toEqual({ thickness: 1.5, color: "#a0a0a0" });
    expect(horizontalLine(`<w:r><w:pict><v:rect style="width:20pt;height:10pt"/></w:pict></w:r>`)).toBeNull();
  });
});

describe("symbols and drop caps", () => {
  it("Symbol and Wingdings characters are the characters they look like", () => {
    expect([[0x57, "Ω"], [0x6d, "μ"], [0x70, "π"], [0x44, "Δ"], [0xb1, "±"], [0xa3, "≤"], [0xb0, "°"]].map(([c]) => symbolChar("Symbol", c as number))).toEqual(["Ω", "μ", "π", "Δ", "±", "≤", "°"]);
    expect([symbolChar("Wingdings", 0xfc), symbolChar("Calibri", 0x41)]).toEqual(["✓", "A"]);
  });
  it("a symbol is shown as text, not a chip", () => {
    const node = schema.nodes.raw_inline.create({ xml: '<w:r><w:sym w:font="Symbol" w:char="F057"/></w:r>', label: "Ω" });
    const [tag, attrs, text] = schema.nodes.raw_inline.spec.toDOM!(node) as [string, Record<string, string>, string];
    expect([tag, attrs.class, text]).toEqual(["span", "sym-char", "Ω"]);
  });
  it("a drop cap's paragraph floats beside the next", () => {
    const p = schema.nodes.paragraph.create({ extra: '<w:framePr w:dropCap="drop" w:lines="3" w:wrap="around" w:vAnchor="text" w:hAnchor="text"/>' }, schema.text("D"));
    expect((schema.nodes.paragraph.spec.toDOM!(p) as [string, Record<string, string>])[1].class).toBe("para drop-cap");
  });
});

describe("arrows, stars and SmartArt", () => {
  it("an arrow is drawn as its outline, with its text over it", () => {
    const arrow = OVAL.replace('prst="ellipse"', 'prst="rightArrow"');
    expect(boxShape(arrow)?.kind).toBe("rightArrow");
    const dom = schema.nodes.raw_inline.spec.toDOM!(schema.nodes.raw_inline.create({ xml: arrow })) as unknown as [string, Record<string, string>, [string, Record<string, string>, [string, Record<string, string>]], unknown];
    expect([dom[0], dom[1].class, dom[2][0], dom[2][2][1].d.startsWith("M0 15L")]).toEqual(["span", "shape-box front", "http://www.w3.org/2000/svg svg", true]);
  });

  it("outlines for the common presets", () => {
    for (const kind of ["triangle", "diamond", "hexagon", "star5", "chevron", "homePlate", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "plus", "flowChartDecision", "flowChartTerminator"]) expect(presetPath(kind, 100, 50), kind).toMatch(/^M[\d. ]+/);
    expect(presetPath("cloud", 100, 50)).toBeNull();
  });

  it("a SmartArt diagram's shapes are read from the drawing Word keeps with it", () => {
    const drawing = `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree>${[0, 1]
      .map((i) => `<dsp:sp modelId="{${i}}"><dsp:spPr><a:xfrm><a:off x="${i * 1905000}" y="0"/><a:ext cx="1270000" cy="635000"/></a:xfrm><a:prstGeom prst="${i ? "chevron" : "rect"}"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln></dsp:spPr><dsp:style><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style><dsp:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="nb-NO" sz="1800"/><a:t>Steg ${i + 1}</a:t></a:r></a:p></dsp:txBody></dsp:sp>`)
      .join("")}</dsp:spTree></dsp:drawing>`;
    const shapes = readDiagram(drawing);
    expect(shapes.map((s) => [s.kind, s.x, s.w, s.h, s.fill, s.stroke, s.text, s.color, s.size])).toEqual([
      ["rect", 0, 100, 50, "#4472C4", "#FFFFFF", ["Steg 1"], "#FFFFFF", 18],
      ["chevron", 150, 100, 50, "#4472C4", "#FFFFFF", ["Steg 2"], "#FFFFFF", 18],
    ]);
    const box = drawShapes(shapes, 250, 50);
    expect([box.className, box.querySelectorAll("path").length, box.textContent]).toEqual(["shape-group", 2, "Steg 1Steg 2"]);
  });
});

describe("Word's colour adjustments", () => {
  it("shade, tint and lighter/darker are applied to theme colours", () => {
    expect(colorIn('<a:schemeClr val="accent1"/>')).toBe("#4472C4");
    expect(colorIn('<a:schemeClr val="accent1"><a:shade val="15000"/></a:schemeClr>')).toBe("#172C51"); // the dark blue outline of Word's own shapes
    expect(colorIn('<a:srgbClr val="000000"><a:tint val="50000"/></a:srgbClr>')).toBe("#BCBCBC");
    // "Text 1, lighter 50%" (black) and "Accent 1, darker 50%".
    expect(colorIn('<a:schemeClr val="tx1"><a:lumMod val="50000"/><a:lumOff val="50000"/></a:schemeClr>')).toBe("#808080");
    expect(colorIn('<a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr>')).toBe("#203864");
  });
});
