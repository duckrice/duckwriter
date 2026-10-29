import { describe, expect, it } from "vitest";
import { symbolChar } from "../src/docx/read";
import { boxShape, horizontalLine, lineShape, schema } from "../src/editor/schema";

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
    expect(boxShape(OVAL)).toMatchObject({ kind: "ellipse", width: 60, height: 60, fill: "#4472C4", stroke: "#4472C4", text: [], middle: true });
  });
  it("is drawn, not shown as a chip", () => {
    const node = schema.nodes.raw_inline.create({ xml: TEXT_BOX, label: "Text box: Text in a box" });
    const [tag, attrs] = schema.nodes.raw_inline.spec.toDOM!(node) as [string, Record<string, string>];
    expect([tag, attrs.class]).toEqual(["span", "shape-box front"]);
    expect(attrs.style).toContain("width:150pt;height:60pt");
  });
  it("shapes it can't draw, and the line, aren't boxes", () => {
    expect(boxShape(OVAL.replace('prst="ellipse"', 'prst="rightArrow"'))).toBeNull();
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
