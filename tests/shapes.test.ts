import { describe, expect, it } from "vitest";
import { lineShape } from "../src/editor/schema";

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
