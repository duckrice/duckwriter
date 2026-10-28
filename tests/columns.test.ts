import { describe, expect, it } from "vitest";
import { columnsOf } from "../src/editor/pagination";
import { withColumns } from "../src/editor/commands";

describe("text columns", () => {
  it("are read from the section's properties: equal, with Word's default space, or each its own", () => {
    expect(columnsOf('<w:sectPr><w:cols w:num="2" w:space="720"/></w:sectPr>', 468)).toEqual({ widths: [216, 216], gaps: [36], line: false });
    expect(columnsOf('<w:sectPr><w:cols w:num="3" w:sep="1"/></w:sectPr>', 468)?.gaps).toEqual([36, 36]);
    expect(columnsOf('<w:sectPr><w:cols w:num="3" w:sep="1"/></w:sectPr>', 468)?.line).toBe(true);
    expect(columnsOf('<w:cols w:equalWidth="0"><w:col w:w="2880" w:space="720"/><w:col w:w="5040"/></w:cols>', 468)).toEqual({ widths: [144, 252], gaps: [36], line: false });
    expect(columnsOf('<w:sectPr><w:cols w:space="720"/></w:sectPr>', 468)).toBeNull();
    expect(columnsOf("<w:sectPr/>", 468)).toBeNull();
  });

  it("are set in the section's properties, in Word's order", () => {
    expect(withColumns('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>', 2)).toBe('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:num="2" w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>');
    expect(withColumns('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:docGrid w:linePitch="360"/></w:sectPr>', 3)).toBe('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:num="3" w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>');
    expect(withColumns("<w:sectPr/>", 1)).toBe('<w:sectPr><w:cols w:space="708"/></w:sectPr>');
  });
});

import { withPageNumbering } from "../src/editor/commands";

describe("page number format", () => {
  const base = '<w:sectPr><w:footerReference w:type="default" r:id="rId9"/><w:pgSz w:w="11906" w:h="16838"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>';
  it("writes roman numerals, a restart and a blank first page in Word's order", () => {
    expect(withPageNumbering(base, { format: "lowerRoman", start: 1, firstBlank: true })).toBe('<w:sectPr><w:footerReference w:type="default" r:id="rId9"/><w:pgSz w:w="11906" w:h="16838"/><w:pgNumType w:fmt="lowerRoman" w:start="1"/><w:cols w:space="708"/><w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr>');
  });
  it("back to plain numbers going on: nothing extra written", () => {
    const set = withPageNumbering(base, { format: "upperRoman", start: 3, firstBlank: true });
    expect(withPageNumbering(set, { format: "decimal", start: null, firstBlank: false })).toBe(base);
  });
  it("the last section's title page is the document's own setting, so it isn't written there", () => {
    expect(withPageNumbering(base, { format: "decimal", start: 1, firstBlank: true }, false)).not.toContain("titlePg");
  });
});
