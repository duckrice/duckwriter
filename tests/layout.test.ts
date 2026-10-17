import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { stylesheetFor } from "../src/editor/stylesheet";

describe("line grid", () => {
  it("is saved as Word's document grid and read back", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    expect(context.page.lineGrid).toBeNull();
    context.page.lineGrid = 15;
    context.lineGridChanged = true;
    const saved = await writeDocx(doc, context);
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:docGrid w:type="lines" w:linePitch="300"/>');
    expect((await readDocx(saved)).context.page.lineGrid).toBe(15);
  });

  it("rounds every style's line height up to whole grid steps", async () => {
    const { context } = await readDocx(await blankDocx());
    expect(stylesheetFor(context, ".doc")).not.toContain("round(up");
    context.page.lineGrid = 15;
    const css = stylesheetFor(context, ".doc");
    expect(css).toMatch(/\.para\[data-style="Heading1"\] \{[^}]*line-height:round\(up, calc\([\d.]+em\), 15pt\)/);
    expect(css).toMatch(/\.para:not\(\[data-style\]\) \{[^}]*line-height:round\(up, calc\([\d.]+em\), 15pt\)/);
  });

  it("switching it off again leaves a plain grid, as Word writes it", async () => {
    const { doc, context } = await readDocx(await blankDocx());
    context.page.lineGrid = null;
    context.lineGridChanged = true;
    const xml = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:docGrid w:linePitch="360"/>');
  });
});
