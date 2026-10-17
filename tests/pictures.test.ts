import * as docx from "docx";
import JSZip from "jszip";
import { readFileSync } from "node:fs";
import type { Node } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { schema } from "../src/editor/schema";

const png = readFileSync("src-tauri/icons/32x32.png");

async function withPictures(): Promise<Uint8Array> {
  const d = new docx.Document({
    sections: [
      {
        children: [
          new docx.Paragraph({ children: [new docx.TextRun("Before "), new docx.ImageRun({ data: png, type: "png", transformation: { width: 64, height: 48 } }), new docx.TextRun(" after")] }),
          new docx.Paragraph({
            children: [
              new docx.ImageRun({
                data: png,
                type: "png",
                transformation: { width: 40, height: 40 },
                floating: {
                  horizontalPosition: { relative: docx.HorizontalPositionRelativeFrom.MARGIN, offset: 914400 },
                  verticalPosition: { relative: docx.VerticalPositionRelativeFrom.PARAGRAPH, offset: 457200 },
                  wrap: { type: docx.TextWrappingType.SQUARE },
                },
              }),
              new docx.TextRun("Text that flows around the picture."),
            ],
          }),
        ],
      },
    ],
  });
  return new Uint8Array(await docx.Packer.toBuffer(d));
}

const images = (doc: Node) => {
  const out: Node[] = [];
  doc.descendants((n) => {
    if (n.type === schema.nodes.image) out.push(n);
  });
  return out;
};

const replaceAttrs = (doc: Node, index: number, attrs: Record<string, unknown>): Node => {
  let result = doc;
  let i = 0;
  doc.descendants((n, pos) => {
    if (n.type !== schema.nodes.image) return;
    if (i++ === index) result = doc.replace(pos, pos + 1, new (doc.slice(0, 0).constructor as typeof import("prosemirror-model").Slice)(schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, n.type.create({ ...n.attrs, ...attrs }, null, n.marks))).content.firstChild!.content, 0, 0));
  });
  return result;
};

describe("pictures", () => {
  it("reads pictures in line with the text and floating ones, with size, position and wrapping", async () => {
    const { doc } = await readDocx(await withPictures());
    const [inline, floating] = images(doc);
    expect(inline.attrs).toMatchObject({ floating: false, width: 48, height: 36 });
    expect(floating.attrs).toMatchObject({ floating: true, width: 30, height: 30, x: 72, y: 36, relX: "margin", relY: "paragraph", wrap: "square" });
    expect(inline.attrs.rid).toMatch(/^rId/);
  });

  it("saves pictures that weren't changed exactly as they were", async () => {
    const original = await withPictures();
    const { doc, context } = await readDocx(original);
    const before = await (await JSZip.loadAsync(original)).file("word/document.xml")!.async("string");
    const after = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/document.xml")!.async("string");
    const drawings = (xml: string) => xml.match(/<w:drawing>[\s\S]*?<\/w:drawing>/g);
    expect(drawings(after)).toEqual(drawings(before));
  });

  it("saves a resized and a moved picture by adjusting its own XML", async () => {
    const { doc, context } = await readDocx(await withPictures());
    let changed = replaceAttrs(doc, 0, { width: 96, height: 72 });
    changed = replaceAttrs(changed, 1, { x: 144, y: 18, wrap: "topAndBottom" });
    const saved = await writeDocx(changed, context);
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("string");
    expect(xml).toContain('<wp:extent cx="1219200" cy="914400"/>');
    expect(xml).toContain('<wp:positionH relativeFrom="margin"><wp:posOffset>1828800</wp:posOffset></wp:positionH>');
    expect(xml).toContain("<wp:wrapTopAndBottom/>");
    expect(xml).toContain('cstate="none"'); // Its original details are still there.
    const [inline, floating] = images((await readDocx(saved)).doc);
    expect(inline.attrs).toMatchObject({ width: 96, height: 72 });
    expect(floating.attrs).toMatchObject({ x: 144, y: 18, wrap: "topAndBottom" });
  });

  it("adds a new picture to the file's media, with its type and relationship", async () => {
    const { context } = await readDocx(await blankDocx());
    context.media.set("p1", { bytes: new Uint8Array(png), ext: "png", rid: null });
    const doc = schema.nodes.doc.create(null, schema.nodes.paragraph.create(null, [schema.text("A picture: "), schema.nodes.image.create({ media: "p1", width: 120, height: 90 })]));
    const saved = await writeDocx(doc, context);
    const zip = await JSZip.loadAsync(saved);
    expect(zip.file("word/media/duckwriter1.png")).not.toBeNull();
    expect(await zip.file("[Content_Types].xml")!.async("string")).toContain('<Default Extension="png" ContentType="image/png"/>');
    expect(await zip.file("word/_rels/document.xml.rels")!.async("string")).toContain('Target="media/duckwriter1.png"');
    const again = await readDocx(saved);
    expect(images(again.doc)[0].attrs).toMatchObject({ width: 120, height: 90, floating: false });
    // Saving again doesn't store the picture twice.
    const twice = await JSZip.loadAsync(await writeDocx(doc, context));
    expect(Object.keys(twice.files).filter((f) => f.startsWith("word/media/") && !f.endsWith("/"))).toEqual(["word/media/duckwriter1.png"]);
  });

  it("turns a picture in line with the text into a floating one", async () => {
    const { doc, context } = await readDocx(await withPictures());
    const changed = replaceAttrs(doc, 0, { floating: true, x: 20, y: 10, relX: "column", relY: "paragraph", wrap: "square" });
    const saved = await writeDocx(changed, context);
    const [first] = images((await readDocx(saved)).doc);
    expect(first.attrs).toMatchObject({ floating: true, x: 20, y: 10, wrap: "square", width: 48, height: 36 });
  });
});
