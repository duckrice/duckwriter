import JSZip from "jszip";
import type { Node } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { schema } from "../src/editor/schema";

// A column chart as Word writes it: two series over three categories, the numbers cached in the chart part.
const ser = (i: number, name: string, values: number[]) =>
  `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:strRef><c:f>Sheet1!$${"BC"[i]}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${name}</c:v></c:pt></c:strCache></c:strRef></c:tx>` +
  `<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>10 °C</c:v></c:pt><c:pt idx="1"><c:v>20 °C</c:v></c:pt><c:pt idx="2"><c:v>30 °C</c:v></c:pt></c:strCache></c:strRef></c:cat>` +
  `<c:val><c:numRef><c:f>Sheet1!$${"BC"[i]}$2:$${"BC"[i]}$4</c:f><c:numCache><c:formatCode>0.0</c:formatCode><c:ptCount val="3"/>${values.map((v, k) => `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:val></c:ser>`;
const CHART = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Resistance</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>${ser(0, "Copper", [1.5, 2.25, 3])}${ser(1, "Steel", [4, 5, 6])}</c:barChart><c:catAx><c:axPos val="b"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Temperature</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx><c:valAx><c:axPos val="l"/></c:valAx></c:plotArea></c:chart></c:chartSpace>`;
const DRAWING = `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="5486400" cy="3200400"/><wp:docPr id="1" name="Chart 1"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdChart1"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

async function withChart(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await blankDocx());
  const doc = await zip.file("word/document.xml")!.async("string");
  zip.file("word/document.xml", doc.replace(/<w:body>/, `<w:body>${DRAWING}`));
  const rels = await zip.file("word/_rels/document.xml.rels")!.async("string");
  zip.file("word/_rels/document.xml.rels", rels.replace("</Relationships>", '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>'));
  zip.file("word/charts/chart1.xml", CHART);
  return zip.generateAsync({ type: "uint8array" });
}

const chartIn = (doc: Node) => {
  let found: Node | null = null;
  doc.descendants((n) => {
    if (n.type.name === "raw_inline" && n.attrs.chart) found = n;
  });
  return found as Node | null;
};

describe("Word's charts", () => {
  it("are read from the numbers Word keeps with them", async () => {
    const chart = chartIn((await readDocx(await withChart())).doc)!;
    expect(chart.attrs.chart).toMatchObject({
      kind: "column",
      title: "Resistance",
      xTitle: "Temperature",
      categories: ["10 °C", "20 °C", "30 °C"],
      series: [
        { name: "Copper", values: [1.5, 2.25, 3] },
        { name: "Steel", values: [4, 5, 6] },
      ],
      decimals: 1,
    });
  });

  it("are drawn at their size, a bar for each number", async () => {
    const chart = chartIn((await readDocx(await withChart())).doc)!;
    const box = schema.nodes.raw_inline.spec.toDOM!(chart) as HTMLElement;
    expect([box.className, box.style.width, box.style.height]).toEqual(["chart-box", "432pt", "252pt"]);
    const bars = box.querySelectorAll("rect[fill]").length - 1 - 2; // less the background and the legend's two keys
    expect(bars).toBe(6);
    expect(box.textContent).toContain("Resistance");
    expect(box.textContent).toContain("6.0"); // the axis, with the numbers' one decimal
  });

  it("stay in the file as they were when saved", async () => {
    const original = await withChart();
    const { doc, context } = await readDocx(original);
    const saved = await JSZip.loadAsync(await writeDocx(doc, context));
    expect(await saved.file("word/charts/chart1.xml")!.async("string")).toBe(CHART);
    expect(await saved.file("word/document.xml")!.async("string")).toContain('r:id="rIdChart1"');
  });
});

it("a chart copied and pasted keeps its numbers", async () => {
  const chart = chartIn((await readDocx(await withChart())).doc)!;
  const box = schema.nodes.raw_inline.spec.toDOM!(chart) as HTMLElement;
  expect([box.getAttribute("data-pm-node"), JSON.parse(box.getAttribute("data-pm")!).chart.title]).toEqual(["raw_inline", "Resistance"]);
});
