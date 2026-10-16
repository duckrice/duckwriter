// Writes the Word documents the app test opens, into samples/. They are made
// with the "docx" library, so they come from a different program than DuckWriter.
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import * as docx from "docx";

const out = join(dirname(fileURLToPath(import.meta.url)), "..", "samples");
mkdirSync(out, { recursive: true });

const { Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;

const report = new docx.Document({
  creator: "Sample",
  styles: { default: { document: { run: { font: "Calibri", size: 22 }, paragraph: { spacing: { after: 120 } } } } },
  numbering: {
    config: [
      {
        reference: "steps",
        levels: [
          { level: 0, format: docx.LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: docx.LevelFormat.LOWER_LETTER, text: "%2)", alignment: AlignmentType.START, style: { paragraph: { indent: { left: 1440, hanging: 360 } } } },
        ],
      },
      {
        reference: "dots",
        levels: [{ level: 0, format: docx.LevelFormat.BULLET, text: "•", alignment: AlignmentType.START, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
    ],
  },
  sections: [
    {
      children: [
        new Paragraph({ heading: HeadingLevel.TITLE, children: [new TextRun("Club trip to Bergen")] }),
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Plan")] }),
        new Paragraph({
          alignment: AlignmentType.JUSTIFIED,
          children: [
            new TextRun("We leave from Oslo on "),
            new TextRun({ text: "Friday 7 June", bold: true }),
            new TextRun(" by train and come back on Sunday evening. The trip costs "),
            new TextRun({ text: "1 450 kr", color: "C00000", bold: true }),
            new TextRun(" per person, which covers the train, two nights in the hostel and the "),
            new TextRun({ text: "Fløibanen", italics: true }),
            new TextRun(" funicular."),
          ],
        }),
        new Paragraph({ numbering: { reference: "steps", level: 0 }, children: [new TextRun("Sign up by 15 May")] }),
        new Paragraph({ numbering: { reference: "steps", level: 1 }, children: [new TextRun("Use the form on the website")] }),
        new Paragraph({ numbering: { reference: "steps", level: 1 }, children: [new TextRun("Or ask Kari at training")] }),
        new Paragraph({ numbering: { reference: "steps", level: 0 }, children: [new TextRun("Pay the deposit")] }),
        new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("What to bring")] }),
        new Paragraph({ numbering: { reference: "dots", level: 0 }, children: [new TextRun("Rain jacket — it is Bergen")] }),
        new Paragraph({ numbering: { reference: "dots", level: 0 }, children: [new TextRun({ text: "Good shoes", highlight: "yellow" })] }),
        new Paragraph({ numbering: { reference: "dots", level: 0 }, children: [new TextRun({ text: "Packed lunch", font: "Georgia", size: 28 })] }),
        new docx.Table({
          rows: [
            new docx.TableRow({ children: ["Day", "Activity"].map((t) => new docx.TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t, bold: true })] })] })) }),
            new docx.TableRow({ children: ["Friday", "Train, check in"].map((t) => new docx.TableCell({ children: [new Paragraph(t)] })) }),
            new docx.TableRow({ children: ["Saturday", "Fløyen and Bryggen"].map((t) => new docx.TableCell({ children: [new Paragraph(t)] })) }),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun("More at "),
            new docx.ExternalHyperlink({ link: "https://example.com/bergen", children: [new TextRun({ text: "the club website", style: "Hyperlink" })] }),
            new TextRun("."),
          ],
        }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "See you there!", size: 32, color: "0F766E" })] }),
      ],
    },
  ],
});

writeFileSync(join(out, "report.docx"), await docx.Packer.toBuffer(report));
writeFileSync(join(out, "not-a-docx.docx"), "This is plain text, not a Word document.");
console.log(`Wrote samples to ${out}`);
