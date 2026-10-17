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
      headers: { default: new docx.Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Bergen trip 2026", color: "808080" })] })] }) },
      footers: {
        default: new docx.Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Page ", docx.PageNumber.CURRENT, " of ", docx.PageNumber.TOTAL_PAGES] })] })] }),
      },
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

// Several pages of text with headings and a page break, for the page layout.
const lorem =
  "Measurements were repeated three times at each temperature, and the mean value was used. The uncertainty was estimated from the spread of the results and the resolution of the instruments. ";
const long = [];
for (let section = 1; section <= 6; section++) {
  long.push(new Paragraph({ heading: HeadingLevel.HEADING_1, keepNext: true, children: [new TextRun(`${section}. Section ${section}`)] }));
  for (let p = 0; p < 5; p++) long.push(new Paragraph({ children: [new TextRun(lorem.repeat(2 + ((section + p) % 3)))] }));
  if (section === 3) long.push(new Paragraph({ children: [new TextRun("The end of part one."), new docx.PageBreak()] }));
  // Longer than a page: it has to continue on the next one.
  if (section === 5) long.push(new Paragraph({ children: [new TextRun("A very long paragraph. " + lorem.repeat(30))] }));
}
const longDoc = new docx.Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 }, paragraph: { spacing: { after: 160 } } } } },
  sections: [{ children: long }],
});
writeFileSync(join(out, "long.docx"), await docx.Packer.toBuffer(longDoc));

// A table longer than a page.
const tableRows = [new docx.TableRow({ tableHeader: true, children: ["Run", "Temperature (°C)", "Rate (mol/s)"].map((t) => new docx.TableCell({ children: [new Paragraph({ children: [new TextRun({ text: t, bold: true })] })] })) })];
for (let i = 1; i <= 70; i++) {
  tableRows.push(new docx.TableRow({ children: [String(i), String(20 + i), (0.01 * i).toFixed(2)].map((t) => new docx.TableCell({ children: [new Paragraph(t)] })) }));
}
const tableDoc = new docx.Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{ children: [new Paragraph("Measurements:"), new docx.Table({ columnWidths: [1500, 3500, 3500], rows: tableRows }), new Paragraph("End of the measurements.")] }],
});
writeFileSync(join(out, "table-long.docx"), await docx.Packer.toBuffer(tableDoc));
writeFileSync(join(out, "not-a-docx.docx"), "This is plain text, not a Word document.");

// A reviewed document: tracked changes by two people, and comments.
const ada = { author: "Ada Lovelace", date: "2026-10-01T09:00:00Z" };
const max = { author: "Max Planck", date: "2026-10-02T14:30:00Z" };
const reviewDoc = new docx.Document({
  styles: { default: { document: { run: { font: "Calibri", size: 22 }, paragraph: { spacing: { after: 160 } } } } },
  comments: {
    children: [
      { id: 0, author: ada.author, date: new Date(ada.date), initials: "AL", children: [new Paragraph("Is this the right figure?")] },
      { id: 1, author: max.author, date: new Date(max.date), initials: "MP", children: [new Paragraph("Please cite the source."), new Paragraph("Any review paper will do.")] },
    ],
  },
  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun("The rate "),
            new docx.DeletedTextRun({ text: "was", id: 10, ...ada }),
            new docx.InsertedTextRun({ text: "is", id: 11, ...ada }),
            new TextRun(" proportional to the temperature."),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun("As shown in "),
            new docx.CommentRangeStart(0),
            new TextRun("figure 3"),
            new docx.CommentRangeEnd(0),
            new TextRun({ children: [new docx.CommentReference(0)] }),
            new TextRun(", the curve "),
            new docx.InsertedTextRun({ text: "clearly ", id: 12, ...max }),
            new TextRun("flattens out."),
          ],
        }),
        new Paragraph({
          children: [
            new docx.CommentRangeStart(1),
            new TextRun("Earlier work found the same."),
            new docx.CommentRangeEnd(1),
            new TextRun({ children: [new docx.CommentReference(1)] }),
            new docx.DeletedTextRun({ text: " This sentence goes.", id: 13, ...max }),
          ],
        }),
      ],
    },
  ],
});
writeFileSync(join(out, "review.docx"), await docx.Packer.toBuffer(reviewDoc));
console.log(`Wrote samples to ${out}`);
