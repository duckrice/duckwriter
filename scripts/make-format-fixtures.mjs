// Makes tests/fixtures/formats.{doc,rtf,odt}: one Word document with what the
// readers must understand (headings, bold, italic, sizes, raised and lowered
// text, a list, a table, a footnote, Norwegian letters), saved by LibreOffice
// in the older formats. Needs LibreOffice; the files are kept in the
// repository, so the tests don't.
//   node scripts/make-format-fixtures.mjs
import { execFileSync } from "node:child_process";
import { copyFileSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import * as docx from "docx";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const soffice = process.env.SOFFICE ?? "C:\\Program Files\\LibreOffice\\program\\soffice.exe";
const work = mkdtempSync(join(tmpdir(), "duckwriter-fixtures-"));
const { Paragraph, TextRun, HeadingLevel } = docx;

const file = new docx.Document({
  footnotes: { 1: { children: [new Paragraph("Measured at noon.")] } },
  numbering: { config: [{ reference: "points", levels: [{ level: 0, format: docx.LevelFormat.BULLET, text: "•", alignment: docx.AlignmentType.LEFT }] }] },
  sections: [
    {
      children: [
        new Paragraph({ text: "Rates of reaction", heading: HeadingLevel.TITLE }),
        new Paragraph({ text: "Introduction", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          children: [
            new TextRun("Plain, "),
            new TextRun({ text: "bold", bold: true }),
            new TextRun(", "),
            new TextRun({ text: "italic", italics: true }),
            new TextRun(", "),
            new TextRun({ text: "underlined", underline: {} }),
            new TextRun(", "),
            new TextRun({ text: "big", size: 32 }),
            new TextRun(", H"),
            new TextRun({ text: "2", subScript: true }),
            new TextRun("O and x"),
            new TextRun({ text: "2", superScript: true }),
            new TextRun(". Blåbærsyltetøy."),
            new docx.FootnoteReferenceRun(1),
          ],
        }),
        new Paragraph({ text: "Methods", heading: HeadingLevel.HEADING_2 }),
        new Paragraph({ text: "First point", numbering: { reference: "points", level: 0 } }),
        new Paragraph({ text: "Second point", numbering: { reference: "points", level: 0 } }),
        new Paragraph({ text: "Centred line", alignment: docx.AlignmentType.CENTER }),
        new docx.Table({
          rows: [
            new docx.TableRow({ children: ["Run", "Rate"].map((t) => new docx.TableCell({ children: [new Paragraph(t)] })) }),
            new docx.TableRow({ children: ["1", "0.5"].map((t) => new docx.TableCell({ children: [new Paragraph(t)] })) }),
          ],
        }),
        new Paragraph("The end."),
      ],
    },
  ],
});
const source = join(work, "formats.docx");
writeFileSync(source, await docx.Packer.toBuffer(file));

const out = join(root, "tests", "fixtures");
mkdirSync(out, { recursive: true });
// A profile of its own: the user's LibreOffice settings stay as they are.
const profile = pathToFileURL(join(work, "lo-profile")).href;
for (const [ext, filter] of [["doc", "doc:MS Word 97"], ["rtf", "rtf"], ["odt", "odt"]]) {
  execFileSync(soffice, [`-env:UserInstallation=${profile}`, "--headless", "--convert-to", filter, "--outdir", work, source], { stdio: "ignore", timeout: 120_000 });
  copyFileSync(join(work, `formats.${ext}`), join(out, `formats.${ext}`));
  console.log(`tests/fixtures/formats.${ext}`);
}
rmSync(work, { recursive: true, force: true });
