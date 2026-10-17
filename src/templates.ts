// Ready-made starts for common documents: a scientific article, a report, a
// letter and a thesis chapter. They use the document's own styles (Title,
// Heading 1…), so they look right in Word too.

import type { Node } from "prosemirror-model";
import { ensureStyles } from "./docx/builtins";
import type { DocContext } from "./docx/model";
import { schema } from "./editor/schema";

export type TemplateKind = "article" | "report" | "letter" | "thesis";

export const TEMPLATES: { kind: TemplateKind; name: string; file: string }[] = [
  { kind: "article", name: "Scientific article", file: "Article.docx" },
  { kind: "report", name: "Report", file: "Report.docx" },
  { kind: "letter", name: "Letter", file: "Letter.docx" },
  { kind: "thesis", name: "Thesis chapter", file: "Chapter.docx" },
];

export function templateDoc(kind: TemplateKind, ctx: DocContext): Node {
  ensureStyles(ctx, "Caption", "Bibliography", "TOCHeading", "TOC1", "TOC2", "TOC3");
  const has = (id: string) => ctx.styles.byId.has(id);
  const p = (text: string, style: string | null = null, attrs: Record<string, unknown> = {}) => schema.nodes.paragraph.create({ style: style && has(style) ? style : null, ...attrs }, text ? schema.text(text) : null);
  const italic = (text: string) => schema.nodes.paragraph.create(null, schema.text(text, [schema.marks.italic.create()]));
  const h1 = (text: string) => p(text, "Heading1");
  const h2 = (text: string) => p(text, "Heading2");
  const today = new Date().toLocaleDateString(undefined, { day: "numeric", month: "long", year: "numeric" });
  let blocks: Node[];
  if (kind === "article") {
    blocks = [
      p("Title of the article", "Title"),
      p("First Author¹, Second Author²", "Subtitle"),
      italic("¹ Department, University, City, Country. ² Institute, City, Country. Corresponding author: name@example.org"),
      h1("Abstract"),
      p("Summarise the background, aim, methods, main results and conclusion in about 200 words."),
      p("Keywords: first keyword; second keyword; third keyword"),
      h1("Introduction"),
      p("What is known, what is not, and what this study does. Cite sources with Alt+Shift+C."),
      h1("Methods"),
      h2("Study design"),
      p("Describe the design, the material and the measurements."),
      h2("Statistical analysis"),
      p("Describe the analysis. Equations: Alt+= in the text, Alt+Shift+E numbered."),
      h1("Results"),
      p("Present the findings, with figures and tables (captions: Ctrl+Alt+C)."),
      h1("Discussion"),
      p("Interpret the results, compare them with earlier work, and name the limitations."),
      h1("Conclusions"),
      p("The main conclusions, briefly."),
      h1("Acknowledgements"),
      p("Funding and help."),
      schema.nodes.bibliography.create({ title: "References" }),
      p(""),
    ];
  } else if (kind === "report") {
    blocks = [
      p("Report title", "Title"),
      p(`Author · ${today}`, "Subtitle"),
      schema.nodes.toc.create({ title: "Contents" }),
      h1("Summary"),
      p("The main points of the report, for readers who read nothing else."),
      h1("Background"),
      p("Why this report was written."),
      h1("Findings"),
      h2("First finding"),
      p("Details."),
      h2("Second finding"),
      p("Details."),
      h1("Recommendations"),
      p("What should be done."),
      schema.nodes.bibliography.create({ title: "References" }),
      p(""),
    ];
  } else if (kind === "letter") {
    blocks = [
      p("Your Name", null, { align: "right" }),
      p("Street 1", null, { align: "right" }),
      p("0000 City", null, { align: "right" }),
      p(""),
      p(today, null, { align: "right" }),
      p(""),
      p("Recipient Name"),
      p("Organisation"),
      p("Street 1"),
      p("0000 City"),
      p(""),
      p("Subject of the letter", "Heading2"),
      p("Dear Recipient,"),
      p("Write the letter here."),
      p("Yours sincerely,"),
      p(""),
      p("Your Name"),
    ];
  } else {
    blocks = [
      p("1 Chapter title", "Heading1"),
      italic("A short introduction to the chapter: what it covers and how it is organised."),
      h2("1.1 Background"),
      p("Text."),
      h2("1.2 Aims"),
      p("Text."),
      h2("1.3 Outline"),
      p("Text."),
      schema.nodes.toc.create({ title: "List of figures", label: "Figure" }),
      p(""),
    ];
  }
  return schema.nodes.doc.create(null, blocks);
}
