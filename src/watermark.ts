// Watermarks, as Word makes them: big see-through words turned across the
// middle of every page. Word keeps one as a WordArt shape (VML) in the page
// header, named PowerPlusWaterMarkObject; DuckWriter writes it the same way,
// so Word shows and edits it, and finds those Word made.

import { Fragment, type Node } from "prosemirror-model";
import type { DocContext } from "./docx/model";
import { schema } from "./editor/schema";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const isMark = (n: Node) => n.type === schema.nodes.raw_inline && /PowerPlusWaterMarkObject/.test(n.attrs.xml ?? "");

/** The run holding the watermark shape, as Word writes it. */
export function watermarkXml(text: string): string {
  // Word sizes the shape to the page's text area; this is its size for one line of words.
  const width = Math.min(468, Math.max(160, text.length * 52));
  const formulas = ["sum #0 0 10800", "prod #0 2 1", "sum 21600 0 @1", "sum 0 0 @2", "sum 21600 0 @3", "if @0 @3 0", "if @0 21600 @1", "if @0 0 @2", "if @0 @4 21600", "mid @5 @6", "mid @8 @5", "mid @7 @8", "mid @6 @7", "sum @6 0 @5"].map((f) => `<v:f eqn="${f}"/>`).join("");
  return (
    `<w:r><w:rPr><w:noProof/></w:rPr><w:pict xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">` +
    `<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800" path="m@7,l@8,m@5,21600l@6,21600e"><v:formulas>${formulas}</v:formulas><v:path textpathok="t" o:connecttype="custom" o:connectlocs="@9,0;@10,10800;@11,21600;@12,10800" o:connectangles="270,180,90,0"/><v:textpath on="t" fitshape="t"/><v:handles><v:h position="#0,bottomRight" xrange="6629,14971"/></v:handles><o:lock v:ext="edit" text="t" shapetype="t"/></v:shapetype>` +
    `<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:${width}pt;height:${Math.round(width / 3)}pt;rotation:315;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="${esc(text)}"/></v:shape>` +
    `</w:pict></w:r>`
  );
}

/** The words of the document's watermark (Word's or DuckWriter's), or null. */
export function watermarkText(ctx: DocContext): string | null {
  for (const story of ctx.stories) {
    if (story.kind !== "header") continue;
    let found: string | null = null;
    story.doc.descendants((n) => {
      if (found === null && isMark(n)) found = /\bstring="([^"]*)"/.exec(n.attrs.xml)?.[1]?.replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&") ?? null;
    });
    if (found !== null) return found;
  }
  return null;
}

/** Takes every watermark out of the headers, and puts `text` in the default one (made if needed). */
export function setWatermark(ctx: DocContext, text: string | null) {
  const strip = (doc: Node) => doc.copy(Fragment.fromArray(doc.content.content.map((p) => (p.isTextblock ? p.copy(Fragment.fromArray(p.content.content.filter((n) => !isMark(n)))) : p))));
  for (const story of ctx.stories) if (story.kind === "header") story.doc = strip(story.doc);
  if (!text) return;
  const mark = schema.nodes.raw_inline.create({ xml: watermarkXml(text), hidden: true });
  let story = ctx.stories.find((s) => s.kind === "header" && s.type === "default");
  if (!story) {
    story = { kind: "header", type: "default", path: null, rid: null, rootStart: null, prefixes: new Set(), rels: new Map(), doc: schema.nodes.doc.create(null, schema.nodes.paragraph.create()), saved: null };
    ctx.stories.push(story);
  }
  const first = story.doc.firstChild;
  if (first?.isTextblock) story.doc = story.doc.copy(story.doc.content.replaceChild(0, first.copy(Fragment.from(mark).append(first.content))));
  else story.doc = story.doc.copy(Fragment.from(schema.nodes.paragraph.create(null, mark)).append(story.doc.content));
}
