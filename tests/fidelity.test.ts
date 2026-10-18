// What Word 365 itself writes, beyond what the docx library makes: paragraph
// ids, content controls, text boxes inside mc:AlternateContent, simple fields,
// smart tags and proofing marks. Opened and saved, it must come back intact
// and pass the package check.

import JSZip from "jszip";
import { Fragment, Slice } from "prosemirror-model";
import { describe, expect, it } from "vitest";
import { checkPackage } from "../src/docx/check";
import { readDocx } from "../src/docx/read";
import { blankDocx } from "../src/docx/template";
import { writeDocx } from "../src/docx/write";
import { schema } from "../src/editor/schema";

const ROOT = `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" mc:Ignorable="w14 w15 wp14">`;

const TEXTBOX = `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="1828800" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Text Box 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>Boxed words</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="Text Box 1" o:spid="_x0000_s1026" type="#_x0000_t202" style="width:144pt;height:36pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Boxed words</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`;

const BODY = `<w:body>
<w:p w14:paraId="1A2B3C4D" w14:textId="77777777" w:rsidR="00AB12CD" w:rsidRDefault="00AB12CD"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r></w:p>
<w:sdt><w:sdtPr><w:id w:val="-1234"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r></w:p><w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr><w:r><w:t>Introduction</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:sdtContent></w:sdt>
<w:p w14:paraId="2B3C4D5E" w14:textId="77777777"><w:r><w:t xml:space="preserve">Written by </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="Author"/><w:tag w:val="author"/><w:id w:val="99"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Kari Nordmann</w:t></w:r></w:sdtContent></w:sdt><w:r><w:t xml:space="preserve"> on </w:t></w:r><w:smartTag w:uri="urn:schemas-microsoft-com:office:smarttags" w:element="date"><w:r><w:t xml:space="preserve">1 May </w:t></w:r></w:smartTag><w:proofErr w:type="spellStart"/><w:r><w:t>speling</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve"> and a box: </w:t></w:r><w:r>${TEXTBOX}</w:r></w:p>
<w:sdt><w:sdtPr><w:id w:val="5"/><w:showingPlcHdr/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Click here to enter text.</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:titlePg/></w:sectPr></w:body>`;

async function wordLike(): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(await blankDocx());
  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${ROOT}${BODY.replace(/\n/g, "")}</w:document>`);
  return zip.generateAsync({ type: "uint8array" });
}

describe("a document as Word 365 writes it", () => {
  it("opens with its text", async () => {
    const { doc } = await readDocx(await wordLike());
    const text = doc.textContent;
    expect(text).toContain("Introduction");
    expect(text).toContain("Written by Kari Nordmann on 1 May speling.");
    expect(text).toContain("and a box");
    // The text box is shown as one, the author box as editable text.
    const pieces: string[] = [];
    doc.descendants((n) => {
      if (n.type === schema.nodes.raw_inline && n.attrs.label) pieces.push(n.attrs.label);
      if (n.isText && n.marks.some((m) => m.type === schema.marks.contentControl)) pieces.push(`control:${n.text}`);
    });
    expect(pieces).toContain("Text box: Boxed words");
    expect(pieces).toContain("control:Kari Nordmann");
  });

  it("is saved unedited with everything in it, and passes the package check", async () => {
    const original = await wordLike();
    expect(await checkPackage(original)).toEqual([]);
    const { doc, context } = await readDocx(original);
    const saved = await writeDocx(doc, context);
    expect(await checkPackage(saved)).toEqual([]);
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("string");
    expect(xml).toContain('mc:Ignorable="w14 w15 wp14"');
    expect(xml).toContain("<mc:AlternateContent>");
    expect(xml).toContain("Boxed words");
    expect(xml).toContain("<w:docPartGallery w:val=\"Table of Contents\"/>");
    expect(xml).toContain("Click here to enter text.");
    expect(xml).toContain("<w:titlePg/>");
    expect(xml).toContain('<w:sdt><w:sdtPr><w:alias w:val="Author"/><w:tag w:val="author"/><w:id w:val="99"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve">Kari Nordmann</w:t></w:r></w:sdtContent></w:sdt>');
  });

  it("keeps it all when the text around it is edited", async () => {
    const { doc, context } = await readDocx(await wordLike());
    // Typing at the end of the first paragraph.
    const end = doc.firstChild!.nodeSize - 1;
    const edited = doc.replace(end, end, new Slice(Fragment.from(schema.text(" again")), 0, 0));
    const saved = await writeDocx(edited, context);
    expect(await checkPackage(saved)).toEqual([]);
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("string");
    expect(xml).toContain("Introduction again");
    expect(xml).toContain("<mc:AlternateContent>");
    expect(xml).toContain("Boxed words");
    expect(xml).toContain("Table of Contents");
    expect(xml).toContain("Click here to enter text.");
    // Word's namespaces are still declared for what is kept.
    for (const prefix of ["mc", "wps", "w14", "v", "o"]) expect(xml).toContain(`xmlns:${prefix}=`);
  });
});

describe("a template's content controls", () => {
  it("lose their placeholder look once typed in", async () => {
    const zip = await JSZip.loadAsync(await blankDocx());
    const control = '<w:sdt><w:sdtPr><w:alias w:val="Title"/><w:id w:val="7"/><w:showingPlcHdr/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>Click to add the title</w:t></w:r></w:sdtContent></w:sdt>';
    zip.file("word/document.xml", (await zip.file("word/document.xml")!.async("string")).replace("<w:body>", `<w:body><w:p>${control}</w:p>`));
    const { doc, context } = await readDocx(await zip.generateAsync({ type: "uint8array" }));
    // Unchanged: still the placeholder.
    let xml = await (await JSZip.loadAsync(await writeDocx(doc, context))).file("word/document.xml")!.async("string");
    expect(xml).toContain("<w:showingPlcHdr/>");
    // The placeholder replaced by a title.
    const para = doc.firstChild!;
    const mark = para.firstChild!.marks.find((m) => m.type === schema.marks.contentControl)!;
    const typed = doc.replace(1, para.nodeSize - 1, new Slice(Fragment.from(schema.text("Ducks in winter", [mark])), 0, 0));
    xml = await (await JSZip.loadAsync(await writeDocx(typed, context))).file("word/document.xml")!.async("string");
    expect(xml).toContain('<w:sdt><w:sdtPr><w:alias w:val="Title"/><w:id w:val="7"/><w:text/></w:sdtPr><w:sdtContent>');
    expect(xml).toContain("Ducks in winter");
  });
});
