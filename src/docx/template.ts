// A new, empty Word document with a sensible set of styles, built in code so
// that new documents go through exactly the same reading and saving as
// documents someone sends you.

import JSZip from "jszip";

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';
const HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

const CONTENT_TYPES = `${HEADER}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Default Extension="jpeg" ContentType="image/jpeg"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

const ROOT_RELS = `${HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;

const DOCUMENT_RELS = `${HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>`;

/** A4, 2.54 cm margins. */
const DOCUMENT = `${HEADER}<w:document ${W} ${R}><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>`;

const SETTINGS = `${HEADER}<w:settings ${W}><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`;

const font = (name: string) => `<w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:eastAsia="${name}" w:cs="${name}"/>`;

const paraStyle = (id: string, name: string, priority: number, pPr: string, rPr: string, extra = "", next = "Normal") =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="${next}"/><w:uiPriority w:val="${priority}"/><w:qFormat/>${extra}<w:pPr>${pPr}</w:pPr><w:rPr>${rPr}</w:rPr></w:style>`;

const heading = (level: number, size: number, color: string, before: number) =>
  paraStyle(
    `Heading${level}`,
    `heading ${level}`,
    9,
    `<w:keepNext/><w:keepLines/><w:spacing w:before="${before}" w:after="80"/><w:outlineLvl w:val="${level - 1}"/>`,
    `${font("Calibri Light")}<w:color w:val="${color}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/>`,
  );

const STYLES = `${HEADER}<w:styles ${W}>
<w:docDefaults>
<w:rPrDefault><w:rPr>${font("Calibri")}<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-GB" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
${paraStyle("Title", "Title", 10, `<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>`, `${font("Calibri Light")}<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>`)}
${paraStyle("Subtitle", "Subtitle", 11, `<w:spacing w:after="160"/>`, `<w:color w:val="595959"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/>`)}
${heading(1, 32, "2F5496", 360)}
${heading(2, 28, "2F5496", 160)}
${heading(3, 24, "1F3763", 160)}
${paraStyle("Quote", "Quote", 29, `<w:spacing w:before="200"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>`, `<w:i/><w:iCs/><w:color w:val="404040"/>`)}
${paraStyle("ListParagraph", "List Paragraph", 34, `<w:ind w:left="720"/><w:contextualSpacing/>`, "", "", "ListParagraph")}
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
</w:styles>`;

const APP = `${HEADER}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>DuckWriter</Application></Properties>`;

function core(now: string) {
  return `${HEADER}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
}

export async function blankDocx(): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", CONTENT_TYPES);
  zip.file("_rels/.rels", ROOT_RELS);
  zip.file("word/document.xml", DOCUMENT);
  zip.file("word/_rels/document.xml.rels", DOCUMENT_RELS);
  zip.file("word/styles.xml", STYLES);
  zip.file("word/settings.xml", SETTINGS);
  zip.file("docProps/app.xml", APP);
  zip.file("docProps/core.xml", core(new Date().toISOString().replace(/\.\d+Z$/, "Z")));
  return zip.generateAsync({ type: "uint8array" });
}
