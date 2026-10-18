// Word's standard styles for the things DuckWriter can add (footnotes,
// captions, a table of contents, headers and footers). A document that doesn't
// have one gets it when the feature is first used, as Word does.

import type { DocContext } from "./model";
import { styleFromXml } from "./read";

const para = (id: string, name: string, pPr: string, rPr: string, extra = "") =>
  `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>${extra}<w:pPr>${pPr}</w:pPr><w:rPr>${rPr}</w:rPr></w:style>`;

const BUILTINS: Record<string, string> = {
  FootnoteText: para("FootnoteText", "footnote text", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>', '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'),
  EndnoteText: para("EndnoteText", "endnote text", '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>', '<w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>'),
  EndnoteReference:
    '<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>',
  FootnoteReference:
    '<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>',
  Caption: para("Caption", "caption", '<w:spacing w:after="200" w:line="240" w:lineRule="auto"/>', '<w:i/><w:iCs/><w:color w:val="44546A"/><w:sz w:val="18"/><w:szCs w:val="18"/>', '<w:uiPriority w:val="35"/><w:unhideWhenUsed/><w:qFormat/>'),
  TOCHeading:
    '<w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:next w:val="Normal"/><w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>',
  TOC1: para("TOC1", "toc 1", '<w:spacing w:after="100"/>', "", '<w:uiPriority w:val="39"/><w:unhideWhenUsed/>'),
  TOC2: para("TOC2", "toc 2", '<w:spacing w:after="100"/><w:ind w:left="220"/>', "", '<w:uiPriority w:val="39"/><w:unhideWhenUsed/>'),
  TOC3: para("TOC3", "toc 3", '<w:spacing w:after="100"/><w:ind w:left="440"/>', "", '<w:uiPriority w:val="39"/><w:unhideWhenUsed/>'),
  TableofFigures: para("TableofFigures", "table of figures", '<w:spacing w:after="0"/>', "", '<w:uiPriority w:val="99"/><w:unhideWhenUsed/>'),
  TableGrid:
    '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>',
  Bibliography: para("Bibliography", "Bibliography", '<w:ind w:left="720" w:hanging="720"/>', "", '<w:uiPriority w:val="37"/><w:unhideWhenUsed/>'),
  Hyperlink:
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>',
  CommentText: para("CommentText", "annotation text", '<w:spacing w:line="240" w:lineRule="auto"/>', '<w:sz w:val="20"/><w:szCs w:val="20"/>', '<w:uiPriority w:val="99"/><w:unhideWhenUsed/>'),
  CommentReference:
    '<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:rPr><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr></w:style>',
  Header: para("Header", "header", '<w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', "", '<w:uiPriority w:val="99"/><w:unhideWhenUsed/>'),
  Footer: para("Footer", "footer", '<w:tabs><w:tab w:val="center" w:pos="4513"/><w:tab w:val="right" w:pos="9026"/></w:tabs><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>', "", '<w:uiPriority w:val="99"/><w:unhideWhenUsed/>'),
};

/**
 * Makes sure the document has these styles: missing ones are added to what
 * is shown now and to the file when it is saved.
 */
export function ensureStyles(ctx: DocContext, ...ids: string[]) {
  for (const id of ids) {
    if (ctx.styles.byId.has(id) || !BUILTINS[id]) continue;
    const style = styleFromXml(BUILTINS[id], ctx.themeFonts);
    if (!style) continue;
    ctx.styles.byId.set(id, style);
    ctx.addedStyles.push(BUILTINS[id]);
  }
}
