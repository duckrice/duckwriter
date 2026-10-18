// Pictures in Word files (DrawingML): reading their size, position and text
// wrapping, and writing them back, either by adjusting the original XML (so
// cropping, alt text and effects survive) or, for new pictures, from scratch.

import { NS, escapeAttr } from "./xml";

export const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
export const EMU_PER_POINT = 12700;

export type Wrap = "square" | "topAndBottom" | "none" | "behind";

export interface PictureLayout {
  width: number;
  height: number;
  floating: boolean;
  /** For a floating picture: its offset from what it is positioned against, in points. */
  x: number;
  y: number;
  /** What x and y are measured from: "page", "margin", "column", "paragraph"… */
  relX: string;
  relY: string;
  wrap: Wrap;
}

export interface PictureInfo extends PictureLayout {
  rid: string;
  /** Its description for people who can't see it (Word's alt text). */
  alt: string;
}

const emu = (value: string | null | undefined) => Number(value ?? 0) / EMU_PER_POINT;

/** The picture in a <w:drawing>, or null for charts, shapes and the like. */
export function readPicture(drawing: Element): PictureInfo | null {
  const holder = Array.from(drawing.children).find((c) => c.namespaceURI === WP && (c.localName === "inline" || c.localName === "anchor"));
  if (!holder) return null;
  const extent = holder.getElementsByTagNameNS(WP, "extent")[0];
  const blip = drawing.getElementsByTagNameNS(NS.a, "blip")[0];
  const rid = blip?.getAttributeNS(NS.r, "embed");
  if (!extent || !rid) return null;
  const docPr = holder.getElementsByTagNameNS(WP, "docPr")[0];
  const base = { rid, alt: docPr?.getAttribute("descr") ?? "", width: emu(extent.getAttribute("cx")), height: emu(extent.getAttribute("cy")) };
  if (holder.localName === "inline") return { ...base, floating: false, x: 0, y: 0, relX: "column", relY: "paragraph", wrap: "square" };
  const position = (name: string) => {
    const el = holder.getElementsByTagNameNS(WP, name)[0];
    const offset = el?.getElementsByTagNameNS(WP, "posOffset")[0]?.textContent;
    return { rel: el?.getAttribute("relativeFrom") ?? (name === "positionH" ? "column" : "paragraph"), offset: emu(offset) };
  };
  const h = position("positionH");
  const v = position("positionV");
  const wrapEl = Array.from(holder.children).find((c) => c.namespaceURI === WP && c.localName.startsWith("wrap"));
  let wrap: Wrap = "square";
  if (!wrapEl || wrapEl.localName === "wrapNone") wrap = holder.getAttribute("behindDoc") === "1" ? "behind" : "none";
  else if (wrapEl.localName === "wrapTopAndBottom") wrap = "topAndBottom";
  return { ...base, floating: true, x: h.offset, y: v.offset, relX: h.rel, relY: v.rel, wrap };
}

const round = (points: number) => Math.round(points * EMU_PER_POINT);

/** A new picture's <w:drawing>, referring to its image by relationship id. */
/** A picture's XML with its description (alt text) set. */
export function withDescription(xml: string, alt: string): string {
  return xml.replace(/<wp:docPr\b([^>]*?)(\/?)>/, (_m, attrs: string, close: string) => {
    const rest = attrs.replace(/\sdescr="[^"]*"/, "");
    return `<wp:docPr${rest}${alt ? ` descr="${escapeAttr(alt)}"` : ""}${close}>`;
  });
}

export function pictureXml(rid: string, layout: PictureLayout, id: number, declareR: boolean): string {
  const cx = round(layout.width);
  const cy = round(layout.height);
  const r = declareR ? ` xmlns:r="${NS.r}"` : "";
  const graphic =
    `<a:graphic xmlns:a="${NS.a}"><a:graphicData uri="${PIC}"><pic:pic xmlns:pic="${PIC}">` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${escapeAttr(rid)}"${r}/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>`;
  const common = `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`;
  const tail = `<wp:docPr id="${id}" name="Picture ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr>${graphic}`;
  if (!layout.floating) return `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="${WP}">${common}${tail}</wp:inline></w:drawing>`;
  return (
    `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${251658240 + id}" ` +
    `behindDoc="${layout.wrap === "behind" ? 1 : 0}" locked="0" layoutInCell="1" allowOverlap="1" xmlns:wp="${WP}">` +
    `<wp:simplePos x="0" y="0"/>${positionXml(layout)}${common}${wrapXml(layout.wrap)}${tail}</wp:anchor></w:drawing>`
  );
}

function positionXml(layout: PictureLayout): string {
  return (
    `<wp:positionH relativeFrom="${escapeAttr(layout.relX)}"><wp:posOffset>${round(layout.x)}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="${escapeAttr(layout.relY)}"><wp:posOffset>${round(layout.y)}</wp:posOffset></wp:positionV>`
  );
}

function wrapXml(wrap: Wrap): string {
  return wrap === "square" ? '<wp:wrapSquare wrapText="bothSides"/>' : wrap === "topAndBottom" ? "<wp:wrapTopAndBottom/>" : "<wp:wrapNone/>";
}

/**
 * Updates a picture's original XML for a new size or position, keeping
 * everything else (cropping, alt text, effects). Returns null when the change
 * is too big for that (in line ↔ floating), so the picture is written afresh.
 */
export function adjustPictureXml(xml: string, before: PictureLayout, after: PictureLayout): string | null {
  if (before.floating !== after.floating) return null;
  let out = xml;
  if (before.width !== after.width || before.height !== after.height) {
    const size = `cx="${round(after.width)}" cy="${round(after.height)}"`;
    out = out.replace(/(<wp:extent\s)[^>]*?(\/>)/, `$1${size}$2`);
    out = out.replace(/(<a:xfrm\b[^>]*>[\s\S]*?<a:ext\s)[^>]*?(\/>)/, `$1${size}$2`);
  }
  if (after.floating) {
    if (before.x !== after.x || before.y !== after.y || before.relX !== after.relX || before.relY !== after.relY) {
      out = out.replace(/<wp:positionH\b[\s\S]*?<\/wp:positionH>\s*<wp:positionV\b[\s\S]*?<\/wp:positionV>/, positionXml(after));
    }
    if (before.wrap !== after.wrap) {
      out = out.replace(/<wp:wrap(None|Square|Tight|Through|TopAndBottom)\b[^>]*?(\/>|>[\s\S]*?<\/wp:wrap\1>)/, wrapXml(after.wrap));
      out = out.replace(/(<wp:anchor\b[^>]*\bbehindDoc=")[^"]*(")/, `$1${after.wrap === "behind" ? 1 : 0}$2`);
    }
  }
  return out;
}

/** The content type of an image, by its file extension. */
export function imageType(ext: string): string {
  const e = ext.toLowerCase();
  return e === "jpg" || e === "jpeg" ? "image/jpeg" : e === "gif" ? "image/gif" : e === "bmp" ? "image/bmp" : e === "webp" ? "image/webp" : e === "svg" ? "image/svg+xml" : e === "emf" ? "image/x-emf" : e === "wmf" ? "image/x-wmf" : e === "tif" || e === "tiff" ? "image/tiff" : "image/png";
}
