// The Word formats that are a .docx inside: with macros (.docm), templates
// (.dotx, .dotm), and Word's single-file XML document (.xml, "Flat OPC"),
// which holds the same parts in one XML file. They open as they are and are
// saved as what they are; saved under another type, the package is changed
// to match (Word refuses a .docx that says it has macros).

import JSZip from "jszip";

const MAIN = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  docm: "application/vnd.ms-word.document.macroEnabled.main+xml",
  dotx: "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
  dotm: "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
} as const;

export type PackageKind = keyof typeof MAIN;

/** The Word package kind a file name asks for; .xml is a .docx in one file. */
export function packageKind(name: string): PackageKind | null {
  const ext = /\.([a-z]+)$/i.exec(name)?.[1]?.toLowerCase();
  if (ext === "xml") return "docx";
  return ext && ext in MAIN ? (ext as PackageKind) : null;
}

/** Whether a file name is one DuckWriter opens as a Word document (not imported). */
export const isWordFile = (name: string) => /\.(docx|docm|dotx|dotm|xml)$/i.test(name);

const escapeAttr = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

/** The package as `kind`: the main part's content type, and no macros unless the kind has them. */
export async function convertPackage(bytes: Uint8Array, kind: PackageKind): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const typesXml = await zip.file("[Content_Types].xml")?.async("string");
  if (!typesXml) return bytes;
  const mainTypes = Object.values(MAIN) as string[];
  const current = mainTypes.find((t) => typesXml.includes(`ContentType="${t}"`));
  if (!current || current === MAIN[kind]) return bytes;
  let types = typesXml.split(`ContentType="${current}"`).join(`ContentType="${MAIN[kind]}"`);
  if (kind === "docx" || kind === "dotx") {
    // Without macros: the macro project and what belongs to it go.
    const relsPath = "word/_rels/document.xml.rels";
    const rels = await zip.file(relsPath)?.async("string");
    if (rels) {
      const macroParts: string[] = [];
      const kept = rels.replace(/<Relationship\b[^>]*\/>/g, (rel) => {
        if (!/Type="[^"]*\/vbaProject"/.test(rel)) return rel;
        const target = /Target="([^"]*)"/.exec(rel)?.[1];
        if (target) macroParts.push(target.startsWith("/") ? target.slice(1) : `word/${target}`);
        return "";
      });
      zip.file(relsPath, kept);
      for (const part of macroParts) {
        const name = part.split("/").pop()!;
        const relsOfPart = part.replace(/[^/]+$/, `_rels/${name}.rels`);
        const linked = await zip.file(relsOfPart)?.async("string");
        for (const t of linked ? [...linked.matchAll(/Target="([^"]*)"/g)].map((m) => part.replace(/[^/]+$/, m[1])) : []) zip.remove(t);
        zip.remove(part);
        zip.remove(relsOfPart);
      }
      types = types.replace(/<Override\b[^>]*PartName="([^"]*)"[^>]*\/>/g, (o, partName: string) => (zip.file(partName.replace(/^\//, "")) ? o : ""));
      if (!Object.keys(zip.files).some((f) => /\.bin$/i.test(f))) types = types.replace(/<Default\b[^>]*Extension="bin"[^>]*\/>/i, "");
    }
  }
  zip.file("[Content_Types].xml", types);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

// ---- Word's single-file XML document (Flat OPC) ----------------------------------------

const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";

/** Whether text is a Word XML document in one file (a pkg:package). */
export const isFlatOpc = (text: string) => /<pkg:package\b/.test(text.slice(0, 4000)) || text.includes(PKG);

/** A Word XML document (one file) as the .docx package it holds. */
export async function fromFlatOpc(text: string): Promise<Uint8Array> {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("This XML file is damaged.");
  const parts = [...doc.getElementsByTagNameNS(PKG, "part")];
  if (!parts.length) throw new Error("This XML file isn't a Word XML document.");
  const zip = new JSZip();
  const overrides: string[] = [];
  for (const part of parts) {
    const name = (part.getAttributeNS(PKG, "name") ?? part.getAttribute("pkg:name") ?? "").replace(/^\//, "");
    const type = part.getAttributeNS(PKG, "contentType") ?? part.getAttribute("pkg:contentType") ?? "";
    if (!name) continue;
    const xmlData = part.getElementsByTagNameNS(PKG, "xmlData")[0];
    const binaryData = part.getElementsByTagNameNS(PKG, "binaryData")[0];
    if (xmlData) {
      const root = Array.from(xmlData.childNodes).find((n) => n.nodeType === 1) as Element | undefined;
      if (!root) continue;
      zip.file(name, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${new XMLSerializer().serializeToString(root)}`);
    } else if (binaryData) {
      zip.file(name, (binaryData.textContent ?? "").replace(/\s+/g, ""), { base64: true });
    }
    if (type) overrides.push(`<Override PartName="/${escapeAttr(name)}" ContentType="${escapeAttr(type)}"/>`);
  }
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">${overrides.join("")}</Types>`);
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

/** The .docx package as one Word XML document. */
export async function toFlatOpc(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const typesXml = (await zip.file("[Content_Types].xml")?.async("string")) ?? "";
  const overrides = new Map([...typesXml.matchAll(/<Override\b[^>]*PartName="([^"]*)"[^>]*ContentType="([^"]*)"/g)].map((m) => [m[1].replace(/^\//, ""), m[2]]));
  const defaults = new Map([...typesXml.matchAll(/<Default\b[^>]*Extension="([^"]*)"[^>]*ContentType="([^"]*)"/g)].map((m) => [m[1].toLowerCase(), m[2]]));
  const out: string[] = [];
  for (const name of Object.keys(zip.files).sort((a, b) => Number(!a.startsWith("_rels/")) - Number(!b.startsWith("_rels/")) || a.localeCompare(b))) {
    if (zip.files[name].dir || name === "[Content_Types].xml") continue;
    const type = overrides.get(name) ?? defaults.get(name.split(".").pop()!.toLowerCase()) ?? "application/octet-stream";
    const isXml = /xml$/i.test(type) || /\.(xml|rels)$/i.test(name);
    if (isXml) {
      const xml = (await zip.file(name)!.async("string")).replace(/^﻿?<\?xml[^>]*\?>\s*/, "");
      out.push(`<pkg:part pkg:name="/${escapeAttr(name)}" pkg:contentType="${escapeAttr(type)}"><pkg:xmlData>${xml}</pkg:xmlData></pkg:part>`);
    } else {
      const data = await zip.file(name)!.async("base64");
      out.push(`<pkg:part pkg:name="/${escapeAttr(name)}" pkg:contentType="${escapeAttr(type)}" pkg:compression="store"><pkg:binaryData>${data}</pkg:binaryData></pkg:part>`);
    }
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<?mso-application progid="Word.Document"?>\n<pkg:package xmlns:pkg="${PKG}">${out.join("")}</pkg:package>\n`;
}

/** A Word XML file's bytes as a .docx package, or why it can't be opened. */
export async function wordXml(bytes: Uint8Array, name: string): Promise<Uint8Array> {
  const text = new TextDecoder().decode(bytes);
  if (isFlatOpc(text)) return fromFlatOpc(text);
  if (text.includes("schemas.microsoft.com/office/word/2003/wordml")) throw new Error(`${name} is a Word 2003 XML document. Open it in Word and save it as .docx first.`);
  throw new Error(`${name} isn't a Word XML document.`);
}

/** A Word file's bytes as a .docx package: as they are, or unpacked from Word XML. */
export async function asPackage(bytes: Uint8Array, name = "This file"): Promise<Uint8Array> {
  // A zip package starts with "PK".
  return bytes[0] === 0x50 && bytes[1] === 0x4b ? bytes : wordXml(bytes, name);
}
