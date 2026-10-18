// Checks a saved .docx the way Word would before calling it damaged: every
// part is well-formed XML with a content type, every relationship points to a
// part that is there, the text refers only to relationships, comments and
// notes that exist, and ids that must be unique are. Used by the tests, and on
// every save while the app runs its own checks.

import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

const parse = (xml: string) => new DOMParser().parseFromString(xml, "application/xml");
const broken = (doc: Document) => doc.getElementsByTagName("parsererror").length > 0 || doc.documentElement?.nodeName === "parsererror";

/** Where a relationship's target is, from the part it belongs to. */
function resolve(from: string, target: string): string {
  if (target.startsWith("/")) return target.slice(1);
  const parts = from.split("/").slice(0, -1);
  for (const piece of target.split("/")) {
    if (piece === "..") parts.pop();
    else if (piece !== ".") parts.push(piece);
  }
  return parts.join("/");
}

/** The problems found; none means Word should open it without complaint. */
export async function checkPackage(bytes: Uint8Array): Promise<string[]> {
  const problems: string[] = [];
  const zip = await JSZip.loadAsync(bytes);
  const files = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const has = (name: string) => !!zip.file(name);

  // Content types.
  const typesXml = await zip.file("[Content_Types].xml")?.async("string");
  if (!typesXml) return ["[Content_Types].xml is missing"];
  const types = parse(typesXml);
  const defaults = new Set([...types.getElementsByTagName("Default")].map((d) => (d.getAttribute("Extension") ?? "").toLowerCase()));
  const overrides = new Set([...types.getElementsByTagName("Override")].map((o) => (o.getAttribute("PartName") ?? "").replace(/^\//, "")));
  for (const name of files) {
    if (name === "[Content_Types].xml") continue;
    const ext = name.split(".").pop()!.toLowerCase();
    if (!overrides.has(name) && !defaults.has(ext)) problems.push(`${name} has no content type`);
  }
  for (const name of overrides) if (!has(name)) problems.push(`a content type is given for ${name}, which isn't there`);

  // Well-formed XML.
  const xml = new Map<string, Document>();
  for (const name of files.filter((n) => /\.(xml|rels)$/i.test(n))) {
    const doc = parse(await zip.file(name)!.async("string"));
    if (broken(doc)) problems.push(`${name} isn't well-formed XML`);
    else xml.set(name, doc);
  }

  // Relationships: targets there, ids unique.
  const relsOf = new Map<string, Map<string, string>>();
  for (const [name, doc] of xml) {
    if (!name.endsWith(".rels")) continue;
    // word/_rels/document.xml.rels belongs to word/document.xml; _rels/.rels to the package.
    const source = name.replace(/_rels\/([^/]*)\.rels$/, "$1");
    const ids = new Map<string, string>();
    for (const rel of doc.getElementsByTagName("Relationship")) {
      const id = rel.getAttribute("Id") ?? "";
      if (ids.has(id)) problems.push(`${name}: relationship ${id} is there twice`);
      ids.set(id, rel.getAttribute("Type") ?? "");
      if (rel.getAttribute("TargetMode") === "External") continue;
      const target = resolve(source, rel.getAttribute("Target") ?? "");
      if (!has(target)) problems.push(`${name}: ${id} points to ${target}, which isn't there`);
    }
    relsOf.set(source, ids);
  }

  // The story parts: their relationship ids, notes and comments referred to.
  const notesIn = (part: string, tag: string) => {
    const doc = xml.get(part);
    return new Set(doc ? [...doc.getElementsByTagNameNS(W, tag)].map((n) => n.getAttributeNS(W, "id") ?? "") : []);
  };
  /** A part's relationships: their type and the part they point to. */
  const targets = (part: string) => {
    const rels = xml.get(part.replace(/([^/]*)$/, "_rels/$1.rels"));
    return rels ? [...rels.getElementsByTagName("Relationship")].filter((r) => r.getAttribute("TargetMode") !== "External").map((r) => ({ type: r.getAttribute("Type") ?? "", part: resolve(part, r.getAttribute("Target") ?? "") })) : [];
  };
  const main = targets("").find((t) => t.type.endsWith("/officeDocument"))?.part ?? "word/document.xml";
  const ofType = (type: string) => targets(main).find((t) => t.type.endsWith(`/${type}`))?.part ?? null;
  const footnotes = ofType("footnotes");
  const endnotes = ofType("endnotes");
  const commentsPart = ofType("comments");
  const footnoteIds = footnotes ? notesIn(footnotes, "footnote") : new Set<string>();
  const endnoteIds = endnotes ? notesIn(endnotes, "endnote") : new Set<string>();
  const commentIds = commentsPart ? notesIn(commentsPart, "comment") : new Set<string>();
  if (commentsPart) {
    const all = [...(xml.get(commentsPart)?.getElementsByTagNameNS(W, "comment") ?? [])].map((c) => c.getAttributeNS(W, "id"));
    if (new Set(all).size !== all.length) problems.push(`${commentsPart}: a comment id is there twice`);
  }

  const stories = [main, ...targets(main).filter((t) => /\/(header|footer|footnotes|endnotes|comments)$/.test(t.type)).map((t) => t.part)];
  const pictureIds = new Map<string, string>();
  const bookmarks = new Map<string, string>();
  for (const part of stories) {
    const doc = xml.get(part);
    if (!doc) continue;
    const rels = relsOf.get(part) ?? new Map();
    for (const el of doc.getElementsByTagName("*")) {
      for (const name of ["id", "embed", "link", "pict"]) {
        const value = el.getAttributeNS(R, name);
        if (value && !rels.has(value)) problems.push(`${part}: ${el.nodeName} refers to relationship ${value}, which isn't there`);
      }
    }
    for (const ref of doc.getElementsByTagNameNS(W, "footnoteReference")) {
      const id = ref.getAttributeNS(W, "id") ?? "";
      if (!footnoteIds.has(id)) problems.push(`${part}: footnote ${id} is referred to but not in the footnotes`);
    }
    for (const ref of doc.getElementsByTagNameNS(W, "endnoteReference")) {
      const id = ref.getAttributeNS(W, "id") ?? "";
      if (!endnoteIds.has(id)) problems.push(`${part}: endnote ${id} is referred to but not in the endnotes`);
    }
    for (const tag of ["commentRangeStart", "commentRangeEnd", "commentReference"]) {
      for (const ref of doc.getElementsByTagNameNS(W, tag)) {
        const id = ref.getAttributeNS(W, "id") ?? "";
        if (!commentIds.has(id)) problems.push(`${part}: ${tag} ${id} has no comment`);
      }
    }
    // Bookmarks: ids unique in the document, each start with an end.
    const ends = new Set([...doc.getElementsByTagNameNS(W, "bookmarkEnd")].map((b) => b.getAttributeNS(W, "id") ?? ""));
    for (const b of doc.getElementsByTagNameNS(W, "bookmarkStart")) {
      const id = b.getAttributeNS(W, "id") ?? "";
      if (bookmarks.has(id)) problems.push(`${part}: bookmark id ${id} is used twice (also in ${bookmarks.get(id)})`);
      bookmarks.set(id, part);
      if (!ends.has(id)) problems.push(`${part}: bookmark ${b.getAttributeNS(W, "name")} (${id}) has no end`);
    }
    // Pictures: the drawing ids must differ across the document.
    for (const pr of doc.getElementsByTagNameNS(WP, "docPr")) {
      const id = pr.getAttribute("id") ?? "";
      if (pictureIds.has(id)) problems.push(`${part}: picture id ${id} is used twice (also in ${pictureIds.get(id)})`);
      pictureIds.set(id, part);
    }
  }
  return problems;
}
