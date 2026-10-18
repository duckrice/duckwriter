// The document's properties (docProps/core.xml): title, subject, author,
// keywords, description, and who saved it last and when. Word shows them
// under File → Info; search engines and reference managers read the title.
// Saving updates "last saved by", the date and the revision, as Word does.

import type JSZip from "jszip";

export interface DocProperties {
  title: string;
  subject: string;
  /** The author (dc:creator). */
  creator: string;
  keywords: string;
  description: string;
  lastModifiedBy: string;
  /** When it was made and last saved (ISO dates), as read. */
  created: string;
  modified: string;
  revision: string;
  /** "Remove names" in Check before sending: saving leaves the names out. */
  anonymous: boolean;
}

const NS = {
  cp: "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  dc: "http://purl.org/dc/elements/1.1/",
  dcterms: "http://purl.org/dc/terms/",
  xsi: "http://www.w3.org/2001/XMLSchema-instance",
};

/** Where each property is kept: namespace prefix and element name. */
const FIELDS: [keyof DocProperties, keyof typeof NS, string][] = [
  ["title", "dc", "title"],
  ["subject", "dc", "subject"],
  ["creator", "dc", "creator"],
  ["keywords", "cp", "keywords"],
  ["description", "dc", "description"],
  ["lastModifiedBy", "cp", "lastModifiedBy"],
  ["revision", "cp", "revision"],
  ["created", "dcterms", "created"],
  ["modified", "dcterms", "modified"],
];

const REL_TYPE = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

export const emptyProperties = (): DocProperties => ({ title: "", subject: "", creator: "", keywords: "", description: "", lastModifiedBy: "", created: "", modified: "", revision: "", anonymous: false });

/** Where the properties part is, from the package's relationships. */
async function corePath(zip: JSZip): Promise<string | null> {
  const rels = await zip.file("_rels/.rels")?.async("string");
  if (rels) {
    const doc = new DOMParser().parseFromString(rels, "application/xml");
    for (const rel of doc.getElementsByTagName("Relationship")) {
      if (rel.getAttribute("Type") === REL_TYPE) return (rel.getAttribute("Target") ?? "").replace(/^\//, "") || null;
    }
  }
  return zip.file("docProps/core.xml") ? "docProps/core.xml" : null;
}

export async function readProperties(zip: JSZip): Promise<DocProperties> {
  const props = emptyProperties();
  const path = await corePath(zip);
  const xml = path ? await zip.file(path)?.async("string") : undefined;
  if (!xml) return props;
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  for (const [key, prefix, name] of FIELDS) {
    const el = doc.getElementsByTagNameNS(NS[prefix], name)[0];
    if (el && key !== "anonymous") (props[key] as string) = el.textContent?.trim() ?? "";
  }
  return props;
}

/**
 * The properties as saved now: "last saved by" (unless the names are to be
 * left out), the date, the next revision. A new copy; the document's own are
 * updated only when it is really saved (not for a recovery copy).
 */
export function stamped(props: DocProperties, savedBy: string | null, now = new Date()): DocProperties {
  const out = { ...props };
  const stamp = now.toISOString().replace(/\.\d+Z$/, "Z");
  if (out.anonymous) {
    out.creator = "";
    out.lastModifiedBy = "";
  } else if (savedBy) {
    out.lastModifiedBy = savedBy;
    if (!out.creator) out.creator = savedBy;
  }
  if (!out.created) out.created = stamp;
  out.modified = stamp;
  out.revision = String((Number(out.revision) || 0) + 1);
  return out;
}

/** Writes the properties, stamped as saved now. */
export async function writeProperties(zip: JSZip, current: DocProperties, savedBy: string | null, now = new Date()) {
  let path = await corePath(zip);
  let xml = path ? await zip.file(path)?.async("string") : undefined;
  if (!path || !xml) {
    // Some programs leave the part out; Word always has one.
    path = "docProps/core.xml";
    xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="${NS.cp}" xmlns:dc="${NS.dc}" xmlns:dcterms="${NS.dcterms}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="${NS.xsi}"></cp:coreProperties>`;
    await addPart(zip, path);
  }
  const props = stamped(current, savedBy, now);

  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  for (const [key, prefix, name] of FIELDS) {
    const value = props[key] as string;
    let el = doc.getElementsByTagNameNS(NS[prefix], name)[0];
    if (!value) {
      el?.remove();
      continue;
    }
    if (!el) {
      el = doc.createElementNS(NS[prefix], `${prefix}:${name}`);
      if (prefix === "dcterms") el.setAttributeNS(NS.xsi, "xsi:type", "dcterms:W3CDTF");
      root.append(el);
    }
    el.textContent = value;
  }
  let out = new XMLSerializer().serializeToString(doc);
  if (!out.startsWith("<?xml")) out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${out}`;
  zip.file(path, out);
}

/** A new properties part: its relationship and its content type. */
async function addPart(zip: JSZip, path: string) {
  const rels = (await zip.file("_rels/.rels")?.async("string")) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  let n = 1;
  while (rels.includes(`Id="rId${n}"`)) n++;
  zip.file("_rels/.rels", rels.replace("</Relationships>", `<Relationship Id="rId${n}" Type="${REL_TYPE}" Target="${path}"/></Relationships>`));
  const types = await zip.file("[Content_Types].xml")?.async("string");
  if (types && !types.includes(`PartName="/${path}"`)) {
    zip.file("[Content_Types].xml", types.replace("</Types>", `<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`));
  }
}
