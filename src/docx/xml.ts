// Small helpers for the XML inside a .docx file.

export const NS = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  ct: "http://schemas.openxmlformats.org/package/2006/content-types",
} as const;

export function parseXml(text: string): Document {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  const error = doc.getElementsByTagName("parsererror")[0];
  if (error) throw new Error(`Damaged XML: ${error.textContent?.slice(0, 200)}`);
  return doc;
}

/** The element's children, as elements. */
export const children = (el: Element): Element[] => Array.from(el.children);

/** Is this a WordprocessingML element called `name`? */
export const isW = (el: Element | null | undefined, name: string): boolean => !!el && el.namespaceURI === NS.w && el.localName === name;

/** The first child element in the `w:` namespace called `name`. */
export function child(el: Element | null | undefined, name: string): Element | null {
  if (!el) return null;
  for (const c of el.children) if (c.namespaceURI === NS.w && c.localName === name) return c;
  return null;
}

export function childrenNamed(el: Element | null | undefined, name: string): Element[] {
  return el ? children(el).filter((c) => isW(c, name)) : [];
}

/** A `w:` attribute (Word puts its attributes in its own namespace). */
export function attr(el: Element | null | undefined, name: string): string | null {
  if (!el) return null;
  return el.getAttributeNS(NS.w, name) ?? el.getAttribute(`w:${name}`);
}

export function numAttr(el: Element | null | undefined, name: string): number | null {
  const value = attr(el, name);
  if (value === null || value === "") return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

/**
 * A toggle property such as <w:b/>: present means on, unless its value says
 * otherwise. Returns null when the element is absent.
 */
export function toggle(el: Element | null | undefined, name: string): boolean | null {
  const c = child(el, name);
  if (!c) return null;
  const value = attr(c, "val");
  return value === null || !["0", "false", "off", "none"].includes(value);
}

export const serialize = (node: Node): string => new XMLSerializer().serializeToString(node);

export function escapeText(text: string): string {
  return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export function escapeAttr(text: string): string {
  return escapeText(text).replace(/"/g, "&quot;");
}

/** Builds an element as a string: `tag("w:b")`, `tag("w:sz", { "w:val": 24 })`, with children. */
export function tag(name: string, attrs: Record<string, string | number | null | undefined> = {}, ...content: string[]): string {
  let out = `<${name}`;
  for (const [key, value] of Object.entries(attrs)) {
    if (value !== null && value !== undefined) out += ` ${key}="${escapeAttr(String(value))}"`;
  }
  const inner = content.join("");
  return inner ? `${out}>${inner}</${name}>` : `${out}/>`;
}

/** An element's start tag exactly as written, prefixes and all: `<w:r w:rsidR="00AB">` (without the closing ">"). */
export function startTag(el: Element): string {
  let out = `<${el.tagName}`;
  for (const a of Array.from(el.attributes)) out += ` ${a.name}="${escapeAttr(a.value)}"`;
  return out;
}

/** Writes an element back out with the same names and prefixes it was read with. */
function verbatim(node: Node): string {
  if (node.nodeType === 3) return escapeText(node.nodeValue ?? "");
  if (node.nodeType === 4) return `<![CDATA[${node.nodeValue ?? ""}]]>`;
  if (node.nodeType === 8) return `<!--${node.nodeValue ?? ""}-->`;
  if (node.nodeType !== 1) return "";
  const el = node as Element;
  let inner = "";
  for (const c of Array.from(el.childNodes)) inner += verbatim(c);
  return inner ? `${startTag(el)}>${inner}</${el.tagName}>` : `${startTag(el)}/>`;
}

/**
 * An element taken from a docx part, as XML to put back in the same part:
 * byte for byte what was read, apart from declarations added for any prefix
 * it uses that only an ancestor below the part's root declared.
 */
export function rawXml(el: Element, rootDeclares: Set<string>): string {
  const xml = verbatim(el);
  const used = new Set<string>();
  for (const m of xml.matchAll(/[<\s/]([\w.-]+):[\w.-]+(?==|[\s/>])/g)) used.add(m[1]);
  const declared = new Set([...xml.matchAll(/\sxmlns:([\w.-]+)=/g)].map((m) => m[1]));
  let extra = "";
  for (const prefix of used) {
    if (prefix === "xml" || prefix === "xmlns" || declared.has(prefix) || rootDeclares.has(prefix)) continue;
    const uri = el.lookupNamespaceURI(prefix);
    if (uri) extra += ` xmlns:${prefix}="${escapeAttr(uri)}"`;
  }
  return extra ? xml.replace(/^(<[^\s/>]+)/, `$1${extra}`) : xml;
}

/** The namespace prefixes a root element declares. */
export function declaredPrefixes(root: Element): Set<string> {
  const out = new Set<string>();
  for (const a of Array.from(root.attributes)) {
    if (a.name === "xmlns") out.add("");
    else if (a.name.startsWith("xmlns:")) out.add(a.name.slice(6));
  }
  return out;
}
