// OpenDocument text (.odt, LibreOffice): opened as a new Word document, as
// HTML the importer reads (headings, bold, italic, lists, tables…).

import JSZip from "jszip";

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

interface Style {
  parent: string | null;
  text: Element | null;
  para: Element | null;
  list: string | null;
  outline: number;
}

export async function odtToHtml(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  const parse = async (name: string) => {
    const xml = await zip.file(name)?.async("string");
    return xml ? new DOMParser().parseFromString(xml, "application/xml") : null;
  };
  const content = await parse("content.xml");
  if (!content) throw new Error("This OpenDocument file has no text in it.");
  const docs = [await parse("styles.xml"), content].filter((d): d is Document => !!d);

  // Styles by name (paragraph and text), and list styles: bullet or numbered, per level.
  const styles = new Map<string, Style>();
  const lists = new Map<string, boolean[]>();
  for (const doc of docs) {
    for (const el of doc.getElementsByTagName("style:style")) {
      styles.set(el.getAttribute("style:name") ?? "", {
        parent: el.getAttribute("style:parent-style-name"),
        text: el.getElementsByTagName("style:text-properties")[0] ?? null,
        para: el.getElementsByTagName("style:paragraph-properties")[0] ?? null,
        list: el.getAttribute("style:list-style-name"),
        outline: Number(el.getAttribute("style:default-outline-level")) || 0,
      });
    }
    for (const el of doc.getElementsByTagName("text:list-style")) {
      const levels: boolean[] = [];
      for (const level of el.children) levels[Number(level.getAttribute("text:level") ?? 1) - 1] = level.localName === "list-level-style-number";
      lists.set(el.getAttribute("style:name") ?? "", levels);
    }
  }
  /** A property of a style or the styles it is based on. */
  const prop = (name: string | null, get: (s: Style) => string | null | undefined): string | null => {
    for (let n = name, depth = 0; n && depth < 20; depth++) {
      const s = styles.get(n);
      if (!s) return null;
      const v = get(s);
      if (v) return v;
      n = s.parent;
    }
    return null;
  };
  const chain = (name: string | null) => {
    const names: string[] = [];
    for (let n = name; n && names.length < 20; n = styles.get(n)?.parent ?? null) names.push(n);
    return names;
  };

  /** Text formatting of a style, as HTML tags around the content. */
  const wrap = (name: string | null, html: string) => {
    const t = (attr: string) => prop(name, (s) => s.text?.getAttribute(attr));
    if (t("fo:font-weight") === "bold") html = `<b>${html}</b>`;
    if (t("fo:font-style") === "italic") html = `<i>${html}</i>`;
    const underline = t("style:text-underline-style");
    if (underline && underline !== "none") html = `<u>${html}</u>`;
    const strike = t("style:text-line-through-style");
    if (strike && strike !== "none") html = `<s>${html}</s>`;
    const position = t("style:text-position") ?? "";
    if (/^(super|[1-9])/.test(position)) html = `<sup>${html}</sup>`;
    else if (/^(sub|-)/.test(position)) html = `<sub>${html}</sub>`;
    return html;
  };

  const notes: string[] = [];
  const inline = (el: Element): string => {
    let out = "";
    for (const node of el.childNodes) {
      if (node.nodeType === 3) {
        out += esc(node.textContent ?? "");
        continue;
      }
      if (node.nodeType !== 1) continue;
      const child = node as Element;
      switch (child.localName) {
        case "span":
          out += wrap(child.getAttribute("text:style-name"), inline(child));
          break;
        case "a":
          out += `<a href="${esc(child.getAttribute("xlink:href") ?? "")}">${inline(child)}</a>`;
          break;
        case "s":
          out += " ".repeat(Number(child.getAttribute("text:c")) || 1);
          break;
        case "tab":
          out += "\t";
          break;
        case "line-break":
          out += "<br>";
          break;
        case "note": {
          // shortcut: a footnote becomes [n] and a note at the end; real footnotes when the importer can make them.
          const body = child.getElementsByTagName("text:note-body")[0];
          notes.push(body ? [...body.children].map(inline).join(" ") : "");
          out += `<sup>[${notes.length}]</sup>`;
          break;
        }
        case "note-citation":
        case "bookmark":
        case "bookmark-start":
        case "bookmark-end":
        case "soft-page-break":
          break;
        default:
          out += inline(child);
      }
    }
    return out;
  };

  const paragraph = (el: Element, list?: { numbered: boolean; level: number }) => {
    const style = el.getAttribute("text:style-name");
    const names = chain(style);
    let level = el.localName === "h" ? Number(el.getAttribute("text:outline-level")) || 1 : Number(prop(style, (s) => (s.outline ? String(s.outline) : null))) || 0;
    if (!level) level = Number(names.map((n) => /^Heading_20_(\d)$/.exec(n)?.[1]).find(Boolean)) || 0;
    const align = prop(style, (s) => s.para?.getAttribute("fo:text-align"));
    const css = align === "center" || align === "end" || align === "justify" ? ` style="text-align:${align === "end" ? "right" : align}"` : "";
    const content = wrap(style, inline(el));
    if (list) return `<p data-list="${list.numbered ? "number" : "bullet"}" data-level="${list.level}"${css}>${content}</p>`;
    if (names.includes("Title")) return `<p class="MsoTitle"${css}>${content}</p>`;
    if (level) return `<h${Math.min(level, 6)}${css}>${content}</h${Math.min(level, 6)}>`;
    return `<p${css}>${content}</p>`;
  };

  const blocks = (el: Element, list?: { style: string | null; level: number }): string => {
    let out = "";
    for (const child of el.children) {
      switch (child.localName) {
        case "p":
        case "h": {
          if (!list) {
            out += paragraph(child);
            break;
          }
          const listStyle = list.style ?? prop(child.getAttribute("text:style-name"), (s) => s.list);
          out += paragraph(child, { numbered: !!lists.get(listStyle ?? "")?.[list.level], level: list.level });
          break;
        }
        case "list": {
          const style = child.getAttribute("text:style-name") ?? list?.style ?? null;
          for (const item of child.children) if (item.localName === "list-item" || item.localName === "list-header") out += blocks(item, { style, level: list ? list.level + 1 : 0 });
          break;
        }
        case "table": {
          const rows = [...child.getElementsByTagName("table:table-row")];
          out += `<table>${rows.map((r) => `<tr>${[...r.children].filter((c) => c.localName === "table-cell").map((c) => `<td>${blocks(c) || "<p></p>"}</td>`).join("")}</tr>`).join("")}</table>`;
          break;
        }
        case "section":
        case "index-body":
        case "table-of-content":
          out += blocks(child, list);
          break;
      }
    }
    return out;
  };

  const text = content.getElementsByTagName("office:text")[0];
  if (!text) throw new Error("This OpenDocument file has no text in it.");
  let html = blocks(text);
  if (notes.length) html += `<h2>Notes</h2>${notes.map((n, i) => `<p>[${i + 1}] ${n}</p>`).join("")}`;
  return html;
}
