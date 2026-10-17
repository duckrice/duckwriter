// Pasting from Word and the web: lists become the document's own lists
// (Word's "mso-list" paragraphs and HTML's ul/ol alike), headings and the
// basic formatting come through (see the schema's rules).

import { Fragment, Slice, type Node } from "prosemirror-model";
import { Plugin } from "prosemirror-state";
import type { DocContext } from "../docx/model";
import { createList } from "./lists";
import { schema } from "./schema";

/** Rewrites pasted HTML so each list item is a paragraph saying which list and level it is in. */
export function flattenPastedLists(html: string): string {
  if (!/<(ul|ol)\b|mso-list/i.test(html)) return html;
  const doc = new DOMParser().parseFromString(html, "text/html");
  // Word: <p style="mso-list:l0 level2 lfo1"><span style="mso-list:Ignore">1.<span> </span></span>Text</p>
  for (const p of doc.querySelectorAll("p")) {
    const level = /mso-list:\s*l\d+\s+level(\d+)/i.exec(p.getAttribute("style") ?? "");
    if (!level) continue;
    const marker = [...p.querySelectorAll("span")].find((s) => /mso-list:\s*Ignore/i.test(s.getAttribute("style") ?? ""));
    const text = (marker?.textContent ?? "").replace(/\s+/g, "");
    marker?.remove();
    p.setAttribute("data-list", /^(\d+|[a-z]|[ivxlc]+)[.)]$/i.test(text) ? "number" : "bullet");
    p.setAttribute("data-level", String(Number(level[1]) - 1));
  }
  // HTML lists: each item a paragraph, a list inside an item a level deeper.
  const top = [...doc.querySelectorAll("ul, ol")].filter((l) => !l.parentElement?.closest("ul, ol"));
  for (const list of top) {
    const out: HTMLElement[] = [];
    const walk = (l: Element, level: number) => {
      for (const li of [...l.children]) {
        if (li.tagName !== "LI") continue;
        const p = doc.createElement("p");
        p.setAttribute("data-list", l.tagName === "OL" ? "number" : "bullet");
        p.setAttribute("data-level", String(Math.min(level, 8)));
        const nested: Element[] = [];
        for (const child of [...li.childNodes]) {
          if (child.nodeType === 1 && /^(UL|OL)$/.test((child as Element).tagName)) nested.push(child as Element);
          else if (child.nodeType === 1 && /^(P|DIV)$/.test((child as Element).tagName)) {
            if (p.childNodes.length) p.append(" ");
            p.append(...(child as Element).childNodes);
          } else p.append(child);
        }
        out.push(p);
        for (const n of nested) walk(n, level + 1);
      }
    };
    walk(list, 0);
    list.replaceWith(...out);
  }
  return doc.body.innerHTML;
}

/** Pasted list paragraphs join one new list of each kind in the document. */
export function realLists(fragment: Fragment, ctx: DocContext, made: Record<string, string>): Fragment {
  const out: Node[] = [];
  fragment.forEach((node) => {
    const numId = node.type === schema.nodes.paragraph ? (node.attrs.numId as string | null) : null;
    if (numId?.startsWith("__")) {
      const kind = numId.slice(2) as "bullet" | "number";
      made[kind] ??= createList(ctx, kind);
      out.push(node.type.create({ ...node.attrs, numId: made[kind], style: node.attrs.style ?? (ctx.styles.byId.has("ListParagraph") ? "ListParagraph" : null) }, node.content, node.marks));
    } else out.push(node.isLeaf ? node : node.copy(realLists(node.content, ctx, made)));
  });
  return Fragment.from(out);
}

export function pastePlugin(ctx: () => DocContext) {
  /** The first pasted paragraph's list, when it is joined to the paragraph pasted into. */
  let firstList: { numId: string; ilvl: number; style: string | null } | null = null;
  return new Plugin({
    props: {
      transformPastedHTML: flattenPastedLists,
      transformPasted: (slice) => {
        const content = realLists(slice.content, ctx(), {});
        const first = content.firstChild;
        firstList = slice.openStart > 0 && first?.type === schema.nodes.paragraph && first.attrs.numId ? { numId: first.attrs.numId, ilvl: first.attrs.ilvl, style: first.attrs.style } : null;
        return new Slice(content, slice.openStart, slice.openEnd);
      },
    },
    // Pasted into an empty paragraph (or one that isn't a list), the first item is still a list item.
    appendTransaction(trs, oldState, newState) {
      const list = firstList;
      firstList = null;
      const paste = trs.find((tr) => tr.getMeta("uiEvent") === "paste");
      if (!list || !paste) return null;
      const $at = newState.doc.resolve(paste.mapping.map(oldState.selection.from, -1));
      const para = $at.parent;
      if (para.type !== schema.nodes.paragraph || para.attrs.numId) return null;
      return newState.tr.setNodeMarkup($at.before(), null, { ...para.attrs, ...list });
    },
  });
}
