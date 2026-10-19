// Word 97–2003 documents (.doc): opened as a new Word document, as HTML the
// importer reads. The text, paragraphs, headings (by style), bold, italic,
// underline, strike-through, raised and lowered text, alignment, lists,
// tables, field results and footnotes; fonts, sizes, colours, pictures,
// headers and comments are left out. Saving goes to .docx.
//
// A .doc is a little file system ("compound file"); its WordDocument stream
// starts with the FIB, which says where the text pieces, formatting pages,
// styles and lists are, mostly in the table stream (0Table or 1Table).

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** The streams of a compound file, by name. */
function streams(bytes: Uint8Array): Map<string, Uint8Array> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u32 = (at: number) => view.getUint32(at, true);
  const sector = 1 << view.getUint16(0x1e, true);
  const mini = 1 << view.getUint16(0x20, true);
  const cutoff = u32(0x38);
  const at = (s: number) => (s + 1) * sector;
  // The FAT, from the sectors the header (and the DIFAT chain) lists.
  const fatSectors: number[] = [];
  for (let i = 0; i < 109; i++) fatSectors.push(u32(0x4c + i * 4));
  for (let d = u32(0x44), guard = 0; d < 0xfffffffa && guard < 10000; guard++) {
    for (let i = 0; i < sector / 4 - 1; i++) fatSectors.push(u32(at(d) + i * 4));
    d = u32(at(d) + sector - 4);
  }
  const fat: number[] = [];
  for (const s of fatSectors.slice(0, u32(0x2c))) for (let i = 0; i < sector / 4; i++) fat.push(u32(at(s) + i * 4));
  const chain = (start: number, next: number[]) => {
    const out: number[] = [];
    for (let s = start; s < 0xfffffffa && out.length <= next.length; s = next[s]) out.push(s);
    return out;
  };
  const read = (start: number, size: number) => {
    const out = new Uint8Array(chain(start, fat).length * sector);
    chain(start, fat).forEach((s, i) => out.set(bytes.subarray(at(s), at(s) + sector), i * sector));
    return out.subarray(0, size);
  };
  const dir = read(u32(0x30), chain(u32(0x30), fat).length * sector);
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  const entries: { name: string; type: number; start: number; size: number }[] = [];
  for (let o = 0; o + 128 <= dir.length; o += 128) {
    const len = dv.getUint16(o + 0x40, true);
    const name = String.fromCharCode(...Array.from({ length: Math.max(0, len / 2 - 1) }, (_, i) => dv.getUint16(o + i * 2, true)));
    entries.push({ name, type: dv.getUint8(o + 0x42), start: dv.getUint32(o + 0x74, true), size: dv.getUint32(o + 0x78, true) });
  }
  const root = entries.find((e) => e.type === 5);
  const miniStream = root ? read(root.start, root.size) : new Uint8Array();
  const miniFat: number[] = [];
  if (u32(0x40)) {
    const m = read(u32(0x3c), u32(0x40) * sector);
    for (let i = 0; i + 4 <= m.length; i += 4) miniFat.push(new DataView(m.buffer, m.byteOffset).getUint32(i, true));
  }
  const out = new Map<string, Uint8Array>();
  for (const e of entries) {
    if (e.type !== 2) continue;
    if (e.size < cutoff) {
      const data = new Uint8Array(chain(e.start, miniFat).length * mini);
      chain(e.start, miniFat).forEach((s, i) => data.set(miniStream.subarray(s * mini, (s + 1) * mini), i * mini));
      out.set(e.name, data.subarray(0, e.size));
    } else out.set(e.name, read(e.start, e.size));
  }
  return out;
}

/** The operand size of a property change (sprm), from its code; null when the first byte says. */
function operandSize(sprm: number): number | null {
  return [1, 1, 2, 4, 2, 2, null, 3][sprm >> 13];
}

/** Property changes in a run of bytes: code → operand (first byte, or a 16-bit word). */
function sprms(data: Uint8Array, from: number, to: number): Map<number, number> {
  const out = new Map<number, number>();
  const v = new DataView(data.buffer, data.byteOffset, data.byteLength);
  for (let o = from; o + 2 <= to; ) {
    const sprm = v.getUint16(o, true);
    o += 2;
    let size = operandSize(sprm);
    if (size === null) size = sprm === 0xd608 ? v.getUint16(o, true) + 1 : data[o] + 1;
    if (o + size > to) break;
    out.set(sprm, size >= 2 && sprm !== 0xd608 ? (size === 2 ? v.getUint16(o, true) : data[o]) : data[o]);
    o += size;
  }
  return out;
}

interface Range {
  start: number;
  end: number;
  props: Map<number, number>;
  istd?: number;
}

export function docToHtml(bytes: Uint8Array): string {
  if (bytes.length < 512 || bytes[0] !== 0xd0 || bytes[1] !== 0xcf) throw new Error("This isn't a Word 97–2003 document (.doc).");
  const files = streams(bytes);
  const word = files.get("WordDocument");
  if (!word) throw new Error("This isn't a Word document (.doc).");
  const w = new DataView(word.buffer, word.byteOffset, word.byteLength);
  if (w.getUint16(0, true) !== 0xa5ec) throw new Error("This isn't a Word document (.doc).");
  if (w.getUint16(2, true) < 0xc1) throw new Error("This is a Word document from before Word 97. Open it in Word and save it as .docx first.");
  const flags = w.getUint16(0x0a, true);
  if (flags & 0x0100) throw new Error("This Word document is protected with a password. Open it in Word, remove the password and save it as .docx first.");
  const table = files.get(flags & 0x0200 ? "1Table" : "0Table");
  if (!table) throw new Error("This Word document is damaged: a part of it is missing.");
  const t = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const fc = (i: number) => [w.getUint32(154 + i * 8, true), w.getUint32(158 + i * 8, true)] as const;
  const ccpText = w.getInt32(76, true);
  const ccpFtn = w.getInt32(80, true);

  // The text, piece by piece: one byte per character (Windows-1252) or two (UTF-16).
  const [clxAt, clxSize] = fc(33);
  let o = clxAt;
  while (o < clxAt + clxSize && table[o] === 0x01) o += 3 + t.getUint16(o + 1, true);
  if (table[o] !== 0x02) throw new Error("This Word document is damaged: its text can't be found.");
  const plcSize = t.getUint32(o + 1, true);
  const plc = o + 5;
  const pieces = (plcSize - 4) / 12;
  const cp1252 = new TextDecoder("windows-1252");
  /** Each character, with where it is in the WordDocument stream (for its formatting). */
  const chars: { ch: string; fc: number }[] = [];
  for (let i = 0; i < pieces; i++) {
    const cpStart = t.getUint32(plc + i * 4, true);
    const cpEnd = t.getUint32(plc + (i + 1) * 4, true);
    const raw = t.getUint32(plc + (pieces + 1) * 4 + i * 8 + 2, true);
    const compressed = (raw & 0x40000000) !== 0;
    const start = compressed ? (raw & 0x3fffffff) / 2 : raw & 0x3fffffff;
    for (let cp = 0; cp < cpEnd - cpStart; cp++) {
      const at = compressed ? start + cp : start + cp * 2;
      if (at + (compressed ? 1 : 2) > word.length) break;
      chars.push({ ch: compressed ? cp1252.decode(word.subarray(at, at + 1)) : String.fromCharCode(w.getUint16(at, true)), fc: at });
    }
  }

  // Formatting pages: runs of characters (CHPX) and paragraphs (PAPX), by stream position.
  const runs = (index: number, paragraphs: boolean): Range[] => {
    const [at, size] = fc(index);
    const n = (size - 4) / 8;
    const out: Range[] = [];
    for (let i = 0; i < n; i++) {
      const page = t.getUint32(at + (n + 1) * 4 + i * 4, true) * 512;
      if (page + 512 > word.length) continue;
      const count = word[page + 511];
      for (let k = 0; k < count; k++) {
        const start = w.getUint32(page + k * 4, true);
        const end = w.getUint32(page + (k + 1) * 4, true);
        const offsetAt = paragraphs ? page + (count + 1) * 4 + k * 13 : page + (count + 1) * 4 + k;
        const offset = word[offsetAt] * 2;
        if (!offset) {
          out.push({ start, end, props: new Map() });
          continue;
        }
        if (paragraphs) {
          let cb = word[page + offset];
          let from = page + offset + 1;
          if (cb === 0) {
            cb = word[from];
            from++;
            cb *= 2;
          } else cb = cb * 2 - 1;
          out.push({ start, end, istd: w.getUint16(from, true), props: sprms(word, from + 2, from + cb) });
        } else {
          const cb = word[page + offset];
          out.push({ start, end, props: sprms(word, page + offset + 1, page + offset + 1 + cb) });
        }
      }
    }
    return out.sort((a, b) => a.start - b.start);
  };
  const charRuns = runs(12, false);
  const paraRuns = runs(13, true);
  const find = (list: Range[], pos: number) => {
    let lo = 0;
    let hi = list.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (list[mid].end <= pos) lo = mid + 1;
      else if (list[mid].start > pos) hi = mid - 1;
      else return list[mid];
    }
    return null;
  };

  // Style names, by number: "heading 1", "Title"…
  const styleNames = new Map<number, string>();
  {
    const [at, size] = fc(1);
    if (size > 4) {
      const cbStshi = t.getUint16(at, true);
      const cstd = t.getUint16(at + 2, true);
      const cbBase = t.getUint16(at + 4, true);
      let p = at + 2 + cbStshi;
      for (let istd = 0; istd < cstd && p + 2 <= at + size; istd++) {
        const cb = t.getUint16(p, true);
        if (cb >= cbBase + 2) {
          const nameAt = p + 2 + cbBase;
          const len = t.getUint16(nameAt, true);
          // Word's own number for a built-in style is the same in every language ("Overskrift 1" is heading 1).
          const sti = t.getUint16(p + 2, true) & 0x0fff;
          const builtIn = sti >= 1 && sti <= 9 ? `heading ${sti}` : sti === 62 ? "title" : null;
          styleNames.set(istd, builtIn ?? String.fromCharCode(...Array.from({ length: Math.min(len, 255) }, (_, i) => t.getUint16(nameAt + 2 + i * 2, true))));
        }
        p += 2 + cb;
      }
    }
  }

  // Lists: which are numbered (and which bulleted), by list number and level.
  const numbered = new Map<number, boolean[]>();
  try {
    const [lstAt, lstSize] = fc(73);
    const [lfoAt, lfoSize] = fc(74);
    if (lstSize && lfoSize) {
      const count = t.getInt16(lstAt, true);
      const lists: { id: number; levels: boolean[] }[] = [];
      let p = lstAt + 2 + count * 28;
      for (let i = 0; i < count; i++) {
        const lstf = lstAt + 2 + i * 28;
        const simple = (table[lstf + 26] & 1) !== 0;
        const levels: boolean[] = [];
        for (let l = 0; l < (simple ? 1 : 9); l++) {
          levels.push(table[p + 4] !== 0x17);
          const papx = table[p + 25];
          const chpx = table[p + 24];
          p += 28 + papx + chpx;
          p += 2 + t.getUint16(p, true) * 2;
        }
        lists.push({ id: t.getInt32(lstf, true), levels });
      }
      const lfos = t.getUint32(lfoAt, true);
      for (let i = 0; i < lfos; i++) {
        const id = t.getInt32(lfoAt + 4 + i * 16, true);
        numbered.set(i + 1, lists.find((l) => l.id === id)?.levels ?? []);
      }
    }
  } catch {
    // Lists then show as bullets.
  }

  // Footnotes: their text follows the document's, one stretch per note.
  const footnoteText: string[] = [];
  {
    const [at, size] = fc(3);
    const n = size >= 8 ? (size - 4) / 4 - 1 : 0;
    for (let i = 0; i < n; i++) {
      const from = ccpText + t.getUint32(at + i * 4, true);
      const to = ccpText + t.getUint32(at + (i + 1) * 4, true);
      footnoteText.push(chars.slice(from, Math.min(to, ccpText + ccpFtn)).map((c) => c.ch).join("").replace(/[\x00-\x08\x0b-\x1f]/g, " ").replace(/\r/g, " ").trim());
    }
  }

  // The document's text, paragraph by paragraph.
  const out: string[] = [];
  const notes: string[] = [];
  let para = "";
  let rows: string[][] = [];
  let row: string[] = [];
  let cell: string[] = [];
  const fields: { inResult: boolean }[] = [];
  const flushTable = () => {
    if (!rows.length) return;
    out.push(`<table>${rows.map((r) => `<tr>${r.map((c) => `<td>${c || "<p></p>"}</td>`).join("")}</tr>`).join("")}</table>`);
    rows = [];
  };
  const paragraphHtml = (pos: number) => {
    const p = find(paraRuns, pos);
    const props = p?.props ?? new Map<number, number>();
    const name = (styleNames.get(p?.istd ?? 0) ?? "").toLowerCase();
    const outline = props.get(0x2640);
    const level = Number(/^heading (\d)/.exec(name)?.[1]) || (outline !== undefined && outline < 9 ? outline + 1 : 0);
    const jc = props.get(0x2461) ?? props.get(0x2403) ?? 0;
    const css = jc === 1 ? ' style="text-align:center"' : jc === 2 ? ' style="text-align:right"' : jc === 3 ? ' style="text-align:justify"' : "";
    const ilfo = props.get(0x460b) ?? 0;
    if (ilfo > 0 && ilfo < 0x7ff) {
      const ilvl = Math.min(8, props.get(0x260a) ?? 0);
      return `<p data-list="${numbered.get(ilfo)?.[ilvl] ? "number" : "bullet"}" data-level="${ilvl}"${css}>${para}</p>`;
    }
    if (name === "title") return `<p class="MsoTitle"${css}>${para}</p>`;
    if (level) return `<h${Math.min(level, 6)}${css}>${para}</h${Math.min(level, 6)}>`;
    return `<p${css}>${para}</p>`;
  };
  const inTable = (pos: number) => {
    const props = find(paraRuns, pos)?.props;
    return !!props && ((props.get(0x2416) ?? 0) > 0 || (props.get(0x6649) ?? 0) > 0);
  };
  const isRowEnd = (pos: number) => {
    const props = find(paraRuns, pos)?.props;
    return !!props && ((props.get(0x2417) ?? 0) > 0 || (props.get(0x244b) ?? 0) > 0);
  };

  let footnote = 0;
  for (let i = 0; i < Math.min(ccpText, chars.length); i++) {
    const { ch, fc: pos } = chars[i];
    const code = ch.charCodeAt(0);
    // Fields: only their result shows (a page number, a citation), not their instruction.
    if (code === 0x13) {
      fields.push({ inResult: false });
      continue;
    }
    if (code === 0x14) {
      if (fields.length) fields[fields.length - 1].inResult = true;
      continue;
    }
    if (code === 0x15) {
      fields.pop();
      continue;
    }
    if (fields.some((f) => !f.inResult)) continue;
    if (code === 0x0d || code === 0x07 || code === 0x0c) {
      const html = paragraphHtml(pos);
      if (code === 0x07 && isRowEnd(pos)) {
        rows.push(row);
        row = [];
        para = "";
        continue;
      }
      if (inTable(pos)) {
        cell.push(html);
        if (code === 0x07) {
          row.push(cell.join(""));
          cell = [];
        }
      } else {
        flushTable();
        out.push(html);
      }
      para = "";
      continue;
    }
    if (code === 0x02) {
      // shortcut: a footnote becomes [n] and a note at the end; real footnotes when the importer can make them.
      notes.push(footnoteText[footnote++] ?? "");
      para += `<sup>[${notes.length}]</sup>`;
      continue;
    }
    if (code === 0x0b) {
      para += "<br>";
      continue;
    }
    if (code === 0x09) {
      para += "\t";
      continue;
    }
    if (code === 0x1e) {
      para += "‑";
      continue;
    }
    if (code < 0x20) continue;
    const props = find(charRuns, pos)?.props;
    const on = (sprm: number) => {
      const v = props?.get(sprm);
      return v === 1 || v === 0x81;
    };
    let html = esc(ch);
    if (props) {
      if (on(0x0835)) html = `<b>${html}</b>`;
      if (on(0x0836)) html = `<i>${html}</i>`;
      if ((props.get(0x2a3e) ?? 0) > 0) html = `<u>${html}</u>`;
      if (on(0x0837)) html = `<s>${html}</s>`;
      const iss = props.get(0x2a48);
      if (iss === 1) html = `<sup>${html}</sup>`;
      else if (iss === 2) html = `<sub>${html}</sub>`;
    }
    para += html;
  }
  if (para) out.push(`<p>${para}</p>`);
  flushTable();
  let html = out.join("").replace(/<\/(b|i|u|s|sup|sub)><\1>/g, "");
  if (notes.length) html += `<h2>Notes</h2>${notes.map((n, i) => `<p>[${i + 1}] ${esc(n)}</p>`).join("")}`;
  return html;
}
