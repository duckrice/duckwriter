// Rich Text Format (.rtf): opened as a new Word document, as HTML the
// importer reads. Text, bold, italic, underline, strike-through, raised and
// lowered text, headings (by their style's name), alignment, lists, tables and
// footnotes; fonts, sizes, colours, pictures and headers are left out.

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

/** Groups whose text isn't the document's (font lists, pictures, headers…). */
const SKIP = new Set(["fonttbl", "colortbl", "info", "listtable", "listoverridetable", "generator", "pict", "object", "header", "headerl", "headerr", "headerf", "footer", "footerl", "footerr", "footerf", "fldinst", "xmlnstbl", "rsidtbl", "themedata", "colorschememapping", "latentstyles", "datastore", "pgdsctbl", "ftnsep", "ftnsepc", "aftnsep", "aftnsepc", "pntxta", "pntxtb", "bkmkstart", "bkmkend", "revtbl", "filetbl", "mmathPr", "userprops", "nonshppict", "shp"]);

interface Run {
  b: boolean;
  i: boolean;
  u: boolean;
  s: boolean;
  v: "" | "sup" | "sub";
  skip: boolean;
  /** Inside the stylesheet, or a footnote. */
  dest: "" | "stylesheet" | "footnote" | "listtext";
  uc: number;
}

export function rtfToHtml(text: string): string {
  if (!text.startsWith("{\\rtf")) throw new Error("This isn't a Rich Text Format (.rtf) file.");
  const decoder = new TextDecoder("windows-1252");
  const styleNames = new Map<number, string>();
  let styleEntry: { n: number; name: string } | null = null;

  let run: Run = { b: false, i: false, u: false, s: false, v: "", skip: false, dest: "", uc: 1 };
  const stack: Run[] = [];
  // The paragraph being read: its text (as HTML), style, alignment, list, and whether it is in a table.
  let para = "";
  let style = 0;
  let align = "";
  let list = false;
  let bulletText = "";
  let inTable = false;
  let cell: string[] = [];
  let row: string[] = [];
  const rows: string[][] = [];
  const out: string[] = [];
  const notes: string[] = [];
  let note = "";
  let skipChars = 0;

  const add = (t: string) => {
    if (!t) return;
    if (skipChars > 0) {
      const n = Math.min(skipChars, t.length);
      skipChars -= n;
      t = t.slice(n);
      if (!t) return;
    }
    if (run.skip) return;
    if (run.dest === "stylesheet") {
      if (styleEntry) styleEntry.name += t;
      return;
    }
    if (run.dest === "listtext") {
      bulletText += t;
      return;
    }
    let html = esc(t);
    if (run.b) html = `<b>${html}</b>`;
    if (run.i) html = `<i>${html}</i>`;
    if (run.u) html = `<u>${html}</u>`;
    if (run.s) html = `<s>${html}</s>`;
    if (run.v) html = `<${run.v}>${html}</${run.v}>`;
    if (run.dest === "footnote") note += html;
    else para += html;
  };
  const endTable = () => {
    if (!rows.length) return;
    out.push(`<table>${rows.map((r) => `<tr>${r.map((c) => `<td>${c || "<p></p>"}</td>`).join("")}</tr>`).join("")}</table>`);
    rows.length = 0;
  };
  /** The paragraph as HTML, by its style's name ("heading 1", "Title") and its list. */
  const paragraphHtml = () => {
    const name = (styleNames.get(style) ?? "").toLowerCase();
    const level = Number(/^heading (\d)/.exec(name)?.[1]) || 0;
    const css = align ? ` style="text-align:${align}"` : "";
    if (list) {
      const numbered = /\d/.test(bulletText);
      return `<p data-list="${numbered ? "number" : "bullet"}" data-level="0"${css}>${para}</p>`;
    }
    if (name === "title") return `<p class="MsoTitle"${css}>${para}</p>`;
    if (level) return `<h${level}${css}>${para}</h${level}>`;
    return `<p${css}>${para}</p>`;
  };
  const endParagraph = () => {
    if (run.dest === "footnote") {
      note += " ";
      return;
    }
    const html = paragraphHtml();
    if (inTable) cell.push(html);
    else {
      endTable();
      out.push(html);
    }
    para = "";
  };

  // Control words, symbols, groups and text, in one pass.
  const re = /\\([a-zA-Z]+)(-?\d+)? ?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z])|([{}])|([^\\{}\r\n]+)|[\r\n]+/g;
  let m: RegExpExecArray | null;
  let groupStart = false;
  while ((m = re.exec(text))) {
    const [, word, num, hex, symbol, brace, plain] = m;
    const first = groupStart;
    groupStart = false;
    if (brace === "{") {
      stack.push(run);
      run = { ...run };
      groupStart = true;
      continue;
    }
    if (brace === "}") {
      if (run.dest === "stylesheet" && styleEntry && stack.at(-1)?.dest === "stylesheet") {
        styleNames.set(styleEntry.n, styleEntry.name.replace(/;\s*$/, "").trim());
        styleEntry = null;
      }
      const leavingNote = run.dest === "footnote" && stack.at(-1)?.dest !== "footnote";
      run = stack.pop() ?? run;
      if (leavingNote) {
        notes.push(note.trim());
        note = "";
        para += `<sup>[${notes.length}]</sup>`;
      }
      continue;
    }
    if (plain !== undefined) {
      add(plain);
      continue;
    }
    if (hex !== undefined) {
      add(decoder.decode(new Uint8Array([parseInt(hex, 16)])));
      continue;
    }
    if (symbol !== undefined) {
      if (symbol === "*") {
        // An optional destination: skipped unless it is one this reader knows.
        run.skip = true;
        continue;
      }
      if (symbol === "~") add(" ");
      else if (symbol === "_") add("‑");
      else if (symbol === "\\" || symbol === "{" || symbol === "}") add(symbol);
      continue;
    }
    if (word === undefined) continue;
    const n = num === undefined ? null : Number(num);
    if (first && SKIP.has(word)) {
      run.skip = true;
      continue;
    }
    switch (word) {
      case "stylesheet":
        run.dest = "stylesheet";
        run.skip = false;
        break;
      case "footnote":
        run.dest = "footnote";
        run.skip = false;
        break;
      case "listtext":
      case "pntext":
        // The bullet or number Word wrote before a list item: tells which kind of list.
        run.dest = "listtext";
        list = true;
        break;
      case "s":
        if (run.dest === "stylesheet") styleEntry = { n: n ?? 0, name: "" };
        else style = n ?? 0;
        break;
      case "b":
        run.b = n !== 0;
        break;
      case "i":
        run.i = n !== 0;
        break;
      case "ul":
        run.u = n !== 0;
        break;
      case "ulnone":
        run.u = false;
        break;
      case "strike":
        run.s = n !== 0;
        break;
      case "super":
        run.v = "sup";
        break;
      case "sub":
        run.v = "sub";
        break;
      case "nosupersub":
        run.v = "";
        break;
      case "plain":
        Object.assign(run, { b: false, i: false, u: false, s: false, v: "" });
        break;
      case "pard":
        style = 0;
        align = "";
        list = false;
        bulletText = "";
        inTable = false;
        break;
      case "qc":
        align = "center";
        break;
      case "qr":
        align = "right";
        break;
      case "qj":
        align = "justify";
        break;
      case "ls":
        list = true;
        break;
      case "intbl":
        inTable = true;
        break;
      case "par":
        endParagraph();
        break;
      case "cell":
        if (para || !cell.length) endParagraph();
        row.push(cell.join(""));
        cell = [];
        break;
      case "row":
        rows.push(row);
        row = [];
        break;
      case "line":
        if (!run.skip && run.dest === "") para += "<br>";
        break;
      case "tab":
        add("\t");
        break;
      case "emdash":
        add("—");
        break;
      case "endash":
        add("–");
        break;
      case "bullet":
        add("•");
        break;
      case "lquote":
        add("‘");
        break;
      case "rquote":
        add("’");
        break;
      case "ldblquote":
        add("“");
        break;
      case "rdblquote":
        add("”");
        break;
      case "uc":
        run.uc = n ?? 1;
        break;
      case "u":
        add(String.fromCharCode(n === null ? 63 : n < 0 ? n + 65536 : n));
        // The plain-text stand-in that follows is not read twice.
        skipChars = run.uc;
        break;
    }
  }
  if (para.trim()) endParagraph();
  endTable();
  let html = out.join("");
  // shortcut: footnotes become [n] and notes at the end; real footnotes when the importer can make them.
  if (notes.length) html += `<h2>Notes</h2>${notes.map((t, i) => `<p>[${i + 1}] ${t}</p>`).join("")}`;
  return html;
}
