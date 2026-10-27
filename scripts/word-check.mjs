// Compares DuckWriter's pages with real Word's, page by page. Word builds test documents (in
// scripts/word-docs) full of numbered paragraphs (P01, P02, …) and saves each with its own PDF;
// DuckWriter opens the same file and lays it out. Then for every page: which paragraphs are on
// it, its header and footer, and whether it is landscape must be the same.
//
//   npx tauri build --no-bundle && npm run word-check
//
// Needs Microsoft Word on this computer (it is driven through COM, in the background).
import { execFileSync, spawn } from "node:child_process";
import { mkdtempSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exe = join(root, "src-tauri", "target", "release", "duckwriter.exe");
const work = mkdtempSync(join(tmpdir(), "duckwriter-word-check-"));
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** For each page of a PDF: the paragraphs on it, its header and footer text, and whether it's landscape. */
async function pagesOf(file) {
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(readFileSync(file)), standardFontDataUrl: join(root, "node_modules", "pdfjs-dist", "standard_fonts") + "/" }).promise;
  const pages = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const items = (await page.getTextContent()).items.filter((i) => i.str.trim());
    const height = page.view[3];
    const ids = [...new Set(items.map((i) => /P\d\d/.exec(i.str)?.[0]).filter(Boolean))].join(",");
    const footer = items.filter((i) => i.transform[5] < height * 0.08).map((i) => i.str.trim()).join(" ");
    const header = items.filter((i) => i.transform[5] > height * 0.92).map((i) => i.str.trim()).join(" ");
    pages.push({ ids, header, footer, landscape: page.view[2] > page.view[3] });
  }
  return pages;
}

/** DuckWriter's pages of a .docx, as the built app lays them out. */
async function duckwriterPages(docx, port) {
  const env = {
    ...process.env,
    DUCKWRITER_TEST_OFFSCREEN: "1",
    WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${port} --disable-features=CalculateNativeWinOcclusion --disable-backgrounding-occluded-windows --disable-renderer-backgrounding`,
    WEBVIEW2_USER_DATA_FOLDER: join(work, `profile-${port}`),
    DUCKWRITER_RECOVERY_DIR: join(work, "recovery"),
  };
  const child = spawn(exe, [docx], { env, stdio: "ignore" });
  try {
    let target;
    for (let i = 0; i < 60 && !target; i++) {
      await sleep(500);
      try {
        target = (await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find((t) => t.type === "page");
      } catch {
        // Not up yet.
      }
    }
    if (!target) throw new Error("DuckWriter did not start");
    const ws = new WebSocket(target.webSocketDebuggerUrl);
    await new Promise((resolve) => (ws.onopen = resolve));
    let id = 0;
    const pending = new Map();
    ws.onmessage = (e) => {
      const m = JSON.parse(e.data);
      if (pending.has(m.id)) (pending.get(m.id)(m), pending.delete(m.id));
    };
    const evaluate = (expression) =>
      new Promise((resolve) => {
        const n = ++id;
        pending.set(n, (m) => resolve(m.result?.result?.value));
        ws.send(JSON.stringify({ id: n, method: "Runtime.evaluate", params: { expression, awaitPromise: true, returnByValue: true } }));
      });
    const waitFor = async (expression, ms = 15000) => {
      for (const end = Date.now() + ms; Date.now() < end; await sleep(200)) if (await evaluate(expression)) return true;
      return false;
    };
    // The app's own page first (not the blank one before it), then its test handle.
    await waitFor(`!!document.querySelector(".sheet .doc")`, 30000);
    await evaluate(`localStorage.setItem("duckwriter.debug", "1")`);
    await evaluate("location.reload()");
    await sleep(1500);
    if (!(await waitFor(`!!window.__dw && document.querySelectorAll(".page-sheet").length > 0`, 45000))) throw new Error(`DuckWriter didn't open the document: ${await evaluate(`[document.title, !!window.__dw, document.querySelectorAll(".page-sheet").length, document.querySelector("dialog[open]")?.textContent ?? ""].join(" / ")`)}`);
    // The pages settle once fonts are in.
    await sleep(1500);
    // The pages as laid out on screen: which paragraph starts on each (by its first line), its header and footer.
    // (On screen, not in a saved PDF: the PDF has one paper size for every page, so sideways pages aren't.)
    const pages = await evaluate(`(() => {
      const sheets = [...document.querySelectorAll(".page-sheet")].map((s) => s.getBoundingClientRect());
      const pages = sheets.map((r) => ({ ids: [], header: "", footer: "", landscape: r.width > r.height }));
      for (const p of document.querySelectorAll(".doc p")) {
        const id = /^P\\d\\d/.exec(p.textContent)?.[0];
        if (!id) continue;
        const range = document.createRange();
        range.selectNodeContents(p);
        const top = [...range.getClientRects()].find((r) => r.height > 0)?.top;
        const k = sheets.findIndex((s) => top >= s.top - 1 && top < s.bottom);
        if (k >= 0 && !pages[k].ids.includes(id)) pages[k].ids.push(id);
      }
      for (const hf of document.querySelectorAll(".hf")) {
        const k = Number(hf.dataset.page);
        if (pages[k]) pages[k][hf.dataset.kind] = hf.textContent.replace(/\\s+/g, " ").trim();
      }
      return pages.map((p) => ({ ...p, ids: p.ids.join(",") }));
    })()`);
    ws.close();
    return pages;
  } finally {
    child.kill();
  }
}

let differences = 0;
let port = 9700;
for (const script of readdirSync(join(root, "scripts", "word-docs")).filter((f) => f.endsWith(".ps1"))) {
  const name = script.replace(/\.ps1$/, "");
  const docx = join(work, `${name}.docx`);
  execFileSync("powershell", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", join(root, "scripts", "word-docs", script), "-Out", docx], { stdio: "ignore", timeout: 180000 });
  const [word, duck] = [await pagesOf(docx.replace(/\.docx$/, ".pdf")), await duckwriterPages(docx, port++)];
  console.log(`\n${name}: Word ${word.length} pages, DuckWriter ${duck.length}`);
  for (let i = 0; i < Math.max(word.length, duck.length); i++) {
    const [w, d] = [word[i], duck[i]];
    const same = !!w && !!d && w.ids === d.ids && w.header === d.header && w.footer === d.footer && w.landscape === d.landscape;
    if (!same) differences++;
    const show = (p) => (p ? `${p.ids || "-"} | ${p.header || "-"} | ${p.footer || "-"}${p.landscape ? " | landscape" : ""}` : "(no page)");
    console.log(`${same ? "same" : "DIFF"}  page ${i + 1}: ${show(w)}${same ? "" : `\n      DuckWriter: ${show(d)}`}`);
  }
}
try {
  rmSync(work, { recursive: true, force: true, maxRetries: 20, retryDelay: 250 });
} catch {
  // WebView2 may still hold a profile for a moment; the temp folder is cleaned up by Windows later.
}
console.log(differences ? `\n${differences} page(s) differ from Word` : "\nEvery page as in Word");
process.exit(differences ? 1 : 0);
