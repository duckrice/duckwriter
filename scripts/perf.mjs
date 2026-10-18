// How DuckWriter copes with a long document: makes one (600 by default, about
// 250 pages, with headings, lists, tables and footnotes), opens it off-screen,
// and measures opening, typing (a key press until the next frame) and saving.
//   node scripts/perf.mjs [size]
import { spawn } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import * as docx from "docx";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exe = join(root, "src-tauri", "target", "release", "duckwriter.exe");
const pages = Number(process.argv[2] ?? 600);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const { Paragraph, TextRun, HeadingLevel } = docx;

// About 5 paragraphs of 70 words per page.
const words = "the rate of the reaction increases with temperature as the measurements in the second series show clearly and the uncertainty stays small across all runs".split(" ");
const sentence = (i) => Array.from({ length: 70 }, (_, k) => words[(i * 7 + k) % words.length]).join(" ") + ".";
const children = [];
const footnotes = {};
let note = 1;
for (let p = 0; p < pages; p++) {
  if (p % 4 === 0) children.push(new Paragraph({ text: `Section ${p / 4 + 1}`, heading: HeadingLevel.HEADING_1 }));
  for (let i = 0; i < 4; i++) {
    const runs = [new TextRun(sentence(p * 5 + i))];
    if (i === 1 && p % 2 === 0) {
      footnotes[note] = { children: [new Paragraph(`Note ${note} about the measurements.`)] };
      runs.push(new docx.FootnoteReferenceRun(note++));
    }
    children.push(new Paragraph({ children: runs }));
  }
  if (p % 10 === 5) {
    children.push(new docx.Table({ rows: Array.from({ length: 6 }, (_, r) => new docx.TableRow({ children: [String(r), String(r * 2), String(r * 3)].map((t) => new docx.TableCell({ children: [new Paragraph(t)] })) })) }));
  }
  if (p % 10 === 7) for (let i = 0; i < 3; i++) children.push(new Paragraph({ text: `Point ${i + 1} of the list`, bullet: { level: 0 } }));
}
const doc = new docx.Document({ footnotes, styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ children }] });
const work = mkdtempSync(join(tmpdir(), "duckwriter-perf-"));
const file = join(work, "long.docx");
writeFileSync(file, await docx.Packer.toBuffer(doc));

const port = 9620;
const child = spawn(exe, [file], {
  env: { ...process.env, DUCKWRITER_TEST_OFFSCREEN: "1", WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${port} --disable-features=CalculateNativeWinOcclusion --disable-backgrounding-occluded-windows --disable-renderer-backgrounding`, WEBVIEW2_USER_DATA_FOLDER: join(work, "profile"), DUCKWRITER_RECOVERY_DIR: join(work, "Recovery") },
  stdio: "ignore",
});
child.on("exit", (code) => console.log("app exited", code));
const started = Date.now();
let target;
for (let i = 0; i < 120 && !target; i++) {
  await sleep(500);
  try {
    target = (await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find((t) => t.type === "page");
  } catch {}
}
const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((r) => (ws.onopen = r));
let id = 1;
const pending = new Map();
ws.onmessage = (e) => {
  const m = JSON.parse(e.data);
  if (m.id && pending.has(m.id)) (pending.get(m.id)(m), pending.delete(m.id));
};
const send = (method, params = {}) => new Promise((r) => { const i = id++; pending.set(i, r); ws.send(JSON.stringify({ id: i, method, params })); });
const ev = async (expression) => (await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true })).result?.result?.value;
for (let i = 0; i < 240; i++) {
  await sleep(250);
  if (await ev(`document.title.startsWith("long.docx") && document.querySelectorAll(".page-sheet").length > 10`)) break;
}
const opened = Date.now() - started;
await ev(`localStorage.setItem("duckwriter.debug", "1"); localStorage.setItem("duckwriter.autosave", "off")`);
await ev(`location.reload()`);
for (let i = 0; i < 240; i++) {
  await sleep(250);
  if (await ev(`!!window.__dw && document.querySelectorAll(".page-sheet").length > 10`)) break;
}
await sleep(3000);
const info = await ev(`({ pages: document.querySelectorAll(".page-sheet").length, paragraphs: window.__dw.body.state.doc.childCount })`);
// Typing: each key press until the next frame after the layout.
const typing = await ev(`(async () => {
  const v = window.__dw.body;
  const mid = Math.floor(v.state.doc.content.size / 2);
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(mid))));
  const times = [];
  for (let i = 0; i < 20; i++) {
    const t = performance.now();
    v.dispatch(v.state.tr.insertText("x"));
    await new Promise((r) => requestAnimationFrame(() => setTimeout(r, 0)));
    times.push(performance.now() - t);
  }
  times.sort((a, b) => a - b);
  return { median: Math.round(times[10]), worst: Math.round(times[19]) };
})()`);
const saving = await ev(`(async () => { const t = performance.now(); await window.__dw.build(); return Math.round(performance.now() - t); })()`);
console.log(JSON.stringify({ pagesAsked: pages, ...info, openMs: opened, typingMs: typing, buildMs: saving }));
ws.close();
child.kill();
await sleep(800);
try {
  rmSync(work, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
} catch {}
