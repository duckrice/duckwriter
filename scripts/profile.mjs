// Profiles opening a long document, and typing in it: which functions take
// the time, and from where. If it hangs, "stacks" shows where it is stuck.
//   node scripts/profile.mjs [size] [seconds to wait for opening] [stacks]
// With REVIEW=1, every paragraph has tracked changes and every third a comment.
import { spawn } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import * as docx from "docx";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exe = join(root, "src-tauri", "target", "release", "duckwriter.exe");
const pages = Number(process.argv[2] ?? 100);
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const { Paragraph, TextRun, HeadingLevel } = docx;
const words = "the rate of the reaction increases with temperature as the measurements in the second series show clearly and the uncertainty stays small across all runs".split(" ");
const sentence = (i) => Array.from({ length: 70 }, (_, k) => words[(i * 7 + k) % words.length]).join(" ") + ".";
const children = [];
const footnotes = {};
const comments = [];
let note = 1;
for (let p = 0; p < pages; p++) {
  if (p % 4 === 0) children.push(new Paragraph({ text: `Section ${p / 4 + 1}`, heading: HeadingLevel.HEADING_1 }));
  for (let i = 0; i < 4; i++) {
    const runs = [new TextRun(sentence(p * 5 + i))];
    const k = p * 5 + i;
    if (process.env.REVIEW) {
      const who = { author: k % 2 ? "Kari" : "Ola", date: "2026-10-01T10:00:00Z" };
      runs.push(new docx.InsertedTextRun({ text: " added here", id: 10000 + k * 2, ...who }), new docx.DeletedTextRun({ text: " taken out", id: 10001 + k * 2, ...who }));
      if (k % 3 === 0) {
        comments.push({ id: k, author: who.author, date: new Date("2026-10-01"), children: [new Paragraph(`Comment ${k}`)] });
        runs.unshift(new docx.CommentRangeStart(k));
        runs.push(new docx.CommentRangeEnd(k), new TextRun({ children: [new docx.CommentReference(k)] }));
      }
    }
    if (i === 1 && p % 2 === 0) {
      footnotes[note] = { children: [new Paragraph(`Note ${note}.`)] };
      runs.push(new docx.FootnoteReferenceRun(note++));
    }
    children.push(new Paragraph({ children: runs }));
  }
  if (p % 10 === 7) for (let i = 0; i < 3; i++) children.push(new Paragraph({ text: `Point ${i + 1}`, bullet: { level: 0 } }));
}
const work = mkdtempSync(join(tmpdir(), "duckwriter-perf-"));
const file = join(work, "long.docx");
writeFileSync(file, await docx.Packer.toBuffer(new docx.Document({ footnotes, ...(comments.length ? { comments: { children: comments } } : {}), styles: { default: { document: { run: { font: "Calibri", size: 22 } } } }, sections: [{ children }] })));

const port = 9621;
const child = spawn(exe, [], {
  env: { ...process.env, DUCKWRITER_TEST_OFFSCREEN: "1", WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${port} --disable-features=CalculateNativeWinOcclusion --disable-backgrounding-occluded-windows --disable-renderer-backgrounding`, WEBVIEW2_USER_DATA_FOLDER: join(work, "profile"), DUCKWRITER_RECOVERY_DIR: join(work, "Recovery") },
  stdio: "ignore",
});
let target;
for (let i = 0; i < 240 && !target; i++) {
  await sleep(500);
  try {
    target = (await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find((t) => t.type === "page");
  } catch {}
}
const ws = new WebSocket(target.webSocketDebuggerUrl);
await new Promise((r) => (ws.onopen = r));
let id = 1;
const pending = new Map();
let onPaused = null;
ws.onmessage = (e) => {
  const m = JSON.parse(e.data);
  if (m.method === "Debugger.paused" && onPaused) onPaused(m.params);
  if (m.id && pending.has(m.id)) (pending.get(m.id)(m), pending.delete(m.id));
};
const send = (method, params = {}) => new Promise((r) => { const i = id++; pending.set(i, r); ws.send(JSON.stringify({ id: i, method, params })); });
const ev = async (expression) => (await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true })).result?.result?.value;
await sleep(2000);
await ev(`localStorage.setItem("duckwriter.debug", "1"); localStorage.setItem("duckwriter.autosave", "off")`);
await ev(`location.reload()`);
for (let i = 0; i < 60; i++) {
  await sleep(250);
  if (await ev(`!!window.__dw`)) break;
}

/** Self time per function, the largest first. */
function hotspots(profile, top = 18) {
  const byId = new Map(profile.nodes.map((n) => [n.id, n]));
  const self = new Map();
  const dt = profile.timeDeltas;
  profile.samples.forEach((sid, i) => {
    const n = byId.get(sid);
    const f = n.callFrame;
    const key = `${f.functionName || "(anonymous)"} ${f.url.split("/").pop()}:${f.lineNumber + 1}`;
    self.set(key, (self.get(key) ?? 0) + (dt[i] ?? 0) / 1000);
  });
  return [...self].sort((a, b) => b[1] - a[1]).slice(0, top).map(([k, ms]) => `${Math.round(ms).toString().padStart(7)} ms  ${k}`);
}

/** Where the time of the hottest functions is spent from: their callers, a few levels up. */
function callers(profile, names, depth = 5) {
  const byId = new Map(profile.nodes.map((n) => [n.id, n]));
  const parent = new Map();
  for (const n of profile.nodes) for (const c of n.children ?? []) parent.set(c, n.id);
  const name = (n) => `${n.callFrame.functionName || "(anon)"}:${n.callFrame.lineNumber + 1}:${n.callFrame.columnNumber}`;
  const out = new Map();
  profile.samples.forEach((sid, i) => {
    let n = byId.get(sid);
    if (!names.includes(n.callFrame.functionName)) return;
    const chain = [];
    for (let d = 0; d < depth && n; d++) {
      chain.push(name(n));
      n = byId.get(parent.get(n.id));
    }
    const key = chain.join(" < ");
    out.set(key, (out.get(key) ?? 0) + (profile.timeDeltas[i] ?? 0) / 1000);
  });
  return [...out].sort((a, b) => b[1] - a[1]).slice(0, 12).map(([k, ms]) => `${Math.round(ms).toString().padStart(6)} ms  ${k}`);
}

await send("Profiler.enable");
if (process.argv[4] === "stacks") await send("Debugger.enable");
await send("Profiler.setSamplingInterval", { interval: 500 });
await send("Profiler.start");
const t0 = Date.now();
const log = (m) => console.log(`${((Date.now() - t0) / 1000).toFixed(1)} s  ${m}`);
const within = (p, ms) => Promise.race([p, sleep(ms).then(() => "busy")]);
ev(`setTimeout(() => window.__dw.openPath(${JSON.stringify(file)}), 0)`);
const wait = Number(process.argv[3] ?? 40) * 1000;
while (Date.now() - t0 < wait) {
  await sleep(2000);
  const r = await within(ev(`document.title + " | " + document.querySelectorAll(".page-sheet").length`), 3000);
  log(r);
  if (r === "busy" && process.argv[4] === "stacks") {
    const paused = new Promise((res) => (onPaused = res));
    ws.send(JSON.stringify({ id: id++, method: "Debugger.pause" }));
    const p = await paused;
    console.log(p.callFrames.slice(0, 14).map((c) => `    ${c.functionName || "(anon)"} @${c.location.lineNumber + 1}:${c.location.columnNumber}`).join("\n"));
    const stats = await send("Debugger.evaluateOnCallFrame", { callFrameId: p.callFrames[0].callFrameId, expression: `JSON.stringify(globalThis.__layoutStats)` });
    console.log("    stats", stats.result?.result?.value);
    ws.send(JSON.stringify({ id: id++, method: "Debugger.resume" }));
  }
  if (typeof r === "string" && r.startsWith("long.docx") && !r.endsWith("| 1")) { await sleep(1500); break; }
}
if ((await within(ev("1"), 3000)) === "busy") {
  console.log("Still busy; giving up.");
  child.kill();
  process.exit(1);
}
const openMs = Date.now() - t0;
const { profile } = (await send("Profiler.stop")).result;
const info = await ev(`({ pages: document.querySelectorAll(".page-sheet").length, paragraphs: window.__dw.body.state.doc.childCount })`);
console.log(`Opened ${info.pages} pages (${info.paragraphs} paragraphs) in ${openMs} ms; hot spots:`);
console.log(hotspots(profile).join("\n"));
console.log("Layout:", JSON.stringify(await ev(`Object.fromEntries(Object.entries(globalThis.__layoutStats ?? {}).map(([k, v]) => [k, Math.round(v)]))`)));

await send("Profiler.start");
const typing = await ev(`(async () => {
  const v = window.__dw.body;
  const mid = Math.floor(v.state.doc.content.size / 2);
  v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.near(v.state.doc.resolve(mid))));
  const times = [];
  for (let i = 0; i < 10; i++) {
    const t = performance.now();
    v.dispatch(v.state.tr.insertText("x"));
    await new Promise((r) => requestAnimationFrame(() => setTimeout(r, 0)));
    times.push(performance.now() - t);
  }
  times.sort((a, b) => a - b);
  // Then until the pages are laid out again.
  const runs = globalThis.__layoutStats?.runs ?? 0;
  const t = performance.now();
  while ((globalThis.__layoutStats?.runs ?? 0) === runs && performance.now() - t < 3000) await new Promise((r) => setTimeout(r, 10));
  return { median: Math.round(times[5]), worst: Math.round(times[9]), relaidAfter: Math.round(performance.now() - t) };
})()`);
const typed = (await send("Profiler.stop")).result.profile;
console.log(`\nTyping: ${JSON.stringify(typing)} ms per key; hot spots:`);
console.log(hotspots(typed).join("\n"));
console.log("\nCalled from:");
console.log(callers(typed, ["nodesBetween", "replaceChildren", "getClientRects", "removeChild", "remove", "descendants"], 7).join("\n"));
ws.close();
child.kill();
await sleep(800);
try {
  rmSync(work, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
} catch {}
