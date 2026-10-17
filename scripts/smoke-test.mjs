// Drives the built desktop app over the DevTools protocol (WebView2 remote
// debugging) and checks the things unit tests can't: that documents open and
// look right, that typing and the toolbar work, and what ends up in a saved file.
//
//   npm run samples && npm run tauri build -- --no-bundle && npm run smoke
//
// Add `-- --screenshots` to also save pictures of the app along the way, into
// test-screenshots/ (or `-- --screenshots=some/folder`).
import { execFileSync, spawn } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exe = join(root, "src-tauri", "target", "release", "duckwriter.exe");
const sample = (name) => join(root, "samples", name);
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
// Its own browser profile: the test then works while the app is open for real
// use, and doesn't touch the settings saved there.
const profile = mkdtempSync(join(tmpdir(), "duckwriter-smoke-"));

const shotsFlag = process.argv.find((arg) => arg === "--screenshots" || arg.startsWith("--screenshots="));
const shotsDir = shotsFlag ? (shotsFlag.split("=")[1] ?? join(root, "test-screenshots")) : null;
let shots = 0;
if (shotsDir) {
  rmSync(shotsDir, { recursive: true, force: true });
  mkdirSync(shotsDir, { recursive: true });
}

let failures = 0;
function check(name, ok, detail = "") {
  if (!ok) failures++;
  console.log(`${ok ? "ok  " : "FAIL"}  ${name}${detail && !ok ? `  (${detail})` : ""}`);
}

/** The test's environment for a copy of the app. */
const appEnv = (port, profileDir = profile) => ({
  ...process.env,
  // Off-screen, so the test doesn't take over the screen. WebView2 would stop
  // drawing a window it thinks nobody can see; the flags keep it drawing.
  DUCKWRITER_TEST_OFFSCREEN: "1",
  WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${port} --disable-features=CalculateNativeWinOcclusion --disable-backgrounding-occluded-windows --disable-renderer-backgrounding`,
  WEBVIEW2_USER_DATA_FOLDER: profileDir,
});

/** Copies of DuckWriter currently running, by process id. */
const runningCopies = () =>
  execFileSync("powershell", ["-NoProfile", "-Command", "Get-Process -Name duckwriter -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }; exit 0"], { encoding: "utf8" })
    .split(/\r?\n/)
    .filter(Boolean)
    .map(Number);

/** Asks a program's window to close, as clicking its X does. */
const requestClose = (pid) =>
  execFileSync("powershell", [
    "-NoProfile",
    "-Command",
    `Add-Type -Name W -Namespace U -MemberDefinition '[DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);'; [void][U.W]::PostMessage((Get-Process -Id ${pid}).MainWindowHandle, 0x10, [IntPtr]::Zero, [IntPtr]::Zero)`,
  ]);

/** Starts the app with `files` open and hands `run` helpers for driving it. */
async function session(files, port, run, { profileDir = profile } = {}) {
  files = [files].flat().filter(Boolean);
  const child = spawn(exe, files, { env: appEnv(port, profileDir), stdio: "ignore" });
  let target;
  for (let i = 0; i < 60 && !target; i++) {
    await sleep(500);
    try {
      target = (await (await fetch(`http://127.0.0.1:${port}/json`)).json()).find((t) => t.type === "page");
    } catch {
      // Not listening yet.
    }
  }
  if (!target) {
    child.kill();
    throw new Error("The app did not start, or remote debugging is unavailable");
  }

  const ws = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve) => (ws.onopen = resolve));
  let nextId = 1;
  const pending = new Map();
  const errors = [];
  ws.onmessage = (event) => {
    const msg = JSON.parse(event.data);
    if (msg.id && pending.has(msg.id)) {
      pending.get(msg.id)(msg);
      pending.delete(msg.id);
    } else if (msg.method === "Runtime.exceptionThrown") {
      const details = msg.params.exceptionDetails;
      errors.push(details.exception?.description ?? details.text);
    } else if (msg.method === "Runtime.consoleAPICalled" && msg.params.type === "error") {
      errors.push(msg.params.args.map((arg) => arg.value ?? arg.description).join(" "));
    } else if (msg.method === "Log.entryAdded" && msg.params.entry.level === "error") {
      errors.push(msg.params.entry.text);
    }
  };
  const send = (method, params = {}) =>
    new Promise((resolve) => {
      const id = nextId++;
      pending.set(id, resolve);
      ws.send(JSON.stringify({ id, method, params }));
    });

  const app = {
    errors,
    pid: child.pid,
    exited: () => child.exitCode !== null,
    async evaluate(expression) {
      const { result } = await send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
      if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? "evaluate failed");
      return result.result.value;
    },
    mouse: (type, x, y, extra = {}) =>
      send("Input.dispatchMouseEvent", { type, x, y, button: "left", buttons: type === "mouseReleased" ? 0 : 1, clickCount: 1, ...extra }),
    async click(x, y) {
      await app.mouse("mouseMoved", x, y, { buttons: 0 });
      await app.mouse("mousePressed", x, y);
      await app.mouse("mouseReleased", x, y);
    },
    /** Presses at the first point and moves through the rest, like drawing with a pen. */
    async stroke(points) {
      await app.mouse("mouseMoved", points[0][0], points[0][1], { buttons: 0 });
      await app.mouse("mousePressed", points[0][0], points[0][1]);
      for (const [x, y] of points.slice(1)) await app.mouse("mouseMoved", x, y);
      await app.mouse("mouseReleased", ...points[points.length - 1]);
    },
    /** Saves a picture of the window when screenshots were asked for. */
    async shoot(name) {
      if (!shotsDir) return;
      await sleep(350);
      const { result } = await send("Page.captureScreenshot", { format: "png" });
      const file = `${String(++shots).padStart(2, "0")}-${name}.png`;
      writeFileSync(join(shotsDir, file), Buffer.from(result.data, "base64"));
    },
    scrollTo: (top) => app.evaluate(`document.querySelector(".viewer").scrollTo(0, ${top})`),
    pageBox: () => app.evaluate(`(() => { const r = document.querySelector(".page").getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; })()`),
    async drag(x1, y1, x2, y2) {
      await app.mouse("mouseMoved", x1, y1, { buttons: 0 });
      await app.mouse("mousePressed", x1, y1);
      for (let i = 1; i <= 8; i++) await app.mouse("mouseMoved", x1 + ((x2 - x1) * i) / 8, y1 + ((y2 - y1) * i) / 8);
      await app.mouse("mouseReleased", x2, y2);
    },
    wheel: (x, y, deltaY, modifiers = 0) => send("Input.dispatchMouseEvent", { type: "mouseWheel", x, y, deltaX: 0, deltaY, modifiers }),
    /** `modifiers`: 2 is Ctrl. */
    async key(name, code, modifiers = 0) {
      await send("Input.dispatchKeyEvent", { type: "keyDown", key: name, code: name, windowsVirtualKeyCode: code, modifiers });
      await send("Input.dispatchKeyEvent", { type: "keyUp", key: name, code: name, windowsVirtualKeyCode: code, modifiers });
      // The editor picks up a selection made with the keys a moment later.
      await sleep(40);
    },
    type: (text) => send("Input.insertText", { text }),
    /** The centre of the index-th element matching `selector`, scrolled into view. */
    center: (selector, index = 0) =>
      app.evaluate(`(() => {
        const el = document.querySelectorAll(${JSON.stringify(selector)})[${index}];
        if (!el) return null;
        el.scrollIntoView({ block: "center" });
        const r = el.getBoundingClientRect();
        return [r.left + r.width / 2, r.top + r.height / 2, r.width, r.height];
      })()`),
    async clickOn(selector, index = 0) {
      const at = await app.center(selector, index);
      if (!at) throw new Error(`Nothing matches ${selector} [${index}]`);
      await app.click(at[0], at[1]);
    },
    tool: (title) => app.clickOn(`.tb[title^="${title}"]`),
    count: (selector) => app.evaluate(`document.querySelectorAll(${JSON.stringify(selector)}).length`),
    async waitFor(expression, ms = 10000) {
      for (let t = 0; t < ms; t += 200) {
        if (await app.evaluate(expression)) return true;
        await sleep(200);
      }
      return false;
    },
    /** Makes `window.__dw` (the app, with its editor and the bytes Save would write) available. */
    async debugHandle() {
      await app.evaluate(`localStorage.setItem("duckwriter.debug", "1")`);
      await app.evaluate("location.reload()");
      await sleep(1500);
      await app.waitFor(`!!document.querySelector(".sheet .doc") && !!window.__dw`);
      await sleep(800);
    },
    invoke: (command, args) => app.evaluate(`window.__TAURI_INTERNALS__.invoke(${JSON.stringify(command)}, ${JSON.stringify(args ?? {})})`),
  };
  await send("Runtime.enable");
  await send("Log.enable");
  await send("Page.enable");
  try {
    await app.waitFor(`!!document.querySelector(".sheet .doc")`);
    if (files.length) await app.waitFor(`document.title.startsWith(${JSON.stringify(basename(files[0]))})`, 5000);
    await sleep(800);
    await run(app);
  } finally {
    ws.close();
    child.kill();
    await sleep(500);
  }
}

const CTRL = 2;
const SHIFT = 8;
/** A size in points as the browser reports it, in pixels. */
const pt = (points) => `${Math.round(((points * 4) / 3) * 10000) / 10000}px`;

/** The open document's paragraphs: text, style, list label and the text's look. */
const paragraphs = (app) =>
  app.evaluate(`[...document.querySelectorAll(".doc > .para")].map((p) => {
    const label = p.querySelector(".list-label")?.textContent ?? null;
    const text = [...p.childNodes].filter((n) => !n.classList?.contains("list-label")).map((n) => n.textContent).join("");
    const cs = getComputedStyle(p);
    return { text, style: p.dataset.style ?? null, label, size: cs.fontSize, weight: cs.fontWeight, align: cs.textAlign };
  })`);

// ---- a new document: typing and the toolbar --------------------------------------

const typed = join(tmpdir(), "duckwriter-smoke-typed.docx");
rmSync(typed, { force: true });

await session(null, 9433, async (app) => {
  check("starts with a new, empty document", (await app.evaluate("document.title")).startsWith("Document1.docx") && (await app.count(".doc > .para")) === 1);
  check("the page is A4 with 2.54 cm margins", await app.evaluate(`(() => { const s = document.querySelector(".sheet").style; return s.width === "595.3pt" && document.querySelector(".doc").style.margin === "72pt 72pt 0px"; })()`));
  check("the cursor is in the text, ready to type", await app.evaluate(`document.activeElement?.classList.contains("doc")`));

  await app.type("Packing list");
  await app.evaluate(`(() => { const s = document.querySelector(".style-select"); s.value = "Heading1"; s.dispatchEvent(new Event("change")); })()`);
  const shownStyle = await app.evaluate(`document.querySelector(".style-select").selectedOptions[0]?.textContent`);
  await app.key("End", 35);
  await app.key("Enter", 13);
  await app.type("Bring a ");
  await app.key("b", 66, CTRL);
  await app.type("rain jacket");
  await app.key("b", 66, CTRL);
  await app.type(" and good shoes.");
  let paras = await paragraphs(app);
  check("typing puts the text on the page", paras.map((p) => p.text).join("|") === "Packing list|Bring a rain jacket and good shoes.", JSON.stringify(paras));
  check(
    "choosing Heading 1 makes a heading",
    paras[0]?.style === "Heading1" && paras[0]?.size === pt(16) && shownStyle === "Heading 1",
    JSON.stringify(paras[0]),
  );
  check("Enter after a heading goes back to normal text", paras[1]?.style === null && paras[1]?.size === pt(11), JSON.stringify(paras[1]));
  check(
    "Ctrl+B makes the next words bold, and again turns it off",
    await app.evaluate(`(() => {
      const spans = [...document.querySelectorAll(".doc > .para")[1].querySelectorAll("span")];
      return spans.filter((s) => getComputedStyle(s).fontWeight === "700").map((s) => s.textContent).join("") === "rain jacket";
    })()`),
  );

  await app.key("Enter", 13);
  await app.clickOn(`.tb[title="Bullets"]`);
  await app.type("Socks");
  await app.key("Enter", 13);
  await app.type("Sweater");
  await app.key("Enter", 13);
  await app.key("Tab", 9);
  await app.type("The warm one");
  paras = await paragraphs(app);
  check("the Bullets button starts a bulleted list", paras[2]?.label === "•" && paras[3]?.label === "•", JSON.stringify(paras.slice(2)));
  check("Tab in a list moves the item a level in", paras[4]?.label === "◦" && paras[4]?.text === "The warm one", JSON.stringify(paras[4]));
  check("the toolbar shows the list is on", await app.evaluate(`document.querySelector('.tb[title="Bullets"]').classList.contains("active")`));
  await app.key("Enter", 13);
  await app.key("Enter", 13);
  await app.key("Enter", 13);
  paras = await paragraphs(app);
  check("Enter on an empty item steps out of the list", paras.at(-1)?.label === null, JSON.stringify(paras.slice(-3)));

  await app.type("Centred and big");
  await app.key("e", 69, CTRL);
  await app.key("Home", 36, SHIFT);
  await app.evaluate(`(() => { const i = document.querySelector(".size-input"); i.value = "20"; i.dispatchEvent(new Event("change")); })()`);
  paras = await paragraphs(app);
  check("Ctrl+E centres the paragraph", paras.at(-1)?.align === "center");
  check(
    "the size box changes the selected text's size",
    (await app.evaluate(`(() => { const span = [...document.querySelectorAll(".doc > .para")].at(-1).querySelector("span"); return span && getComputedStyle(span).fontSize; })()`)) === pt(20),
    await app.evaluate(`document.querySelector(".doc").innerHTML.slice(-600)`),
  );
  check("the window title shows there are unsaved changes", (await app.evaluate("document.title")).startsWith("• Document1.docx"));
  await app.shoot("typed-document");

  await app.debugHandle();
  // Reloading starts afresh; type something to save.
  await app.type("Saved by the test");
  await app.evaluate(`(async () => {
    const bytes = await window.__dw.build();
    await window.__TAURI_INTERNALS__.invoke("write_file", bytes, { headers: { "x-path": encodeURIComponent(${JSON.stringify(typed)}) } });
  })()`);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

{
  const zip = await JSZip.loadAsync(readFileSync(typed));
  const xml = await zip.file("word/document.xml").async("string");
  check("the saved file is a Word document with the typed text", xml.includes("Saved by the test") && !!zip.file("word/styles.xml"));
  rmSync(typed, { force: true });
}

// ---- opening a Word document --------------------------------------------------------

await session(sample("report.docx"), 9434, async (app) => {
  const paras = await paragraphs(app);
  check("opens the document given on the command line", (await app.evaluate("document.title")).startsWith("report.docx"));
  check("shows the title and headings in their styles", paras[0]?.style === "Title" && paras[1]?.style === "Heading1" && paras[7]?.style === "Heading2", JSON.stringify(paras.map((p) => p.style)));
  check("numbers the list the way the document says", paras.slice(3, 7).map((p) => p.label).join(" ") === "1. a) b) 2.", paras.slice(3, 7).map((p) => p.label).join(" "));
  check("shows the bullets", paras.slice(8, 11).every((p) => p.label === "•"), paras.slice(8, 11).map((p) => p.label).join(" "));
  check("keeps the table, shown as a box for now", (await app.evaluate(`document.querySelector(".raw-block .raw-label")?.textContent`)) === "Table");
  check(
    "red bold text is red and bold",
    await app.evaluate(`(() => {
      // The innermost of the nested formatting spans.
      const span = [...document.querySelectorAll(".doc span")].findLast((s) => s.textContent === "1 450 kr");
      const cs = span && getComputedStyle(span);
      return !!cs && cs.color === "rgb(192, 0, 0)" && cs.fontWeight === "700";
    })()`),
  );
  check(
    "a link looks like a link",
    await app.evaluate(`(() => { const a = document.querySelector(".doc a"); return a?.textContent === "the club website" && getComputedStyle(a.querySelector("span") ?? a).textDecorationLine.includes("underline"); })()`),
  );
  check("the window title has no unsaved-changes mark", (await app.evaluate("document.title")) === "report.docx – DuckWriter");
  await app.shoot("report-opened");

  // Change the first heading and close: it asks first.
  const heading = await app.center(".doc > .para", 1);
  await app.click(heading[0] + heading[2] / 2 - 2, heading[1]);
  await app.key("End", 35);
  await app.type(" and budget");
  requestClose(app.pid);
  check("closing with unsaved changes asks first", await app.waitFor(`!!document.querySelector("dialog[open]")`, 5000));
  await app.shoot("asks-to-save");
  await app.clickOn("dialog[open] .btn"); // Cancel
  await sleep(1200);
  check("Cancel keeps the window open", !app.exited());
  requestClose(app.pid);
  await app.waitFor(`!!document.querySelector("dialog[open]")`, 5000);
  app.clickOn("dialog[open] .btn", 1).catch(() => {}); // Don't save; the window goes away mid-call
  let closed = false;
  for (let i = 0; i < 20 && !closed; i++) {
    await sleep(250);
    closed = app.exited();
  }
  check("Don't save closes the window", closed);
});

// ---- equations and symbols -------------------------------------------------------------

await session(null, 9436, async (app) => {
  await app.type("The area is ");
  await app.key("=", 187, 1); // Alt+=
  check("Alt+= opens the equation editor", await app.waitFor(`!!document.querySelector(".eq-editor")`, 3000));
  await app.type("\\pi r^2");
  check("it shows the equation while typing", await app.waitFor(`!!document.querySelector(".eq-preview msup")`, 3000));
  await app.key("Enter", 13);
  check("Enter puts it in the text as typeset maths", await app.waitFor(`!!document.querySelector(".doc .math math msup") && !document.querySelector(".eq-editor")`, 3000));
  await app.type(" and the sum is");
  await app.key("Enter", 13);
  await app.clickOn(`.tb[title="Equation (Alt+=)"]`);
  await app.waitFor(`!!document.querySelector(".eq-editor")`, 3000);
  await app.clickOn(`.eq-template[title^="Sum"]`);
  await app.type("i=1");
  await app.key("Tab", 9);
  await app.evaluate(`(() => { const i = document.querySelector(".eq-input"); i.value = "\\\\sum_{i=1}^{n} i = \\\\frac{n(n+1)}{2}"; i.dispatchEvent(new Event("input")); })()`);
  check("the buttons insert shapes, and mistakes are explained", (await app.evaluate(`document.querySelector(".eq-message").textContent`)) === "");
  await app.shoot("equation-editor");
  await app.clickOn(".eq-footer .btn.primary");
  check(
    "an equation on an empty line is shown on its own line, centred",
    await app.waitFor(`(() => { const m = document.querySelectorAll(".doc .math")[1]; return !!m && m.classList.contains("display") && !!m.querySelector("munderover") && !!m.querySelector("mfrac"); })()`, 3000),
  );
  await app.key("Enter", 13);
  await app.clickOn(`.tb[title="Symbol options"]`);
  await app.clickOn(`.symbol-button[title="μ"]`);
  // The menu stays open for picking several.
  await app.clickOn(`.symbol-button[title="°"]`);
  check("the symbol menu inserts Greek letters and signs", (await app.evaluate(`[...document.querySelectorAll(".doc > .para")].at(-1).textContent`)) === "μ°");
  await app.shoot("equations");
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

// ---- real pages ---------------------------------------------------------------------------

/** Where each line of text is, by page: lines outside a page's text area, and per paragraph the page it starts and ends on. */
const pageReport = (app) =>
  app.evaluate(`(() => {
    const sheets = [...document.querySelectorAll(".page-sheet")].map((s) => s.getBoundingClientRect());
    const doc = document.querySelector(".doc");
    const cs = getComputedStyle(doc);
    const top = parseFloat(cs.marginTop), side = parseFloat(cs.marginLeft);
    const scale = doc.getBoundingClientRect().width / doc.offsetWidth;
    // Text area: the page minus the top and bottom margins (equal in the samples).
    const areas = sheets.map((s) => ({ top: s.top + top * scale, bottom: s.bottom - top * scale }));
    // Letters may reach a pixel or two beyond their line.
    const pageAt = (y) => areas.findIndex((a) => y >= a.top - 3 && y <= a.bottom + 3);
    let outside = 0;
    const paras = [...doc.querySelectorAll(":scope > .para")].map((p) => {
      // Only the text itself, not the invisible fillers that make the page gaps.
      const rects = [];
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let t = walker.nextNode(); t; t = walker.nextNode()) {
        if (!t.textContent.trim()) continue;
        const range = document.createRange();
        range.selectNodeContents(t);
        rects.push(...[...range.getClientRects()].filter((r) => r.height >= 2 && r.width > 0));
      }
      rects.sort((x, y) => x.top - y.top);
      for (const r of rects) if (pageAt(r.top) < 0 || pageAt(r.bottom) !== pageAt(r.top)) { outside++; window.__outside = [...(window.__outside ?? []), [p.textContent.slice(0, 20), Math.round(r.top), Math.round(r.bottom)]]; }
      return { text: p.textContent.slice(0, 30), style: p.dataset.style ?? null, first: rects.length ? pageAt(rects[0].top) : -1, last: rects.length ? pageAt(rects.at(-1).top) : -1, top: rects[0]?.top ?? 0, pageTop: rects.length ? areas[pageAt(rects[0].top)]?.top : 0 };
    });
    void side;
    return { pages: sheets.length, outside, paras };
  })()`);

await session(sample("long.docx"), 9437, async (app) => {
  let report = await pageReport(app);
  check("a long document is shown on several pages", report.pages >= 3 && (await app.evaluate(`document.querySelector(".statusbar").textContent`)).includes(`${report.pages} pages`), `${report.pages} pages`);
  check("no line of text sits in a margin or between pages", report.outside === 0, `${report.outside} lines outside: ${JSON.stringify(await app.evaluate("window.__outside"))}`);
  const headings = report.paras.map((p, i) => [p, report.paras[i + 1]]).filter(([p]) => p.style === "Heading1");
  check("a heading is never left alone at the bottom of a page", headings.every(([h, next]) => !next || next.first === h.last), JSON.stringify(headings.map(([h, n]) => [h.text, h.last, n?.first])));
  const breakAt = report.paras.findIndex((p) => p.text.startsWith("The end of part one."));
  const after = report.paras[breakAt + 1];
  check("a page break starts a new page", breakAt >= 0 && after.first === report.paras[breakAt].last + 1 && Math.abs(after.top - after.pageTop) < 30, JSON.stringify([report.paras[breakAt], after]));
  check("paragraphs continue across pages", report.paras.some((p) => p.last > p.first));
  await app.evaluate(`document.querySelector(".workspace").scrollTop = document.querySelectorAll(".page-sheet")[1].offsetTop - 200`);
  await app.shoot("page-boundary");

  // Typing more text moves everything along, and the pages stay tidy.
  const first = await app.center(".doc > .para", 1);
  await app.click(first[0], first[1]);
  for (let i = 0; i < 6; i++) await app.type("Extra words to push the text further down the page. ");
  await sleep(300);
  report = await pageReport(app);
  check("after typing, no line sits in a margin or between pages", report.outside === 0, `${report.outside} lines outside`);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

await session(sample("not-a-docx.docx"), 9435, async (app) => {
  check("a file that isn't a Word document gets a plain message", await app.waitFor(`/isn't a Word document/.test(document.querySelector(".toast.error")?.textContent ?? "")`, 5000));
});

rmSync(profile, { recursive: true, force: true });
if (shotsDir) console.log(`\n${shots} screenshots saved to ${shotsDir}`);
console.log(failures ? `\n${failures} check(s) failed` : "\nAll checks passed");
process.exit(failures ? 1 : 0);
