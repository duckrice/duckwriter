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
    async doubleClick(x, y) {
      await app.click(x, y);
      await app.mouse("mousePressed", x, y, { clickCount: 2 });
      await app.mouse("mouseReleased", x, y, { clickCount: 2 });
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
    /** Presses a key. `physical` is the key's place (KeyB, Digit2, IntlBackslash…), worked out for letters and digits. */
    async key(name, code, modifiers = 0, physical = null, text = undefined) {
      const where = physical ?? (/^[a-z]$/i.test(name) ? `Key${name.toUpperCase()}` : /^[0-9]$/.test(name) ? `Digit${name}` : name);
      await send("Input.dispatchKeyEvent", { type: "keyDown", key: name, code: where, windowsVirtualKeyCode: code, modifiers, ...(text ? { text } : {}) });
      await send("Input.dispatchKeyEvent", { type: "keyUp", key: name, code: where, windowsVirtualKeyCode: code, modifiers });
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
    /** Clicks the first element matching `selector` whose text starts with `text`. */
    async clickText(selector, text) {
      const at = await app.evaluate(`(() => {
        const el = [...document.querySelectorAll(${JSON.stringify(selector)})].find((e) => e.textContent.trim().startsWith(${JSON.stringify(text)}));
        if (!el) return null;
        el.scrollIntoView({ block: "nearest" });
        const r = el.getBoundingClientRect();
        return [r.left + r.width / 2, r.top + r.height / 2];
      })()`);
      if (!at) throw new Error(`No ${selector} reading "${text}"`);
      await app.click(at[0], at[1]);
    },
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
const ESC = ["Escape", 27];
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
  await app.clickOn(`.tb[title^="Bullets"]`);
  await app.type("Socks");
  await app.key("Enter", 13);
  await app.type("Sweater");
  await app.key("Enter", 13);
  await app.key("Tab", 9);
  await app.type("The warm one");
  paras = await paragraphs(app);
  check("the Bullets button starts a bulleted list", paras[2]?.label === "•" && paras[3]?.label === "•", JSON.stringify(paras.slice(2)));
  check("Tab in a list moves the item a level in", paras[4]?.label === "◦" && paras[4]?.text === "The warm one", JSON.stringify(paras[4]));
  check("the toolbar shows the list is on", await app.evaluate(`document.querySelector('.tb[title^="Bullets"]').classList.contains("active")`));
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
  await app.clickOn(`.tb[title^="Symbol"][title$="options"]`);
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

// ---- headers, footers and page numbers -------------------------------------------------

const withHeaders = join(tmpdir(), "duckwriter-smoke-headers.docx");
rmSync(withHeaders, { force: true });

await session(null, 9438, async (app) => {
  await app.type("First page");
  await app.key("Enter", 13, CTRL); // page break
  await app.type("Second page");
  await app.waitFor(`document.querySelectorAll(".page-sheet").length === 2`, 5000);
  const area = (kind, page) => app.center(`.hf.${kind}[data-page="${page}"]`);
  let at = await area("header", 0);
  await app.doubleClick(at[0], at[1]);
  check("double-clicking the top of a page edits its header", await app.waitFor(`!!document.querySelector(".hf.header.editing .ProseMirror") && !document.querySelector(".hf-bar").hidden`, 3000));
  await app.type("Lab report");
  await app.key(...ESC);
  const headerText = (page) => app.evaluate(`document.querySelector('.hf.header[data-page="${page}"]').textContent`);
  check("the header shows on every page", (await headerText(0)) === "Lab report" && (await headerText(1)) === "Lab report", `${await headerText(0)} / ${await headerText(1)}`);

  at = await area("footer", 0);
  await app.doubleClick(at[0], at[1]);
  await app.waitFor(`!!document.querySelector(".hf.footer.editing .ProseMirror")`, 3000);
  await app.type("Page ");
  await app.clickOn(`.hf-bar .btn[title="Insert the page number"]`);
  await app.type(" of ");
  await app.clickOn(`.hf-bar .btn[title="Insert the number of pages"]`);
  await app.shoot("editing-footer");
  await app.key(...ESC);
  const footerText = (page) => app.evaluate(`document.querySelector('.hf.footer[data-page="${page}"]').textContent`);
  check("page numbers count up from page to page", (await footerText(0)) === "Page 1 of 2" && (await footerText(1)) === "Page 2 of 2", `${await footerText(0)} / ${await footerText(1)}`);

  at = await area("header", 0);
  await app.doubleClick(at[0], at[1]);
  await app.waitFor(`!!document.querySelector(".hf.header.editing")`, 3000);
  await app.evaluate(`(() => { const c = document.querySelector(".hf-check input"); c.checked = true; c.dispatchEvent(new Event("change")); })()`);
  await app.waitFor(`document.activeElement?.closest(".hf.header.editing") && document.querySelector(".hf-bar-label").textContent.includes("first page")`, 3000);
  await app.type("Cover page");
  await app.key(...ESC);
  check("the first page can have its own header", (await headerText(0)) === "Cover page" && (await headerText(1)) === "Lab report", `${await headerText(0)} / ${await headerText(1)}`);
  check("the window title shows there are unsaved changes", (await app.evaluate("document.title")).startsWith("•"));
  await app.shoot("headers-and-footers");

  await app.debugHandle();
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

await session(sample("report.docx"), 9439, async (app) => {
  const headerText = await app.evaluate(`document.querySelector('.hf.header[data-page="0"]').textContent`);
  const footerText = await app.evaluate(`document.querySelector('.hf.footer[data-page="0"]').textContent`);
  check("headers and footers of an opened document are shown, with page numbers", headerText === "Bergen trip 2026" && footerText === "Page 1 of 1", `${headerText} / ${footerText}`);
});

// ---- footnotes ------------------------------------------------------------------------------

const withNotes = join(tmpdir(), "duckwriter-smoke-notes.docx");
rmSync(withNotes, { force: true });

await session(null, 9450, async (app) => {
  await app.debugHandle();
  await app.type("The reaction rate doubled");
  await app.key("f", 70, CTRL | 1); // Ctrl+Alt+F
  check("Ctrl+Alt+F adds a footnote and opens it at the bottom of the page", await app.waitFor(`!!document.querySelector(".fn-note.editing .ProseMirror")`, 3000));
  await app.type("Measured at 20 °C.");
  await app.key(...ESC);
  const notes = () => app.evaluate(`[...document.querySelectorAll(".fn-note")].map((n) => n.textContent.trim())`);
  check("the note is shown under a line at the bottom of the page, numbered", JSON.stringify(await notes()) === JSON.stringify(["1 Measured at 20 °C."]), JSON.stringify(await notes()));

  // A note added earlier in the text comes first.
  await sleep(300);
  await app.key("Home", 36);
  await app.type("Surprisingly, ");
  await app.key("f", 70, CTRL | 1);
  await app.waitFor(`!!document.querySelector(".fn-note.editing .ProseMirror")`, 3000);
  await app.type("Nobody expected this.");
  await app.key(...ESC);
  check("notes are numbered in the order of their marks", JSON.stringify(await notes()) === JSON.stringify(["1 Nobody expected this.", "2 Measured at 20 °C."]), JSON.stringify(await notes()));
  const box = await app.evaluate(`(() => { const a = document.querySelector(".fn-area").getBoundingClientRect(); const s = document.querySelector(".page-sheet").getBoundingClientRect(); return [a.bottom, s.bottom, s.bottom - 96 * (document.querySelector(".sheet").getBoundingClientRect().width / document.querySelector(".sheet").offsetWidth)]; })()`);
  check("the notes sit just above the bottom margin", Math.abs(box[0] - box[2]) < 2, JSON.stringify(box));
  await app.shoot("footnotes");
  await app.evaluate(`(async () => {
    const bytes = await window.__dw.build();
    await window.__TAURI_INTERNALS__.invoke("write_file", bytes, { headers: { "x-path": encodeURIComponent(${JSON.stringify(withNotes)}) } });
  })()`);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

{
  const zip = await JSZip.loadAsync(readFileSync(withNotes));
  const xml = await zip.file("word/footnotes.xml")?.async("string");
  check("the saved file has the footnotes, in order", !!xml && xml.indexOf("Nobody expected") < xml.indexOf("Measured at 20") && xml.includes("<w:footnoteRef/>"));
  rmSync(withNotes, { force: true });
}

// Footnotes near the end of a page: the note goes where its mark goes, and text keeps clear of it.
await session(sample("long.docx"), 9451, async (app) => {
  // Put a long note in the last paragraph that starts on page 1.
  const target = await app.evaluate(`(() => {
    const first = document.querySelectorAll(".page-sheet")[0].getBoundingClientRect();
    const paras = [...document.querySelectorAll(".doc > .para")].filter((p) => p.getBoundingClientRect().top < first.bottom - 200);
    const p = paras.at(-1);
    p.scrollIntoView({ block: "center" });
    const r = p.getBoundingClientRect();
    return [r.left + 40, r.top + 8];
  })()`);
  await app.click(target[0], target[1]);
  await app.key("End", 35);
  await app.key("f", 70, CTRL | 1);
  await app.waitFor(`!!document.querySelector(".fn-note.editing .ProseMirror")`, 3000);
  await app.type("A long note. ".repeat(40));
  await app.key(...ESC);
  await sleep(500);
  const report = await app.evaluate(`(() => {
    const sheets = [...document.querySelectorAll(".page-sheet")].map((s) => s.getBoundingClientRect());
    const pageAt = (y) => sheets.findIndex((s) => y >= s.top && y <= s.bottom);
    const mark = document.querySelector(".doc .footnote-ref").getBoundingClientRect();
    const note = document.querySelector(".fn-note").getBoundingClientRect();
    const area = document.querySelector(".fn-area:not([hidden])").getBoundingClientRect();
    // Text on the note's page must end above the notes.
    let overlap = 0;
    for (const p of document.querySelectorAll(".doc > .para")) {
      const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let t = w.nextNode(); t; t = w.nextNode()) {
        const r = document.createRange(); r.selectNodeContents(t);
        for (const x of r.getClientRects()) if (x.height >= 2 && pageAt(x.top) === pageAt(area.top) && x.bottom > area.top + 3) overlap++;
      }
    }
    return { markPage: pageAt(mark.top), notePage: pageAt(note.top), overlap };
  })()`);
  check("a note is on the same page as its mark", report.markPage === report.notePage && report.markPage >= 0, JSON.stringify(report));
  check("text never runs into the notes", report.overlap === 0, JSON.stringify(report));
  await app.evaluate(`document.querySelector(".fn-area:not([hidden])").scrollIntoView({ block: "center" })`);
  await app.shoot("footnote-near-page-end");
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

// ---- captions, cross-references, numbered headings, table of contents -------------------------

const withRefs = join(tmpdir(), "duckwriter-smoke-references.docx");
rmSync(withRefs, { force: true });

await session(null, 9460, async (app) => {
  await app.debugHandle();
  const setStyle = (id) => app.evaluate(`(() => { const s = document.querySelector(".style-select"); s.value = "${id}"; s.dispatchEvent(new Event("change")); })()`);
  const references = () => app.clickOn(`.tb[title^="References"][title$="options"]`);
  await app.type("Introduction");
  await setStyle("Heading1");
  await app.key("End", 35);
  await app.key("Enter", 13);
  await app.type("Some text about the experiment.");
  await app.key("Enter", 13);
  await app.type("Setup");
  await setStyle("Heading2");
  await app.key("End", 35);
  await app.key("Enter", 13);
  await app.type("Picture of the setup goes here.");

  await references();
  await app.clickText(".references-menu .menu-item", "Numbered headings");
  const labels = () => app.evaluate(`[...document.querySelectorAll(".doc > .para")].map((p) => p.querySelector(".list-label")?.textContent ?? "")`);
  check("numbered headings: 1, 1.1", JSON.stringify((await labels()).filter(Boolean)) === JSON.stringify(["1", "1.1"]), JSON.stringify(await labels()));

  // A caption under the picture paragraph.
  await references();
  await app.clickText(".references-menu .menu-item", "Insert caption");
  await app.waitFor(`!!document.querySelector("dialog[open] .dialog-form")`, 3000);
  await app.type("The setup");
  await app.clickOn("dialog[open] .btn.primary");
  const captions = () => app.evaluate(`[...document.querySelectorAll('.doc > .para[data-style="Caption"]')].map((p) => p.textContent)`);
  check("a caption is added and numbered", JSON.stringify(await captions()) === JSON.stringify(["Figure 1: The setup"]), JSON.stringify(await captions()));
  check("it looks like a caption: small and italic", await app.evaluate(`(() => { const cs = getComputedStyle(document.querySelector('.doc > .para[data-style="Caption"]')); return cs.fontStyle === "italic" && cs.fontSize === "12px"; })()`));

  // A second figure before the first renumbers both.
  const intro = await app.center(".doc > .para", 1);
  await app.click(intro[0], intro[1]);
  await references();
  await app.clickText(".references-menu .menu-item", "Insert caption");
  await app.waitFor(`!!document.querySelector("dialog[open] .dialog-form")`, 3000);
  await app.type("Overview");
  await app.clickOn("dialog[open] .btn.primary");
  check("captions renumber themselves", JSON.stringify(await captions()) === JSON.stringify(["Figure 1: Overview", "Figure 2: The setup"]), JSON.stringify(await captions()));

  // "As Figure 2 shows": a cross-reference at the end of the text.
  const last = await app.center(".doc > .para", 1);
  await app.click(last[0], last[1]);
  await app.key("End", 35);
  await app.type(" As ");
  await references();
  await app.clickText(".references-menu .menu-item", "Figure 2: The setup");
  await app.type(" shows, it works.");
  const refText = await app.evaluate(`[...document.querySelectorAll(".doc .field")].find((f) => f.dataset.instr === "REF")?.textContent`);
  check("a cross-reference shows the figure's number", refText === "Figure 2", String(refText));

  // A table of contents at the top.
  await app.key("Home", 36, CTRL);
  await references();
  await app.clickText(".references-menu .menu-item", "Table of contents");
  check(
    "the table of contents lists the headings with their page numbers",
    await app.waitFor(`[...document.querySelectorAll(".toc-entry")].map((e) => e.textContent).join("|") === "1 Introduction1|1.1 Setup1"`, 5000),
    await app.evaluate(`[...document.querySelectorAll(".toc-entry")].map((e) => e.textContent).join("|")`),
  );
  await app.shoot("references");
  await app.evaluate(`(async () => {
    const bytes = await window.__dw.build();
    await window.__TAURI_INTERNALS__.invoke("write_file", bytes, { headers: { "x-path": encodeURIComponent(${JSON.stringify(withRefs)}) } });
  })()`);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

{
  const zip = await JSZip.loadAsync(readFileSync(withRefs));
  const xml = await zip.file("word/document.xml").async("string");
  check("the saved file has Word fields for the captions, the reference and the contents", xml.includes("SEQ Figure") && xml.includes(" REF _Ref") && xml.includes(" TOC "));
  rmSync(withRefs, { force: true });
}

// ---- pictures and snapping to the grid ------------------------------------------------------

const withPictures = join(tmpdir(), "duckwriter-smoke-pictures.docx");
rmSync(withPictures, { force: true });
const GRID = 14.173;

await session(null, 9470, async (app) => {
  await app.debugHandle();
  await app.evaluate(`localStorage.removeItem("duckwriter.snap")`);
  await app.type("The apparatus is shown here. ".repeat(12));
  await app.key("Home", 36, CTRL);
  // Paste a 400 × 300 picture, as from the clipboard.
  await app.evaluate(`(async () => {
    const canvas = document.createElement("canvas");
    canvas.width = 400; canvas.height = 300;
    const g = canvas.getContext("2d");
    g.fillStyle = "#0f766e"; g.fillRect(0, 0, 400, 300);
    g.fillStyle = "#facc15"; g.beginPath(); g.arc(200, 150, 90, 0, 7); g.fill();
    const blob = await new Promise((r) => canvas.toBlob(r, "image/png"));
    const data = new DataTransfer();
    data.items.add(new File([blob], "setup.png", { type: "image/png" }));
    window.__dw.body.dom.dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }));
  })()`);
  check("a pasted picture goes in line with the text, at its own size", await app.waitFor(`(() => { const img = document.querySelector(".doc .picture-box.inline img"); return !!img && img.style.width === "300pt"; })()`, 5000));
  const image = () => app.evaluate(`(() => { const n = []; window.__dw.body.state.doc.descendants((x) => { if (x.type.name === "image") n.push(x.attrs); }); return n[0]; })()`);

  // Resize from the bottom-right corner: the width snaps to the grid.
  const handle = await app.center(".picture-box .resize-handle.se");
  check("selecting a picture shows its handles and the wrapping choices", !!handle && (await app.count(".picture-bar button")) === 5);
  await app.drag(handle[0], handle[1], handle[0] - 97, handle[1] - 40);
  let attrs = await image();
  check("resizing snaps the width to the grid and keeps the proportions", Math.abs(attrs.width / GRID - Math.round(attrs.width / GRID)) < 0.01 && Math.abs(attrs.height / attrs.width - 0.75) < 0.01, JSON.stringify(attrs));

  // Text flows around it.
  await app.clickText(".picture-bar button", "Square");
  attrs = await image();
  check("Square makes it float with the text around it", attrs.floating && attrs.wrap === "square", JSON.stringify(attrs));
  const beside = await app.evaluate(`(() => {
    const img = document.querySelector(".picture-box img").getBoundingClientRect();
    const p = document.querySelector(".doc > .para");
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      const r = document.createRange(); r.selectNodeContents(t);
      for (const x of r.getClientRects()) if (x.top < img.bottom - 5 && x.bottom > img.top + 5 && x.left >= img.right) return true;
    }
    return false;
  })()`);
  check("the text runs beside the picture", beside);

  // Drag it: it lands on the grid, and the grid shows while dragging.
  const pic = await app.center(".picture-box img");
  await app.mouse("mouseMoved", pic[0], pic[1], { buttons: 0 });
  await app.mouse("mousePressed", pic[0], pic[1]);
  for (let i = 1; i <= 6; i++) await app.mouse("mouseMoved", pic[0] + i * 23, pic[1] + i * 11);
  check("the grid shows while a picture is moved", (await app.count(".snap-grid")) === 1);
  await app.shoot("dragging-a-picture");
  await app.mouse("mouseReleased", pic[0] + 138, pic[1] + 66);
  attrs = await image();
  const onGrid = (v) => Math.abs(v / GRID - Math.round(v / GRID)) < 0.01;
  check("a moved picture lands on the grid", onGrid(attrs.x) && onGrid(attrs.y) && attrs.x > 0, JSON.stringify(attrs));
  check("the grid goes away afterwards", (await app.count(".snap-grid")) === 0);
  await app.shoot("picture-with-text-around");

  await app.evaluate(`(async () => {
    const bytes = await window.__dw.build();
    await window.__TAURI_INTERNALS__.invoke("write_file", bytes, { headers: { "x-path": encodeURIComponent(${JSON.stringify(withPictures)}) } });
  })()`);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

{
  const zip = await JSZip.loadAsync(readFileSync(withPictures));
  const xml = await zip.file("word/document.xml").async("string");
  check("the saved file has the picture, floating with square wrapping", !!zip.file("word/media/duckwriter1.png") && xml.includes("<wp:anchor") && xml.includes("<wp:wrapSquare"));
}

await session(withPictures, 9471, async (app) => {
  check("the saved picture opens again where it was", await app.waitFor(`(() => { const img = document.querySelector(".doc .picture-box.floating img"); return !!img && img.naturalWidth === 400; })()`, 5000));
});
rmSync(withPictures, { force: true });

// ---- the line grid ----------------------------------------------------------------------------

await session(sample("long.docx"), 9480, async (app) => {
  await app.clickOn(`.tb[title="Layout options"]`);
  await app.clickText(".layout-menu .menu-item", "Line grid");
  await sleep(500);
  const report = await app.evaluate(`(() => {
    const doc = document.querySelector(".doc");
    const sheets = [...document.querySelectorAll(".page-sheet")].map((s) => s.getBoundingClientRect().top);
    const step = parseFloat(getComputedStyle(document.querySelector('.doc > .para:not([data-style])')).lineHeight);
    // Each line's top, measured from the top of the text area of its page, should be a whole number of steps.
    const margin = parseFloat(getComputedStyle(doc).marginTop);
    let off = 0, lines = 0;
    for (const p of document.querySelectorAll(".doc > .para:not([data-style])")) {
      const range = document.createRange(); range.selectNodeContents(p);
      const pTop = p.getBoundingClientRect().top + parseFloat(getComputedStyle(p).paddingTop);
      const page = sheets.filter((t) => t <= pTop + 1).length - 1;
      const fromTop = pTop - (sheets[page] + margin);
      const steps = fromTop / step;
      lines++;
      if (Math.abs(steps - Math.round(steps)) > 0.08) off++;
    }
    const heading = parseFloat(getComputedStyle(document.querySelector('.doc > .para[data-style="Heading1"]')).lineHeight);
    return { step, heading, off, lines, title: document.title };
  })()`);
  check("with the line grid on, headings take a whole number of lines", Math.abs(report.heading / report.step - Math.round(report.heading / report.step)) < 0.01 && report.heading > report.step, JSON.stringify(report));
  check("…and paragraphs start on the grid on every page", report.off === 0 && report.lines > 20, JSON.stringify(report));
  check("switching it on is a change to save", report.title.startsWith("•"));
  await app.shoot("line-grid");
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

// ---- keyboard shortcuts --------------------------------------------------------------------------

await session(null, 9490, async (app) => {
  const ALT = 1;
  await app.type("Size test");
  await app.key("a", 65, CTRL);
  const size = () => app.evaluate(`document.querySelector(".size-input").value`);
  // Norwegian keyboard: Ctrl+Shift+< types ">", Ctrl+< types "<".
  await app.key(">", 226, CTRL | SHIFT, "IntlBackslash");
  check("Ctrl+> (Ctrl+Shift+< on a Norwegian keyboard) makes text bigger", (await size()) === "12", await size());
  await app.key("¨", 221, CTRL, "BracketRight");
  check("Ctrl+] (the ¨ key on a Norwegian keyboard) adds one point", (await size()) === "13", await size());
  await app.key("<", 226, CTRL, "IntlBackslash");
  check("Ctrl+< makes it smaller again", (await size()) === "12", await size());
  await app.key("å", 219, CTRL, "BracketLeft");
  check("Ctrl+[ takes a point off", (await size()) === "11", await size());
  await app.key("P", 80, CTRL | SHIFT);
  check("Ctrl+Shift+P goes to the font size box", await app.evaluate(`document.activeElement === document.querySelector(".size-input")`));
  await app.type("18");
  await app.key("Enter", 13);
  check("typing a size there and Enter sets it, and goes back to the text", (await app.evaluate(`getComputedStyle(document.querySelector(".doc .para span")).fontSize`)) === pt(18) && (await app.evaluate(`document.activeElement?.classList.contains("doc")`)));

  // AltGr+2 types @ on a Norwegian keyboard: it must not become Heading 2.
  await app.key("End", 35);
  await app.key("@", 50, CTRL | ALT, "Digit2", "@");
  check("AltGr characters are typed, not taken as shortcuts", (await app.evaluate(`document.querySelector(".doc .para").textContent`)).endsWith("@") && !(await app.evaluate(`document.querySelector(".doc .para").dataset.style`)));

  await app.key("1", 49, CTRL | ALT);
  check("Ctrl+Alt+1 makes a heading", (await app.evaluate(`document.querySelector(".doc .para").dataset.style`)) === "Heading1");

  // Menus open from the keyboard and work with the arrow keys.
  await app.key("End", 35);
  await app.key("Y", 89, CTRL | SHIFT);
  check("Ctrl+Shift+Y opens the symbols, ready for the arrow keys", await app.waitFor(`document.activeElement?.classList.contains("symbol-button")`, 2000));
  await app.key("ArrowRight", 39);
  await app.key("Enter", 13);
  await app.key("Escape", 27);
  check("…and Enter puts the chosen one in", (await app.evaluate(`document.querySelector(".doc .para").textContent`)).endsWith("β"));

  await app.key("F1", 112);
  check("F1 lists every shortcut", await app.waitFor(`document.querySelectorAll(".shortcut-help .shortcut-row").length > 50`, 2000));
  await app.type("size");
  check("the list can be searched", (await app.evaluate(`[...document.querySelectorAll(".shortcut-help .shortcut-row")].map((r) => r.querySelector(".shortcut-label").textContent).join("|")`)).includes("Font size up by one point"));
  await app.shoot("keyboard-shortcuts");
  await app.key("Escape", 27);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

await session(sample("not-a-docx.docx"), 9435, async (app) => {
  check("a file that isn't a Word document gets a plain message", await app.waitFor(`/isn't a Word document/.test(document.querySelector(".toast.error")?.textContent ?? "")`, 5000));
});

rmSync(profile, { recursive: true, force: true });
if (shotsDir) console.log(`\n${shots} screenshots saved to ${shotsDir}`);
console.log(failures ? `\n${failures} check(s) failed` : "\nAll checks passed");
process.exit(failures ? 1 : 0);
