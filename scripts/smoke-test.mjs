// Drives the built desktop app over the DevTools protocol (WebView2 remote
// debugging) and checks the things unit tests can't: that documents open and
// look right, that typing and the toolbar work, and what ends up in a saved file.
//
//   npm run samples && npm run tauri build -- --no-bundle && npm run smoke
//
// Add `-- --screenshots` to also save pictures of the app along the way, into
// test-screenshots/ (or `-- --screenshots=some/folder`).
import { execFileSync, spawn } from "node:child_process";
import { copyFileSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import * as pdfjs from "pdfjs-dist/legacy/build/pdf.mjs";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const exe = join(root, "src-tauri", "target", "release", "duckwriter.exe");
const sample = (name) => join(root, "samples", name);
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
// Its own browser profile: the test then works while the app is open for real
// use, and doesn't touch the settings saved there.
const profile = mkdtempSync(join(tmpdir(), "duckwriter-smoke-"));
const recoveryDir = join(profile, "Recovery");

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

/** Removes a test folder; WebView2 may hold on to a profile for a moment after the app closes. */
function removeFolder(path) {
  try {
    rmSync(path, { recursive: true, force: true, maxRetries: 20, retryDelay: 250 });
  } catch {
    // Left for Windows to clean up with the other temporary files.
  }
}

/** The test's environment for a copy of the app. */
const appEnv = (port, profileDir = profile) => ({
  ...process.env,
  // Off-screen, so the test doesn't take over the screen. WebView2 would stop
  // drawing a window it thinks nobody can see; the flags keep it drawing.
  DUCKWRITER_TEST_OFFSCREEN: "1",
  WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: `--remote-debugging-port=${port} --disable-features=CalculateNativeWinOcclusion --disable-backgrounding-occluded-windows --disable-renderer-backgrounding`,
  WEBVIEW2_USER_DATA_FOLDER: profileDir,
  DUCKWRITER_RECOVERY_DIR: recoveryDir,
  // The same spelling languages on any computer the test runs on.
  DUCKWRITER_SPELL_LANGUAGES: "en-GB,nb-NO",
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
async function session(files, port, run, { profileDir = profile, keepAutoSave = false } = {}) {
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
    send,
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
        const tab = el.closest(".ribbon-panel [data-tab]");
        if (tab?.hidden) document.querySelector('.ribbon-tab[data-tab="' + tab.dataset.tab + '"]').click();
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
    // Tests edit the sample files; auto-save would change them. One test switches it on.
    if (!keepAutoSave) await app.evaluate(`localStorage.setItem("duckwriter.autosave", "off")`);
    if (files.length) await app.waitFor(`document.title.startsWith(${JSON.stringify(basename(files[0]))})`, 5000);
    await sleep(800);
    await run(app);
  } finally {
    ws.close();
    child.kill();
    await sleep(500);
    // Its unsaved work stays out of the next sessions: three minutes on, the app would offer
    // it back to whichever test starts then. (The recovery test keeps it on purpose.)
    if (!keepAutoSave) {
      let left = [];
      try {
        left = readdirSync(recoveryDir);
      } catch {
        // No recovery folder yet.
      }
      for (const f of left) rmSync(join(recoveryDir, f), { force: true });
    }
  }
}

const CTRL = 2;
const SHIFT = 8;
const ALT = 1;
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
  // Screen readers say each button's name: its text, aria-label or title. Every ribbon tab, and the status bar.
  const unnamed = await app.evaluate(`[...document.querySelectorAll("button, select, input")].filter((b) => !(b.getAttribute("aria-label") || b.textContent.trim() || b.title || b.labels?.length || b.closest("label"))).map((b) => b.outerHTML.slice(0, 120))`);
  check("every button and box has a name for screen readers", unnamed.length === 0, unnamed.join(" | "));

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
  // The size box's list shows every size, the current one marked, not only what matches the box.
  const [ax, ay] = await app.center(".size-input + .combo-arrow");
  await app.click(ax, ay);
  const offered = await app.evaluate(`(() => { const items = [...document.querySelectorAll(".size-input ~ .combo-list .combo-item")]; return items.length + ":" + (items.find((i) => i.classList.contains("active"))?.textContent ?? "none"); })()`);
  check("the font size list shows every size, with the current one marked", offered === "17:11", offered);
  await app.clickText(".size-input ~ .combo-list .combo-item", "16");
  check("choosing a size applies it and gives the text back the cursor", (await app.evaluate(`document.querySelector(".size-input").value`)) === "16" && (await app.evaluate(`!!document.activeElement?.closest(".doc")`)));
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
  check("shows the table, ready to edit", (await app.evaluate(`document.querySelector(".doc table")?.rows.length`)) === 3);
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
  // As in Word: a click selects the equation (for formatting), a double-click opens it.
  const eq = await app.center(".doc .math");
  await app.click(eq[0], eq[1]);
  check("clicking an equation selects it, without opening it", await app.waitFor(`document.querySelector(".doc .math").classList.contains("selected") && !document.querySelector(".eq-editor")`, 2000));
  await app.clickOn(".combo-arrow");
  await app.evaluate(`[...document.querySelectorAll(".combo-item")].find((i) => i.textContent === "Times New Roman").dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))`);
  check("a font from the list keeps it typeset", await app.waitFor(`!!document.querySelector(".doc .math math msup") && !!document.querySelector(".doc .math").closest('[style*="Times New Roman"]')`, 2000));
  await app.mouse("mousePressed", eq[0], eq[1], { clickCount: 2 });
  await app.mouse("mouseReleased", eq[0], eq[1], { clickCount: 2 });
  check("double-clicking opens it for editing", await app.waitFor(`!!document.querySelector(".eq-editor")`, 2000));
  await app.key("Escape", 27);
  await app.key("End", 35);
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
  check("a long document is shown on several pages", report.pages >= 3 && (await app.evaluate(`document.querySelector(".statusbar").textContent`)).includes(`of ${report.pages}`), `${report.pages} pages`);
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
  check("selecting a picture shows its handles, the wrapping choices and Description", !!handle && (await app.count(".picture-bar button")) === 6 && (await app.evaluate(`[...document.querySelectorAll(".picture-bar button")].pop().textContent`)) === "Description…");
  await app.drag(handle[0], handle[1], handle[0] - 97, handle[1] - 40);
  let attrs = await image();
  check("resizing snaps the width to the grid and keeps the proportions", Math.abs(attrs.width / GRID - Math.round(attrs.width / GRID)) < 0.01 && Math.abs(attrs.height / attrs.width - 0.75) < 0.01, JSON.stringify(attrs));

  // Text flows around it.
  await app.clickText(".picture-bar button", "Square");
  attrs = await image();
  check("Square makes it float with the text around it", attrs.floating && attrs.wrap === "square", JSON.stringify(attrs));
  // The picture settles into place just after the layout.
  const beside = await app.waitFor(`(() => {
    const img = document.querySelector(".picture-box img").getBoundingClientRect();
    const p = document.querySelector(".doc > .para");
    const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
    for (let t = w.nextNode(); t; t = w.nextNode()) {
      const r = document.createRange(); r.selectNodeContents(t);
      for (const x of r.getClientRects()) if (x.top < img.bottom - 5 && x.bottom > img.top + 5 && x.left >= img.right) return true;
    }
    return false;
  })()`, 3000);
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

  // Copying a picture puts the picture itself on the clipboard, which other programs (and DuckWriter) can paste.
  await app.send("Browser.grantPermissions", { permissions: ["clipboardReadWrite", "clipboardSanitizedWrite"] });
  const picture = await app.center(".picture-box img");
  await app.click(picture[0], picture[1]);
  check("clicking a picture selects it", (await app.evaluate(`window.__dw.body.state.selection.node?.type.name`)) === "image");
  await app.send("Runtime.evaluate", { expression: `document.execCommand("copy")`, userGesture: true });
  const copied = await app.evaluate(`(async () => {
    for (let i = 0; i < 30; i++) {
      const items = await navigator.clipboard.read().catch(() => []);
      const item = items.find((it) => it.types.includes("image/png"));
      if (item) {
        const bitmap = await createImageBitmap(await item.getType("image/png"));
        return bitmap.width + "x" + bitmap.height;
      }
      await new Promise((r) => setTimeout(r, 100));
    }
    return "nothing";
  })()`);
  check("copying a picture puts the picture itself on the clipboard", copied === "400x300", copied);
  // Pasting HTML with a picture in it (as when pictures are copied with text) gives it its own copy.
  await app.key("End", 35, 2);
  const pasted = await app.evaluate(`(async () => {
    const view = window.__dw.body;
    const blob = await (await fetch(document.querySelector(".doc .picture-box img").src)).blob();
    const url = await new Promise((r) => { const f = new FileReader(); f.onload = () => r(f.result); f.readAsDataURL(blob); });
    view.pasteHTML('<p>Copied <img src="' + url + '" style="width:30pt;height:22.5pt"> here</p>');
    let media = null;
    view.state.doc.descendants((n) => { if (n.type.name === "image" && n.attrs.width === 30) media = n.attrs.media; });
    return !!media && window.__dw.ctx.media.get(media)?.bytes.length === new Uint8Array(await blob.arrayBuffer()).length;
  })()`);
  check("pasting HTML with a picture in it keeps the picture", pasted === true);
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

// ---- auto-save and recovery ------------------------------------------------------------------

const autosaved = join(tmpdir(), "duckwriter-smoke-autosave.docx");
copyFileSync(sample("report.docx"), autosaved);

await session(autosaved, 9500, async (app) => {
  await app.evaluate(`localStorage.removeItem("duckwriter.autosave")`);
  const first = await app.center(".doc > .para", 0);
  await app.click(first[0] + first[2] / 2 - 2, first[1]);
  await app.key("End", 35);
  await app.type(" — auto-saved");
  check("auto-save says it is on", (await app.evaluate(`document.querySelector(".save-status").textContent`)) === "Auto-save on");
  check("a few seconds after typing stops, the file is saved by itself", await app.waitFor(`/^Saved \\d/.test(document.querySelector(".save-status").textContent) && !document.title.startsWith("•")`, 12000), await app.evaluate(`document.querySelector(".save-status").textContent`));
  await app.clickOn(".save-status");
  check("clicking the status switches auto-save off", (await app.evaluate(`document.querySelector(".save-status").textContent`)) === "Auto-save off");
  await app.shoot("auto-save");
}, { keepAutoSave: true });

{
  const zip = await JSZip.loadAsync(readFileSync(autosaved));
  const xml = await zip.file("word/document.xml").async("string");
  check("the auto-saved file on disk has the new text", xml.includes("auto-saved"));
  rmSync(autosaved, { force: true });
}

// A new, never-saved document, and then DuckWriter "crashes" (the test kills it).
await session(null, 9501, async (app) => {
  await app.type("Work that must not be lost");
  // Without a file, its recovery copy is kept a few seconds after typing stops.
  let copies = [];
  for (let i = 0; i < 40 && copies.length < 2; i++) {
    await sleep(250);
    try {
      copies = readdirSync(recoveryDir);
    } catch {
      copies = [];
    }
  }
  check("unsaved work is copied to the recovery folder on this computer", copies.some((f) => f.endsWith(".docx")) && copies.some((f) => f.endsWith(".json")), copies.join(", "));
}, { keepAutoSave: true });

// Make the copy look old, as after a real crash some time ago.
for (const f of readdirSync(recoveryDir).filter((f) => f.endsWith(".json"))) {
  const info = JSON.parse(readFileSync(join(recoveryDir, f), "utf8"));
  writeFileSync(join(recoveryDir, f), JSON.stringify({ ...info, time: Date.now() - 10 * 60_000 }));
}

await session(null, 9502, async (app) => {
  check("on the next start, DuckWriter offers the work back", await app.waitFor(`/Restore unsaved work/.test(document.querySelector("dialog[open]")?.textContent ?? "")`, 5000));
  await app.shoot("restore-offer");
  await app.clickOn("dialog[open] .btn.primary");
  check("restoring brings the text back, still to be saved", await app.waitFor(`document.querySelector(".doc").textContent.includes("Work that must not be lost") && document.title.startsWith("•")`, 5000));
  check("the old copy is gone from the recovery folder", await app.waitFor(`true`) && readdirSync(recoveryDir).length === 0, readdirSync(recoveryDir).join(", "));
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

// ---- tables -------------------------------------------------------------------------------------

const withTable = join(tmpdir(), "duckwriter-smoke-table.docx");
rmSync(withTable, { force: true });

await session(null, 9510, async (app) => {
  await app.debugHandle();
  await app.type("Results:");
  await app.key("Enter", 13);
  // Ctrl+Alt+Shift+T: choose 2 rows × 3 columns with the arrow keys.
  await app.key("T", 84, CTRL | 1 | SHIFT);
  check("Ctrl+Alt+Shift+T opens the table grid", await app.waitFor(`document.activeElement?.classList.contains("table-grid-cell")`, 2000));
  await app.key("ArrowRight", 39);
  await app.key("ArrowRight", 39);
  await app.key("ArrowDown", 40);
  await app.key("Enter", 13);
  const shape = () => app.evaluate(`[...document.querySelectorAll(".doc table")].map((t) => [...t.rows].filter((r) => !r.classList.contains("page-gap-row")).map((r) => r.cells.length).join(","))`);
  check("a 2 × 3 table is inserted, with the cursor in the first cell", JSON.stringify(await shape()) === JSON.stringify(["3,3"]) && (await app.evaluate(`!!document.getSelection().anchorNode?.parentElement?.closest("td")`)), JSON.stringify(await shape()));
  check("the table toolbar appears", await app.evaluate(`!document.querySelector(".table-tools").hidden`));
  for (const text of ["Sample", "Mass", "Notes", "A", "1.2"]) {
    await app.type(text);
    await app.key("Tab", 9);
  }
  await app.type("clear");
  await app.key("Tab", 9);
  check("Tab moves from cell to cell, and in the last cell adds a row", JSON.stringify(await shape()) === JSON.stringify(["3,3,3"]), JSON.stringify(await shape()));
  await app.type("B");
  await app.key("ArrowDown", 40, CTRL | 1 | SHIFT);
  check("Ctrl+Alt+Shift+Down adds a row below", JSON.stringify(await shape()) === JSON.stringify(["3,3,3,3"]), JSON.stringify(await shape()));
  // Select "1.2" and "clear" by dragging, and merge them.
  const from = await app.evaluate(`(() => { const r = [...document.querySelectorAll(".doc td")].find((c) => c.textContent === "1.2").getBoundingClientRect(); return [r.left + 8, r.top + r.height / 2]; })()`);
  const to = await app.evaluate(`(() => { const r = [...document.querySelectorAll(".doc td")].find((c) => c.textContent === "clear").getBoundingClientRect(); return [r.left + 8, r.top + r.height / 2]; })()`);
  await app.drag(from[0], from[1], to[0], to[1]);
  check("dragging across cells selects them", (await app.count(".doc .selectedCell")) === 2);
  await app.clickText(".table-tools .tb-text", "Merge");
  check("Merge makes them one cell", JSON.stringify(await shape()) === JSON.stringify(["3,2,3,3"]), JSON.stringify(await shape()));
  check("new tables have lines (Word's Table Grid style)", await app.evaluate(`getComputedStyle(document.querySelector(".doc td")).borderTopStyle === "solid"`));
  await app.shoot("table");
  await app.evaluate(`(async () => {
    const bytes = await window.__dw.build();
    await window.__TAURI_INTERNALS__.invoke("write_file", bytes, { headers: { "x-path": encodeURIComponent(${JSON.stringify(withTable)}) } });
  })()`);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

{
  const zip = await JSZip.loadAsync(readFileSync(withTable));
  const xml = await zip.file("word/document.xml").async("string");
  check("the saved table is a Word table, with the merged cell", xml.includes("<w:tbl>") && xml.includes('<w:gridSpan w:val="2"/>') && xml.includes('<w:tblStyle w:val="TableGrid"/>') && (await zip.file("word/styles.xml").async("string")).includes('w:styleId="TableGrid"'));
  rmSync(withTable, { force: true });
}

await session(sample("table-long.docx"), 9511, async (app) => {
  const report = await app.evaluate(`(() => {
    const sheets = [...document.querySelectorAll(".page-sheet")].map((s) => s.getBoundingClientRect());
    const margin = parseFloat(getComputedStyle(document.querySelector(".doc")).marginTop);
    const inside = (r) => sheets.some((s) => r.top >= s.top + margin - 3 && r.bottom <= s.bottom - margin + 3);
    const rows = [...document.querySelectorAll(".doc table tr:not(.page-gap-row):not(.repeated-header)")];
    const repeated = document.querySelector(".doc tr.repeated-header");
    const second = sheets[1];
    return {
      pages: sheets.length,
      rows: rows.length,
      outside: [...rows, ...document.querySelectorAll(".repeated-header")].filter((r) => !inside(r.getBoundingClientRect())).length,
      gaps: document.querySelectorAll(".page-gap-row").length,
      header: repeated?.textContent ?? "",
      headerOnTop: !!repeated && Math.abs(repeated.getBoundingClientRect().top - (second.top + margin)) < 3 && repeated.previousElementSibling?.classList.contains("page-gap-row"),
    };
  })()`);
  check("a long table continues on the next page between rows", report.pages >= 2 && report.gaps >= 1 && report.rows === 71 && report.outside === 0, JSON.stringify(report));
  check("its header row is repeated at the top of the next page", report.header === "RunTemperature (°C)Rate (mol/s)" && report.headerOnTop, JSON.stringify(report));
  await app.evaluate(`document.querySelector(".page-gap-row").scrollIntoView({ block: "center" })`);
  await app.shoot("table-across-pages");
});

// ---- find and replace, spelling -----------------------------------------------------------------

await session(null, 9540, async (app) => {
  await app.debugHandle();
  await app.type("The shade of the sky. Shade is light; the shade changes.");
  await app.key("Enter", 13);
  await app.type("Ths sentense has speling misstakes. Denne setningen har stavefeill.");
  await app.key("Home", 36, CTRL);
  await app.key("F", 70, CTRL);
  check("Ctrl+F opens the find bar with the cursor in it", await app.waitFor(`!document.querySelector(".find-dock").hidden && document.activeElement?.classList.contains("find-input")`, 2000));
  await app.type("shade");
  const count = () => app.evaluate(`document.querySelector(".find-count").textContent`);
  check("every match is highlighted and counted", (await app.count(".doc .find-match")) === 3 && (await count()) === "1 of 3", await count());
  await app.key("Enter", 13);
  await app.key("Enter", 13);
  check("Enter goes to the next match and selects it", (await count()) === "2 of 3" && (await app.evaluate(`(() => { const s = window.__dw.body.state; return s.doc.textBetween(s.selection.from, s.selection.to); })()`)) === "Shade" && (await app.evaluate(`document.querySelector(".find-match.current").textContent`)) === "Shade", await count());
  await app.key("Enter", 13, SHIFT);
  check("Shift+Enter goes back", (await count()) === "1 of 3", await count());
  await app.shoot("find");
  await app.key("H", 72, CTRL);
  check("Ctrl+H shows the replace box", await app.waitFor(`document.activeElement?.placeholder === "Replace with"`, 2000));
  await app.type("tint");
  await app.key("Enter", 13);
  const text = () => app.evaluate(`document.querySelector(".doc p").textContent`);
  check("Enter in the replace box replaces the match and goes to the next", (await text()) === "The tint of the sky. Shade is light; the shade changes." && (await count()) === "1 of 2", `${await text()} / ${await count()}`);
  await app.clickText(".find-bar .btn", "Replace all");
  check("Replace all replaces the rest", (await text()) === "The tint of the sky. tint is light; the tint changes." && (await count()) === "Not found", `${await text()} / ${await count()}`);
  await app.key("Z", 90, CTRL);
  await app.key("Escape", 27);
  check("Esc closes the bar, back in the text, highlights gone", (await app.evaluate(`document.querySelector(".find-dock").hidden && !!document.activeElement?.closest(".doc")`)) && (await app.count(".doc .find-match")) === 0);
  await app.key("F3", 114);
  check("F3 opens it again", await app.evaluate(`!document.querySelector(".find-dock").hidden`));
  await app.key("Escape", 27);
  // Spelling, with Windows' own spell checker in English and Norwegian at once.
  const marked = () => app.evaluate(`[...document.querySelectorAll(".doc .spell-error")].map((e) => e.textContent)`);
  await app.waitFor(`document.querySelectorAll(".doc .spell-error").length >= 5`, 8000);
  check("misspelt words get a red line, in English and in Norwegian, and right ones don't", JSON.stringify(await marked()) === JSON.stringify(["Ths", "sentense", "speling", "misstakes", "stavefeill"]), JSON.stringify(await marked()));
  check("the status bar says which languages", /English.*Norwegian|Norwegian.*English|norsk/i.test(await app.evaluate(`document.querySelector(".spell-status").title`)), await app.evaluate(`document.querySelector(".spell-status").title`));
  await app.key("Home", 36, CTRL);
  await app.key("F7", 118);
  check("F7 selects the first mistake and offers suggestions", await app.waitFor(`document.activeElement?.classList.contains("spell-suggestion")`, 5000));
  const firstSuggestion = await app.evaluate(`document.activeElement.textContent`);
  await app.shoot("spelling-suggestions");
  await app.key("Enter", 13);
  const second = () => app.evaluate(`document.querySelectorAll(".doc p")[1].textContent`);
  check("Enter puts the suggestion in", (await second()).startsWith(firstSuggestion + " sentense") && (await marked()).length === 4, `${firstSuggestion}: ${await second()}`);
  const at = await app.evaluate(`(() => { const r = [...document.querySelectorAll(".doc .spell-error")].find((e) => e.textContent === "stavefeill").getBoundingClientRect(); return [r.left + 10, r.top + r.height / 2]; })()`);
  await app.mouse("mouseMoved", at[0], at[1], { buttons: 0, button: "none" });
  await app.mouse("mousePressed", at[0], at[1], { button: "right", buttons: 2 });
  await app.mouse("mouseReleased", at[0], at[1], { button: "right", buttons: 0 });
  check("right-clicking a marked word opens its menu", await app.waitFor(`!!document.querySelector(".spell-menu .menu-item")`, 5000));
  await app.clickText(".spell-menu .menu-item", "Add to dictionary");
  check("Add to dictionary takes the line away", JSON.stringify(await marked()) === JSON.stringify(["sentense", "speling", "misstakes"]), JSON.stringify(await marked()));
  await app.key("F7", 118, CTRL);
  check("Ctrl+F7 switches spell checking off", (await marked()).length === 0 && (await app.evaluate(`document.querySelector(".spell-status").textContent`)) === "Spelling off");
  await app.key("F7", 118, CTRL);
  check("and on again", (await marked()).length === 3);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

// ---- everyday helpers: recent files, page setup, headings, paste, format painter, AutoFormat, counts --

await session(sample("long.docx"), 9570, async (app) => {
  await app.debugHandle();
  // Opened at startup, so it is in the recent list.
  await app.key("F", 70, ALT);
  check("Alt+F lists the recent documents", await app.waitFor(`[...document.querySelectorAll(".recent-menu .recent-name")].some((n) => n.textContent === "long.docx")`, 3000));
  await app.key("Escape", 27);

  await app.key("N", 78, ALT | SHIFT);
  const navItems = () => app.evaluate(`[...document.querySelectorAll(".nav-pane .nav-item")].map((i) => i.textContent)`);
  check("Alt+Shift+N shows the headings beside the page", await app.waitFor(`!document.querySelector(".nav-pane").hidden && document.querySelectorAll(".nav-item").length >= 6`, 3000), JSON.stringify(await navItems()));
  await app.key("ArrowDown", 40);
  await app.key("ArrowDown", 40);
  const target = await app.evaluate(`document.activeElement.textContent`);
  await app.key("Enter", 13);
  const heading = () => app.evaluate(`(() => { const s = window.__dw.body.state; return s.selection.$from.parent.textContent; })()`);
  check("Enter on a heading goes there", (await heading()) === target && (await app.evaluate(`document.querySelector(".nav-item.current")?.textContent`)) === target, `${target} / ${await heading()}`);
  await app.shoot("navigation");
  await app.key("N", 78, ALT | SHIFT);

  await app.key("P", 80, ALT);
  check("Alt+P opens page setup", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Page setup"`, 3000));
  await app.evaluate(`(() => { const s = document.querySelectorAll("dialog[open] select")[1]; s.value = "landscape"; })()`);
  await app.clickText("dialog[open] .btn", "OK");
  check("landscape turns the pages sideways", await app.waitFor(`(() => { const r = document.querySelector(".page-sheet").getBoundingClientRect(); return r.width > r.height; })()`, 5000));
  await app.key("Z", 90, CTRL); // nothing to undo in the text; the page stays sideways
  check("the document counts as changed", await app.evaluate(`document.title.startsWith("•")`));
});

await session(null, 9571, async (app) => {
  await app.debugHandle();
  // AutoFormat: "1. " starts a list; an address becomes a link.
  await app.type("1.");
  await app.type(" ");
  await app.type("First point");
  const first = () => app.evaluate(`(() => { const p = window.__dw.body.state.doc.firstChild; return [p.textContent, p.attrs.numId !== null, document.querySelector(".doc .list-label")?.textContent ?? ""]; })()`);
  check("typing 1. and a space starts a numbered list", JSON.stringify(await first()) === JSON.stringify(["First point", true, "1."]), JSON.stringify(await first()));
  await app.key("Enter", 13);
  await app.key("Enter", 13);
  await app.type("Read www.duck.no");
  await app.type(" ");
  await app.type("today.");
  check("a web address becomes a link when it's finished, and what follows isn't part of it", (await app.evaluate(`document.querySelector(".doc a")?.getAttribute("href")`)) === "https://www.duck.no" && (await app.evaluate(`[...document.querySelectorAll(".doc a")].map((a) => a.textContent).join("|")`)) === "www.duck.no" && (await app.evaluate(`document.querySelectorAll('.doc [data-cstyle="Hyperlink"]').length`)) === 1, await app.evaluate(`document.querySelectorAll(".doc p")[1].innerHTML`));

  // Format painter: bold copied onto other text.
  await app.key("Enter", 13);
  await app.type("Bold and plain");
  const select = (word) => app.evaluate(`(() => { const v = window.__dw.body; let at = -1; v.state.doc.descendants((n, pos) => { if (at < 0 && n.isText && n.text.includes(${JSON.stringify(word)})) at = pos + n.text.indexOf(${JSON.stringify(word)}); }); const TS = v.state.selection.constructor; v.dispatch(v.state.tr.setSelection(TS.create(v.state.doc, at, at + ${word.length}))); v.focus(); })()`);
  await select("Bold");
  await app.key("B", 66, CTRL);
  await app.key("C", 67, CTRL | SHIFT);
  await select("plain");
  await app.key("V", 86, CTRL | SHIFT);
  const boldWords = () => app.evaluate(`[...document.querySelectorAll('.doc [data-pm-mark="bold"]')].map((e) => e.textContent).join(",")`);
  check("Ctrl+Shift+C, then Ctrl+Shift+V copies formatting", (await boldWords()) === "Bold,plain", await boldWords());

  // Pasting a list from a web page.
  await app.key("End", 35, CTRL);
  await app.key("Enter", 13);
  await app.evaluate(`(() => {
    const data = new DataTransfer();
    data.setData("text/html", "<ul><li>Apples</li><li>Pears<ul><li>Green</li></ul></li></ul>");
    data.setData("text/plain", "Apples Pears Green");
    window.__dw.body.dom.dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }));
  })()`);
  const pasted = () => app.evaluate(`[...document.querySelectorAll(".doc p")].slice(-3).map((p) => (p.querySelector(".list-label") ? "•" : "") + p.textContent.replace(/^[•◦▪\\-–]\\s*/, "")).join("|")`);
  check("a list pasted from the web becomes a list", (await app.count(".doc .list-label")) >= 4 && (await pasted()).includes("Green"), await pasted());

  // Word count of a selection.
  await select(" and ");
  check("the status bar counts the selection", /^1 of \d+ words$/.test(await app.evaluate(`document.querySelector(".status-button.words").textContent`)), await app.evaluate(`document.querySelector(".status-button.words").textContent`));
  await app.key("G", 71, CTRL | SHIFT);
  check("Ctrl+Shift+G shows the word count", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Word count" && document.querySelectorAll("dialog[open] .count-table thead th").length === 3`, 3000));
  await app.shoot("word-count");
  await app.key("Escape", 27);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

await session(sample("report.docx"), 9573, async (app) => {
  const pages = async () => (await (await fetch("http://127.0.0.1:9573/json")).json()).filter((t) => t.type === "page");
  await app.key("N", 78, CTRL);
  let found = false;
  for (let i = 0; i < 30 && !found; i++) {
    await sleep(300);
    found = (await pages()).length === 2;
  }
  check("Ctrl+N opens a new document in a window of its own, leaving this one", found && (await app.evaluate(`document.title.startsWith("report.docx")`)));
});

await session(null, 9574, async (app) => {
  await app.debugHandle();
  await app.type("Ducks swim faster in warm water");
  await app.key("S", 83, ALT | SHIFT);
  check("Alt+Shift+S opens the sources", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Sources"`, 3000));
  await app.evaluate(`(() => {
    const box = document.querySelector('dialog[open] textarea[placeholder^="Paste BibTeX"]');
    box.value = "@article{s, author={Smith, Anna and Jones, Bob}, title={Rates of things}, journal={Journal of Ducks}, year=2020, volume={3}, pages={1--10}}";
  })()`);
  await app.clickText("dialog[open] .btn", "Import");
  check("BibTeX is imported", (await app.count("dialog[open] .source-item")) === 1);
  await app.clickText("dialog[open] .btn", "Done");
  await app.key("S", 83, ALT | SHIFT);
  await app.evaluate(`(() => { const s = document.querySelector('dialog[open] select'); s.value = "APA"; })()`);
  await app.clickText("dialog[open] .btn", "Done");
  await app.key("C", 67, ALT | SHIFT);
  check("Alt+Shift+C asks which source to cite", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Cite"`, 8000));
  await app.type("smith");
  await app.key("Enter", 13);
  const cite = () => app.evaluate(`document.querySelector('.doc .field[data-instr="CITATION"]')?.textContent ?? ""`);
  check("the citation shows in the text, after a space", (await app.waitFor(`document.querySelector('.doc .field[data-instr="CITATION"]')?.textContent === "(Smith & Jones, 2020)"`, 3000)) && (await app.evaluate(`window.__dw.body.state.doc.firstChild.textContent`)) === "Ducks swim faster in warm water ", await cite());
  await app.key("End", 35, CTRL);
  await app.key("B", 66, ALT | SHIFT);
  const entry = () => app.evaluate(`document.querySelector(".doc .bibliography-entry")?.textContent ?? ""`);
  check("Alt+Shift+B adds the reference list", (await entry()) === "Smith, A., & Jones, B. (2020). Rates of things. Journal of Ducks, 3, 1–10.", await entry());
  await app.shoot("citations");
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

await session(sample("letterhead.docx"), 9572, async (app) => {
  check("a logo in the header is shown", await app.waitFor(`!!document.querySelector(".hf-layer img")`, 5000));
  await app.shoot("header-logo");
});

// ---- round 3: pages, line numbers, markup, styles, equations, versions, templates -------------------

const versioned = join(tmpdir(), "duckwriter-smoke-versions.docx");
copyFileSync(sample("long.docx"), versioned);

await session(versioned, 9590, async (app) => {
  await app.debugHandle();
  await app.key("G", 71, CTRL);
  check("Ctrl+G asks which page", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Go to page"`, 3000));
  await app.type("3");
  await app.key("Enter", 13);
  const pageOfCursor = () => app.evaluate(`(() => { const s = window.__dw.body.state; const r = window.__dw.body.coordsAtPos(s.selection.from); const sheets = [...document.querySelectorAll(".page-sheet")].map((p) => p.getBoundingClientRect()); return sheets.findIndex((p) => r.top >= p.top && r.top <= p.bottom) + 1; })()`);
  check("and goes there", (await pageOfCursor()) === 3, String(await pageOfCursor()));
  check("the status bar says which page the cursor is on", /^Page 3 of \d+$/.test(await app.evaluate(`document.querySelector(".statusbar .status-item").textContent`)), await app.evaluate(`document.querySelector(".statusbar .status-item").textContent`));
  await app.key("Z", 90, ALT | SHIFT);
  const zoomed = () => app.evaluate(`(() => { const w = document.querySelector(".workspace"); const s = document.querySelector(".page-sheet").getBoundingClientRect(); return [Math.round(s.width), w.clientWidth - 48]; })()`);
  const [width, room] = await zoomed();
  check("Alt+Shift+Z fits the page to the window's width", Math.abs(width - room) < 12, `${width} vs ${room}`);
  await app.key("Z", 90, ALT | SHIFT);
  check("then the whole page", await app.evaluate(`document.querySelector(".page-sheet").getBoundingClientRect().height <= document.querySelector(".workspace").clientHeight`));
  await app.key("Z", 90, ALT | SHIFT);
  check("then back to 100%", (await app.evaluate(`document.querySelector(".statusbar").textContent`)).includes("100%"));

  // Alt+scroll wheel: the selected words a size bigger per notch.
  await app.evaluate(`(() => {
    const para = document.querySelector(".doc .para");
    para.scrollIntoView({ block: "center" });
    const text = document.createTreeWalker(para, NodeFilter.SHOW_TEXT).nextNode();
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, Math.min(4, text.length));
    window.__dw.body.focus();
    getSelection().removeAllRanges();
    getSelection().addRange(range);
  })()`);
  await sleep(300);
  const sizeBefore = Number(await app.evaluate(`document.querySelector(".size-input").value`));
  const word = await app.center(".doc .para");
  await app.wheel(word[0], word[1], -100, ALT);
  await sleep(200);
  const sizeAfter = Number(await app.evaluate(`document.querySelector(".size-input").value`));
  check("Alt+scroll wheel makes the selected text bigger", sizeBefore > 0 && sizeAfter > sizeBefore, `${sizeBefore} → ${sizeAfter}`);
  await app.wheel(word[0], word[1], 100, ALT);
  await sleep(200);
  check("and smaller again", Number(await app.evaluate(`document.querySelector(".size-input").value`)) === sizeBefore);
  check("without zooming", (await app.evaluate(`document.querySelector(".statusbar").textContent`)).includes("100%"));

  await app.key("L", 76, ALT | SHIFT);
  const numbers = () => app.evaluate(`[...document.querySelectorAll(".line-number")].map((n) => Number(n.textContent))`);
  const first = await numbers();
  check("Alt+Shift+L numbers the lines in the margin", first.length > 100 && first[0] === 1 && first.every((n, i) => n === i + 1), `${first.length} numbers, starting ${first.slice(0, 3)}`);
  await app.evaluate(`document.querySelector(".line-number")?.scrollIntoView({ block: "center" })`);
  await app.shoot("line-numbers");
  await app.key("L", 76, ALT | SHIFT);
  const perPage = await numbers();
  check("then from 1 on each page", perPage.filter((n) => n === 1).length === (await app.count(".page-sheet")));
  await app.key("L", 76, ALT | SHIFT);
  check("then none", (await app.count(".line-number")) === 0);

  // A page border: a line round every page, 24 pt in from the edge.
  await app.evaluate(`window.__dw.setPageBorder("double")`);
  const border = await app.evaluate(`(() => { const s = document.querySelector(".page-sheet"); const a = getComputedStyle(s, "::after"); return [s.classList.contains("bordered"), a.borderTopStyle, a.top, document.querySelectorAll(".page-sheet.bordered").length === document.querySelectorAll(".page-sheet").length]; })()`);
  check("a page border goes round every page", border[0] && border[1] === "double" && border[2] === "32px" && border[3], JSON.stringify(border));
  await app.shoot("page-border");
  await app.evaluate(`window.__dw.setPageBorder(null)`);
  check("and comes off again", (await app.count(".page-sheet.bordered")) === 0);
  await app.evaluate(`window.__dw.setWatermark("DRAFT")`);
  check("a watermark goes across every page", (await app.count('.page-sheet[data-watermark="DRAFT"]')) === (await app.count(".page-sheet")));
  await app.shoot("watermark");
  await app.evaluate(`window.__dw.setWatermark(null)`);
  check("and comes off again", (await app.count(".page-sheet[data-watermark]")) === 0);

  // Insert → Date and time: today, in the chosen form.
  await app.evaluate(`(() => { const v = window.__dw.body; v.dispatch(v.state.tr.setSelection(window.__dw.body.state.selection.constructor.atEnd(v.state.doc))); v.focus(); })()`);
  await app.clickOn('.split-arrow[title^="Today\'s date"]');
  const longDate = await app.evaluate(`new Date().toLocaleDateString(undefined, { day: "numeric", month: "long", year: "numeric" })`);
  await app.clickText(".menu .menu-item", longDate);
  check("Insert → Date and time puts today's date in", await app.waitFor(`window.__dw.body.state.doc.textContent.includes(${JSON.stringify(longDate)})`, 3000), longDate);

  // Page numbers in the footer.
  await app.clickOn('.split-arrow[title^="Header and footer"]');
  await app.clickText(".menu .menu-item", "At the bottom, centred");
  check("page numbers go in the footer: Page 1 of 6", await app.waitFor(`[...document.querySelectorAll(".hf.footer")].some((f) => /Page 1 of \\d/.test(f.textContent))`, 5000));

  // Changing a style changes all its paragraphs.
  await app.clickText(".doc p", "1. Section 1");
  await app.key("Y", 89, ALT | SHIFT);
  check("Alt+Shift+Y opens the style", await app.waitFor(`/^Style: /.test(document.querySelector("dialog[open] h2")?.textContent ?? "")`, 3000));
  await app.evaluate(`(() => { const f = document.querySelector('dialog[open] input[aria-label="Font"]'); f.value = "Georgia"; })()`);
  await app.clickText("dialog[open] .btn", "OK");
  const headingFonts = () => app.evaluate(`[...document.querySelectorAll('.doc p[data-style="Heading1"]')].map((p) => getComputedStyle(p).fontFamily.split(",")[0].replace(/"/g, ""))`);
  check("every Heading 1 is in the new font", (await headingFonts()).length >= 6 && (await headingFonts()).every((f) => f === "Georgia"), JSON.stringify(await headingFonts()));

  // Versions: saved twice, the first file is kept and can be restored.
  await app.key("S", 83, CTRL);
  await app.waitFor(`/Saved/.test(document.querySelector(".toasts").textContent)`, 8000);
  await app.key("Home", 36, CTRL);
  await app.type("Second version. ");
  await app.key("S", 83, CTRL);
  await sleep(1500);
  await app.key("V", 86, ALT | SHIFT);
  check("Alt+Shift+V lists the earlier versions", await app.waitFor(`document.querySelectorAll("dialog[open] .source-item").length === 2`, 5000), String(await app.count("dialog[open] .source-item")));
  // Comparing with the newest kept one (from before "Second version" was typed) shows that as added.
  await app.clickText("dialog[open] .source-item .btn", "Compare");
  check("Compare shows what changed since a version, as tracked changes", await app.waitFor(`[...document.querySelectorAll(".doc ins.tc-ins")].some((i) => i.textContent.includes("Second version")) && /differences? from/.test(document.querySelector(".toasts").textContent)`, 5000), await app.evaluate(`document.querySelector(".toasts").textContent`));
  await app.key("V", 86, ALT | SHIFT);
  // Saved under its own name? No: the comparison is a new, unsaved document; versions belong to the file.
  check("the comparison is a document of its own", await app.waitFor(`document.title.includes("(compared)")`, 3000));
  await app.key("Escape", 27);
  await app.evaluate(`window.__dw.openPath(${JSON.stringify(versioned)})`);
  await app.waitFor(`document.title.startsWith("duckwriter-smoke-versions.docx")`, 5000);
  await app.key("V", 86, ALT | SHIFT);
  await app.waitFor(`document.querySelectorAll("dialog[open] .source-item").length === 2`, 5000);
  await app.clickText("dialog[open] .source-item .btn", "Restore");
  check("restoring one shows it, unsaved", await app.waitFor(`!window.__dw.body.state.doc.textContent.startsWith("Second version") && document.title.startsWith("•")`, 5000));
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});
rmSync(versioned, { force: true });

const round3Profile = mkdtempSync(join(tmpdir(), "duckwriter-smoke-r3-"));
await session(null, 9591, async (app) => {
  await app.evaluate(`localStorage.getItem("duckwriter.author") || localStorage.setItem("duckwriter.author", "Kari Nordmann")`);
  await app.debugHandle();
  // A template, in this unused window.
  await app.key("T", 84, ALT | SHIFT);
  await app.clickText(".menu .menu-item", "Scientific article");
  check("Alt+Shift+T: a scientific article from the template", await app.waitFor(`document.querySelector('.doc p[data-style="Title"]')?.textContent === "Title of the article" && !!document.querySelector(".doc .bibliography")`, 5000));
  await app.shoot("template");
  // A numbered equation.
  await app.clickText(".doc p", "Describe the analysis");
  await app.key("End", 35);
  await app.key("Enter", 13);
  await app.key("E", 69, ALT | SHIFT);
  check("Alt+Shift+E: a numbered equation, writing it straight away", await app.waitFor(`!!document.querySelector(".eq-editor") && document.activeElement?.classList.contains("eq-input")`, 3000));
  await app.type("E = mc^2");
  await app.key("Enter", 13);
  const number = () => app.evaluate(`[...document.querySelectorAll(".doc td")].map((c) => c.textContent).join("|")`);
  check("it is numbered (1) on the right", await app.waitFor(`[...document.querySelectorAll(".doc td")].some((c) => c.textContent === "(1)")`, 3000), await number());
  // Show the text as if the changes were accepted.
  await app.key("E", 69, CTRL | SHIFT);
  await app.clickText(".doc p", "Funding and help");
  await app.key("End", 35);
  await app.type(" Thanks to all.");
  const deleted = await app.evaluate(`document.querySelectorAll(".doc ins.tc-ins").length`);
  await app.key("M", 77, ALT | SHIFT);
  check("Alt+Shift+M shows the final text, without markup", deleted > 0 && (await app.evaluate(`getComputedStyle(document.querySelector(".doc ins.tc-ins")).textDecorationLine`)) === "none");
  await app.key("M", 77, ALT | SHIFT);
  check("and the original, without what was added", (await app.evaluate(`getComputedStyle(document.querySelector(".doc ins.tc-ins")).display`)) === "none");
  await app.key("M", 77, ALT | SHIFT);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
}, { profileDir: round3Profile });
removeFolder(round3Profile);

// ---- round 4: ribbon, maths shorthand, focus mode, tables, Markdown and HTML -----------------------

const markdownIn = join(tmpdir(), "duckwriter-smoke-notes.md");
const htmlOut = join(tmpdir(), "duckwriter-smoke-notes.html");
writeFileSync(markdownIn, "# Lab notes\n\n## Results\n\nThe angle was measured.\n\n- first\n- second\n\n| Run | Rate |\n| --- | --- |\n| 2 | 0.7 |\n| 1 | 0.5 |\n");
rmSync(htmlOut, { force: true });

await session(markdownIn, 9592, async (app) => {
  await app.debugHandle();
  const styles = () => app.evaluate(`[...document.querySelectorAll(".doc > p")].map((p) => p.dataset.style ?? "").join(",")`);
  check("a Markdown file opens as a document, with its title, headings, list and table", await app.waitFor(`document.querySelector('.doc p[data-style="Title"]')?.textContent === "Lab notes" && !!document.querySelector(".doc table") && document.querySelectorAll(".doc .list-label").length === 2`, 5000), await styles());
  check("and is a new Word document", await app.evaluate(`document.title.includes("duckwriter-smoke-notes.docx")`));

  // Maths shorthand while typing.
  await app.clickText(".doc p", "The angle");
  await app.key("End", 35);
  await app.type(" It was \\alpha");
  await app.type(" ");
  await app.type("->");
  await app.type(" ");
  await app.type("done ");
  const sentence = () => app.evaluate(`[...document.querySelectorAll(".doc p")].find((p) => p.textContent.startsWith("The angle")).textContent`);
  check("\\alpha and -> become α and → as you type", (await sentence()) === "The angle was measured. It was α → done ", await sentence());

  // Sorting the table by its first column.
  // The cursor in the first column's "2".
  await app.evaluate(`(() => { const v = window.__dw.body; let at = -1; v.state.doc.descendants((n, p) => { if (at < 0 && n.isText && n.text === "2") at = p; }); v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, at))); v.focus(); })()`);
  await app.key("O", 79, ALT | SHIFT);
  const runs = () => app.evaluate(`[...document.querySelectorAll(".doc table tr")].map((r) => r.cells[0].textContent).join(",")`);
  check("Alt+Shift+O sorts the table by the cursor's column", (await runs()) === "Run,1,2", await runs());

  // The ribbon: Alt+3 shows References.
  await app.key("3", 51, ALT);
  check("Alt+3 shows the References tab", await app.evaluate(`document.querySelector(".ribbon-tab.active").textContent === "References" && !!document.querySelector('.ribbon-panel [data-tab="references"]:not([hidden]) .tb-text')`));
  await app.key("1", 49, ALT);

  // Focus mode and the dark page.
  await app.key("F11", 122);
  check("F11: focus mode hides everything but the page", await app.evaluate(`getComputedStyle(document.querySelector(".toolbar")).display === "none" && getComputedStyle(document.querySelector(".statusbar")).display === "none"`));
  await app.key("Escape", 27);
  check("Esc leaves it", await app.evaluate(`getComputedStyle(document.querySelector(".toolbar")).display !== "none"`));
  await app.key("D", 68, ALT | SHIFT);
  check("Alt+Shift+D: a dark page", await app.evaluate(`document.querySelector(".sheet").classList.contains("dark-page") && getComputedStyle(document.querySelector(".doc")).filter.includes("invert")`));
  await app.shoot("dark-page");
  await app.key("D", 68, ALT | SHIFT);

  // An endnote.
  await app.clickText(".doc p", "The angle");
  await app.key("End", 35);
  await app.key("D", 68, CTRL | ALT);
  check("Ctrl+Alt+D adds an endnote and opens it", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Endnote i" && !!document.activeElement?.closest(".endnote-editor")`, 3000));
  await app.type("Measured twice.");
  await app.key("Escape", 27);
  check("the endnote is listed at the end, numbered i", await app.waitFor(`document.querySelector(".doc .endnote-item")?.textContent === "iMeasured twice."`, 3000), await app.evaluate(`document.querySelector(".doc .endnotes")?.textContent ?? ""`));
  await app.shoot("endnote");

  // A copy as a web page.
  await app.evaluate(`window.__dw.exportTo("html", ${JSON.stringify(htmlOut)})`);
  check("saving a copy as a web page", await app.waitFor(`/Saved duckwriter-smoke-notes.html/.test(document.querySelector(".toasts").textContent)`, 5000));
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});
{
  const html = readFileSync(htmlOut, "utf8");
  check("the web page has the title, the list and the table", html.includes("<h1>Lab notes</h1>") && html.includes("<li>first</li>") && /<table><tr><th>.*Run/.test(html));
}
rmSync(markdownIn, { force: true });
rmSync(htmlOut, { force: true });

// ---- writing together: two DuckWriters, three ways of connecting --------------------------------

const relayPort = 8799;
const relay = spawn(process.execPath, [join(root, "server", "relay.mjs")], { env: { ...process.env, PORT: String(relayPort) }, stdio: "ignore" });
const liveFolder = mkdtempSync(join(tmpdir(), "duckwriter-live-"));
await sleep(800);

for (const [method, place, ports] of [
  ["lan", "", [9580, 9581]],
  ["relay", `ws://127.0.0.1:${relayPort}`, [9582, 9583]],
  ["folder", liveFolder, [9584, 9585]],
]) {
  // Profiles of their own: the names they type stay out of the other tests.
  const kariProfile = mkdtempSync(join(tmpdir(), "duckwriter-smoke-a-"));
  const otherProfile = mkdtempSync(join(tmpdir(), "duckwriter-smoke-b-"));
  await session(null, ports[0], async (kari) => {
    await kari.evaluate(`localStorage.setItem("duckwriter.author", "Kari")`);
    await kari.debugHandle();
    await kari.type("Hello");
    if (method === "lan") {
      // The dialog: this network first and chosen; a shared folder under Other ways; no relay.
      await kari.key("W", 87, ALT | SHIFT);
      await kari.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Write together"`, 3000);
      const offered = await kari.evaluate(`(() => {
        const options = [...document.querySelectorAll('dialog[open] input[name="share-method"]')];
        const more = document.querySelector("dialog[open] .share-more");
        return options.map((o) => o.value + (o.checked ? "*" : "") + (more?.contains(o) ? "(more)" : "")).join(",") + (more?.open ? " open" : " closed");
      })()`);
      check("Write together offers this network first, and a shared folder under Other ways", offered === "lan*,folder(more) closed", offered);
      check("joining asks for a folder only for a folder's invitation", await kari.evaluate(`getComputedStyle(document.querySelector('dialog[open] [aria-label="Shared folder"][placeholder^="The shared folder"]').parentElement).display === "none"`));
      await kari.shoot("write-together-dialog");
      await kari.clickText("dialog[open] .dialog-footer .btn", "Close");
    }
    const invite = await kari.evaluate(`window.__dw.shareBy(${JSON.stringify(method)}, ${JSON.stringify(place)})`);
    check(`${method}: sharing gives an invitation`, typeof invite === "string" && invite.startsWith("DW1."), String(invite));
    await session(null, ports[1], async (ola) => {
      await ola.evaluate(`localStorage.setItem("duckwriter.author", "Ola")`);
      await ola.debugHandle();
      await ola.evaluate(`window.__dw.joinBy(${JSON.stringify(invite)}, ${JSON.stringify(method === "folder" ? liveFolder : "")})`);
      const textOf = (app) => app.evaluate(`window.__dw.body.state.doc.textContent`);
      check(`${method}: joining opens the shared document`, (await textOf(ola)) === "Hello" && (await ola.evaluate(`document.title.includes("(shared)")`)), await textOf(ola));
      // Both type at once.
      await kari.key("End", 35, CTRL);
      await ola.key("Home", 36, CTRL);
      await kari.type(", world");
      await ola.type("Dear Kari: ");
      const same = async () => (await textOf(kari)) === "Dear Kari: Hello, world" && (await textOf(ola)) === "Dear Kari: Hello, world";
      let ok = false;
      for (let i = 0; i < 40 && !ok; i++) {
        await sleep(250);
        ok = await same();
      }
      check(`${method}: what both type at the same time ends up the same for both`, ok, `${await textOf(kari)} | ${await textOf(ola)}`);
      const caret = await kari.waitFor(`[...document.querySelectorAll(".doc .peer-label")].some((l) => l.textContent === "Ola")`, 10000);
      check(`${method}: each sees where the other's cursor is, with their name`, caret && (await ola.waitFor(`[...document.querySelectorAll(".doc .peer-label")].some((l) => l.textContent === "Kari")`, 10000)));
      check(`${method}: the status bar says who is there`, /Sharing · 2/.test(await kari.evaluate(`document.querySelector(".share-status").textContent`)));
      if (method === "lan") await kari.shoot("writing-together");
      await kari.evaluate(`window.__dw.collab.stop()`);
      check(`${method}: when the sharer stops, the other keeps their copy`, await ola.waitFor(`document.querySelector(".share-status").hidden && window.__dw.body.state.doc.textContent === "Dear Kari: Hello, world"`, 8000));
      await ola.waitFor(`!document.querySelector(".ended-bar").hidden`, 5000);
      const ended = await ola.evaluate(`(() => { const bar = document.querySelector(".ended-bar"); return bar.hidden ? "hidden: " + bar.textContent : bar.textContent; })()`);
      check(`${method}: and a bar tells them it has ended, and who ended it`, ended.includes("Writing together has ended") && ended.includes("Kari stopped sharing") && ended.includes("Save a copy"), ended);
      check(`${method}: the sharer sees it has ended too`, (await kari.evaluate(`document.querySelector(".ended-bar").textContent`)).includes("You stopped sharing"));
      check(`${method}: no errors in either app`, kari.errors.length === 0 && ola.errors.length === 0, [...kari.errors, ...ola.errors].join("; "));
    }, { profileDir: otherProfile });
  }, { profileDir: kariProfile });
  removeFolder(kariProfile);
  removeFolder(otherProfile);
}
relay.kill();
removeFolder(liveFolder);

// ---- check before sending -----------------------------------------------------------------------

await session(sample("review.docx"), 9596, async (app) => {
  await app.debugHandle();
  await app.key("I", 73, ALT | SHIFT);
  const texts = () => app.evaluate(`[...document.querySelectorAll("dialog[open] .inspect-text")].map((t) => t.textContent)`);
  check("Alt+Shift+I lists what to look at before sending", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Check before sending"`, 3000) && (await texts()).some((t) => /4 tracked changes are still there/.test(t)) && (await texts()).some((t) => /2 comments are not resolved/.test(t)), JSON.stringify(await texts()));
  await app.clickText("dialog[open] .inspect-item .btn", "Show (1 of 4)");
  const selected = () => app.evaluate(`(() => { const s = window.__dw.body.state; return s.selection.from; })()`);
  const first = await selected();
  await app.clickText("dialog[open] .inspect-item .btn", "Show (2 of 4)");
  check("Show goes from one to the next", (await selected()) > first, `${first} → ${await selected()}`);
  await app.shoot("check-before-sending");
  await app.clickText("dialog[open] .btn", "Close");
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

// ---- round 6: sections, notes, word goal, LaTeX -----------------------------------------------

const paperIn = join(tmpdir(), "duckwriter-smoke-paper.md");
const texOut = join(tmpdir(), "duckwriter-smoke-paper.tex");
writeFileSync(paperIn, "# A paper\n\n## Intro\n\nIntro text.\n\n## Methods\n\nMethods text.\n\n## Results\n\nResults text at 10 kg.\n");
rmSync(texOut, { force: true });
const round6Profile = join(tmpdir(), "duckwriter-smoke-round6-profile");
removeFolder(round6Profile);

await session(paperIn, 9597, async (app) => {
  await app.debugHandle();
  await app.waitFor(`[...document.querySelectorAll(".doc p")].some((p) => p.textContent === "Results")`, 5000);
  const headings = () => app.evaluate(`[...document.querySelectorAll('.doc > p[data-style^="Heading"]')].map((p) => p.textContent + ":" + p.dataset.style).join(",")`);
  const order = () => app.evaluate(`[...document.querySelectorAll(".doc > p")].map((p) => p.textContent).filter((t) => t).join("|")`);

  await app.clickText(".doc p", "Results text");
  // The editor takes the click's cursor a moment later.
  await app.waitFor(`window.__dw.body.state.selection.$from.parent.textContent.startsWith("Results text")`, 3000);
  await app.key("ArrowUp", 38, ALT | SHIFT);
  check("Alt+Shift+Up moves the section (heading and text) above the one before", (await order()) === "A paper|Intro|Intro text.|Results|Results text at 10 kg.|Methods|Methods text.", await order());
  await app.key("ArrowRight", 39, ALT | SHIFT);
  check("Alt+Shift+Right takes its heading a level down", (await headings()).includes("Results:Heading2"), await headings());
  await app.key("ArrowLeft", 37, ALT | SHIFT);
  check("and Alt+Shift+Left back up", (await headings()).includes("Results:Heading1"), await headings());

  // Dragging Methods before Intro in the headings pane.
  await app.key("N", 78, ALT | SHIFT);
  await app.waitFor(`document.querySelectorAll(".nav-item").length >= 3`, 3000);
  await app.evaluate(`(() => {
    const items = [...document.querySelectorAll(".nav-item")];
    const from = items.find((i) => i.textContent === "Methods");
    const to = items.find((i) => i.textContent === "Intro");
    const data = new DataTransfer();
    const r = to.getBoundingClientRect();
    from.dispatchEvent(new DragEvent("dragstart", { bubbles: true, dataTransfer: data }));
    to.dispatchEvent(new DragEvent("dragover", { bubbles: true, cancelable: true, dataTransfer: data, clientY: r.top + 2 }));
    to.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: data, clientY: r.top + 2 }));
  })()`);
  check("dragging a heading in the headings pane moves its section", (await order()).startsWith("A paper|Methods|Methods text.|Intro"), await order());
  await app.shoot("headings-pane");
  await app.key("N", 78, ALT | SHIFT);

  // A non-breaking space.
  await app.clickText(".doc p", "Results text");
  await app.evaluate(`(() => { const v = window.__dw.body; let at = -1; v.state.doc.descendants((n, p) => { if (at < 0 && n.isText && n.text.includes("10 kg")) at = p + n.text.indexOf("10 kg") + 2; }); v.dispatch(v.state.tr.setSelection(v.state.selection.constructor.create(v.state.doc, at, at + 1))); v.focus(); })()`);
  await app.key(" ", 32, CTRL | SHIFT, "Space");
  check("Ctrl+Shift+Space puts a non-breaking space", await app.evaluate(`window.__dw.body.state.doc.textContent.includes("10" + String.fromCharCode(160) + "kg")`));

  // A footnote, then all footnotes made endnotes.
  await app.key("End", 35);
  await app.key("F", 70, CTRL | ALT);
  await app.waitFor(`!!document.activeElement?.closest(".fn-note")`, 3000);
  await app.type("Weighed twice.");
  await app.key("Escape", 27);
  await app.evaluate(`window.__dw.openMenu("references")`);
  await app.clickText(".menu-item", "Make all footnotes endnotes");
  check("all footnotes become endnotes, listed at the end", await app.waitFor(`/^i ?Weighed twice[.]$/.test(document.querySelector(".doc .endnote-item")?.textContent ?? "") && !document.querySelector(".doc .footnote-ref")`, 3000), await app.evaluate(`document.querySelector(".doc .endnotes")?.textContent ?? ""`));

  // A word goal.
  await app.key("G", 71, CTRL | SHIFT);
  await app.waitFor(`!!document.querySelector("dialog[open] .goal-input")`, 3000);
  await app.evaluate(`(() => { const box = document.querySelector("dialog[open] .goal-input"); box.value = "1000"; box.dispatchEvent(new Event("input")); })()`);
  await app.shoot("word-goal");
  await app.clickText("dialog[open] .btn", "Close");
  check("a word goal shows in the status bar, with how far along it is", await app.waitFor(`document.querySelector(".status-button.words").classList.contains("has-goal") && / of 1[,.\\s ]?000 words/.test(document.querySelector(".status-button.words").textContent)`, 3000), await app.evaluate(`document.querySelector(".status-button.words").textContent`));

  // A copy as LaTeX.
  await app.evaluate(`window.__dw.exportTo("latex", ${JSON.stringify(texOut)})`);
  check("saving a copy as LaTeX", await app.waitFor(`/Saved duckwriter-smoke-paper.tex/.test(document.querySelector(".toasts").textContent)`, 5000));
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
}, { profileDir: round6Profile });
removeFolder(round6Profile);
{
  const tex = readFileSync(texOut, "utf8");
  check("the LaTeX has the title, the sections in their new order and the endnote", tex.includes("\\title{A paper}") && tex.indexOf("\\section{Methods}") < tex.indexOf("\\section{Intro}") && tex.includes("\\endnote{Weighed twice.}") && tex.includes("10~kg"));
}
rmSync(paperIn, { force: true });
rmSync(texOut, { force: true });

// ---- sections: front matter i, ii, then chapters from 1 -----------------------------------------

{
  const docx = await import("docx");
  const thesisFile = join(tmpdir(), "duckwriter-smoke-thesis.docx");
  const footer = (name) => new docx.Footer({ children: [new docx.Paragraph({ children: [new docx.TextRun({ children: [`${name} `, docx.PageNumber.CURRENT] })] })] });
  writeFileSync(
    thesisFile,
    await docx.Packer.toBuffer(
      new docx.Document({
        sections: [
          { properties: { page: { pageNumbers: { start: 1, formatType: docx.NumberFormat.LOWER_ROMAN } } }, footers: { default: footer("Front") }, children: [new docx.Paragraph("Abstract")] },
          { properties: { type: docx.SectionType.NEXT_PAGE, page: { pageNumbers: { start: 1, formatType: docx.NumberFormat.DECIMAL } } }, footers: { default: footer("Chapter") }, children: [new docx.Paragraph("Chapter one")] },
        ],
      }),
    ),
  );
  await session(thesisFile, 9599, async (app) => {
    await app.waitFor(`document.querySelectorAll(".page-sheet").length === 2`, 5000);
    const page = (text) => app.evaluate(`(() => { const p = [...document.querySelectorAll(".doc p")].find((p) => p.textContent === ${JSON.stringify(text)}); const pitch = document.querySelectorAll(".page-sheet")[1].getBoundingClientRect().top - document.querySelectorAll(".page-sheet")[0].getBoundingClientRect().top; return Math.floor((p.getBoundingClientRect().top - document.querySelector(".page-sheet").getBoundingClientRect().top) / pitch); })()`);
    check("a section break starts a new page", (await page("Abstract")) === 0 && (await page("Chapter one")) === 1, `${await page("Abstract")}, ${await page("Chapter one")}`);
    const footers = await app.evaluate(`[...document.querySelectorAll(".hf.footer")].map((f) => f.textContent.trim()).join(",")`);
    check("each section has its footer, and its page numbers: i in the front matter, 1 in the first chapter", footers === "Front i,Chapter 1", footers);
    check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
  });
  rmSync(thesisFile, { force: true });

  // A landscape section in the middle, for a wide table.
  const wideFile = join(tmpdir(), "duckwriter-smoke-wide.docx");
  writeFileSync(
    wideFile,
    await docx.Packer.toBuffer(
      new docx.Document({
        sections: [
          { footers: { default: footer("Portrait") }, children: [new docx.Paragraph("Before the table")] },
          { properties: { type: docx.SectionType.NEXT_PAGE, page: { size: { orientation: docx.PageOrientation.LANDSCAPE } } }, footers: { default: footer("Landscape") }, children: [new docx.Paragraph("A wide table goes here, across the whole width of a landscape page.")] },
          { properties: { type: docx.SectionType.NEXT_PAGE }, footers: { default: footer("Portrait") }, children: [new docx.Paragraph("After the table")] },
        ],
      }),
    ),
  );
  await session(wideFile, 9600, async (app) => {
    await app.waitFor(`document.querySelectorAll(".page-sheet").length === 3`, 5000);
    const sizes = await app.evaluate(`[...document.querySelectorAll(".page-sheet")].map((s) => { const r = s.getBoundingClientRect(); return Math.round(r.width) + "x" + Math.round(r.height); }).join(",")`);
    const [a, b, c] = sizes.split(",").map((s) => s.split("x").map(Number));
    check("a landscape section's page is drawn landscape, between portrait ones", a[0] < a[1] && b[0] > b[1] && c[0] < c[1] && Math.abs(b[0] - a[1]) < 3, sizes);
    const widths = await app.evaluate(`["Before the table", "A wide table", "After the table"].map((t) => Math.round([...document.querySelectorAll(".doc p")].find((p) => p.textContent.startsWith(t)).getBoundingClientRect().width)).join(",")`);
    const [w1, w2, w3] = widths.split(",").map(Number);
    check("its text is as wide as the landscape page allows", w2 > w1 * 1.3 && w1 === w3, widths);
    const placed = await app.evaluate(`(() => {
      const sheets = [...document.querySelectorAll(".page-sheet")].map((s) => s.getBoundingClientRect());
      const on = (el) => { const r = el.getBoundingClientRect(); return sheets.findIndex((s) => r.top >= s.top - 1 && r.bottom <= s.bottom + 1 && r.left >= s.left - 1 && r.right <= s.right + 1); };
      const text = ["Before the table", "A wide table", "After the table"].map((t) => on([...document.querySelectorAll(".doc p")].find((p) => p.textContent.startsWith(t))));
      const footers = [...document.querySelectorAll(".hf.footer")].map((f) => on(f) + ":" + f.textContent.trim().split(" ")[0]);
      return text.join(",") + " | " + footers.join(",");
    })()`);
    check("each section's text and footer are on its own page", placed === "0,1,2 | 0:Portrait,1:Landscape,2:Portrait", placed);
    await app.shoot("landscape-section");
    // On paper, one size for all pages: three pages still, and back to landscape on screen afterwards.
    await app.debugHandle();
    const widePdf = join(tmpdir(), "duckwriter-smoke-wide.pdf");
    rmSync(widePdf, { force: true });
    await app.evaluate(`window.__dw.savePdfTo(${JSON.stringify(widePdf)})`);
    await app.waitFor(`/Saved duckwriter-smoke-wide.pdf/.test(document.querySelector(".toasts").textContent)`, 20000);
    const pdfPages = (readFileSync(widePdf).toString("latin1").match(/\/Type\s*\/Page(?![a-z])/g) ?? []).length;
    check("a PDF of it has its three pages", pdfPages === 3, String(pdfPages));
    rmSync(widePdf, { force: true });
    const after = await app.evaluate(`(() => { const r = document.querySelectorAll(".page-sheet")[1].getBoundingClientRect(); return r.width > r.height; })()`);
    check("and the screen shows the landscape page again", after);
    check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
  });
  rmSync(wideFile, { force: true });
}

// ---- round 7: properties, readability, Zotero -------------------------------------------------

const round7Profile = join(tmpdir(), "duckwriter-smoke-round7-profile");
removeFolder(round7Profile);
await session(null, 9598, async (app) => {
  await app.debugHandle();
  await app.clickText(".doc p", "");
  await app.type("The duck swam across the lake. It was cold.");

  // Document properties.
  await app.key("P", 80, ALT | SHIFT);
  check("Alt+Shift+P opens the document's properties", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Document properties"`, 3000));
  await app.evaluate(`(() => {
    const set = (key, value) => { const el = document.querySelector('dialog[open] [data-key="' + key + '"]'); el.value = value; };
    set("title", "Ducks in winter");
    set("creator", "Kari Nordmann");
  })()`);
  await app.shoot("properties");
  await app.clickText("dialog[open] .btn", "OK");
  check("the title and author are kept, to be saved", await app.evaluate(`window.__dw.ctx.properties.title === "Ducks in winter" && window.__dw.ctx.properties.creator === "Kari Nordmann"`));

  // Check before sending names them, and can leave them out.
  await app.key("I", 73, ALT | SHIFT);
  check("Check before sending lists the names in the properties", await app.waitFor(`[...document.querySelectorAll("dialog[open] .inspect-text")].some((t) => t.textContent.includes("Kari Nordmann"))`, 3000));
  await app.clickText("dialog[open] .inspect-item .btn", "Leave out");
  check("Leave out removes them", await app.evaluate(`window.__dw.ctx.properties.anonymous === true && window.__dw.ctx.properties.creator === ""`));
  await app.clickText("dialog[open] .btn", "Close");

  // Readability in the word count.
  await app.key("G", 71, CTRL | SHIFT);
  check("the word count has sentences and LIX", await app.waitFor(`[...document.querySelectorAll("dialog[open] .count-table th")].some((t) => t.textContent === "Readability (LIX)") && [...document.querySelectorAll("dialog[open] .count-table tr")].some((r) => r.textContent === "Sentences2")`, 3000), await app.evaluate(`document.querySelector("dialog[open] .count-table")?.textContent ?? ""`));
  await app.clickText("dialog[open] .btn", "Close");

  // Zotero: results if it runs, otherwise why not.
  await app.key("S", 83, ALT | SHIFT);
  await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Sources"`, 3000);
  await app.evaluate(`(() => { const box = document.querySelector('dialog[open] [aria-label="Search Zotero"]'); box.value = "duck"; })()`);
  await app.clickText("dialog[open] .btn", "Search Zotero");
  check("searching Zotero lists what it has, or says why it can't", await app.waitFor(`!document.querySelector("dialog[open] .zotero-results")?.hidden || /Zotero/.test(document.querySelector("dialog[open] .dialog-error")?.textContent ?? "")`, 20000), await app.evaluate(`document.querySelector("dialog[open] .dialog-error")?.textContent ?? ""`));
  await app.clickText("dialog[open] .btn", "Done");
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
}, { profileDir: round7Profile });
removeFolder(round7Profile);

// ---- links ----------------------------------------------------------------------------------------

await session(null, 9560, async (app) => {
  await app.debugHandle();
  await app.type("Read the report");
  await app.key("ArrowLeft", 37, CTRL | SHIFT);
  await app.key("K", 75, CTRL);
  check("Ctrl+K asks for the address of the selected text", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Insert link" && document.activeElement?.placeholder?.startsWith("https")`, 3000));
  await app.type("duck.no/report");
  await app.key("Enter", 13);
  const link = () => app.evaluate(`(() => { const a = document.querySelector(".doc a"); return a ? a.textContent + " → " + a.getAttribute("href") : ""; })()`);
  check("the text becomes a link, blue and underlined", (await link()) === "report → https://duck.no/report" && (await app.evaluate(`getComputedStyle(document.querySelector('.doc a [data-cstyle="Hyperlink"]') ?? document.body).color`)) === "rgb(5, 99, 193)", `${await link()} ${await app.evaluate(`getComputedStyle(document.querySelector('.doc a [data-cstyle="Hyperlink"]') ?? document.body).color`)}`);
  await app.clickText(".doc a", "report");
  check("a plain click puts the cursor in the link, without following it", (await app.evaluate(`location.href.startsWith("http://tauri.localhost") || location.protocol === "tauri:"`)) && !(await app.evaluate(`!!document.querySelector("dialog[open]")`)));
  await app.key("K", 75, CTRL);
  check("Ctrl+K on a link edits it", await app.waitFor(`document.querySelector("dialog[open] h2")?.textContent === "Edit link"`, 3000));
  await app.shoot("link");
  await app.clickText("dialog[open] .btn", "Remove link");
  check("Remove link leaves the text", (await app.count(".doc a")) === 0 && (await app.evaluate(`document.querySelector(".doc p").textContent`)) === "Read the report");
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

// ---- reviewing: tracked changes and comments ------------------------------------------------------

// The reviewing pane lists every change, to accept or reject one by one.
await session(sample("review.docx"), 9621, async (app) => {
  await app.debugHandle();
  await app.evaluate(`[...document.querySelectorAll(".status-button")].find((b) => b.textContent.includes("Changes")).click()`);
  const items = () => app.evaluate(`[...document.querySelectorAll(".review-item")].map((i) => i.querySelector(".review-who").textContent + ": " + i.querySelector(".review-text").textContent)`);
  check("the reviewing pane lists the tracked changes, with who made them", await app.waitFor(`document.querySelectorAll(".review-item").length === 4`, 3000) && /^4 changes: 2 inserted, 2 deleted/.test(await app.evaluate(`document.querySelector(".review-summary").textContent`)), JSON.stringify(await items()));
  await app.evaluate(`document.querySelector(".review-item .btn").click()`);
  check("accepting one there takes it off the list and out of the text", await app.waitFor(`document.querySelectorAll(".review-item").length === 3`, 3000) && (await app.count(".doc del.tc-del")) + (await app.count(".doc ins.tc-ins")) === 3, JSON.stringify(await items()));
  await app.evaluate(`[...document.querySelectorAll(".status-button")].find((b) => b.textContent.includes("Changes")).click()`);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

const reviewed = join(tmpdir(), "duckwriter-smoke-review.docx");
rmSync(reviewed, { force: true });


await session(sample("review.docx"), 9550, async (app) => {
  await app.debugHandle();
  const text = (i) => app.evaluate(`document.querySelectorAll(".doc p")[${i}].textContent`);
  check("a Word file's tracked changes are shown: deleted text struck through, new text underlined", (await app.count(".doc del.tc-del")) === 2 && (await app.count(".doc ins.tc-ins")) === 2 && (await text(0)) === "The rate wasis proportional to the temperature.");
  const cards = () => app.evaluate(`[...document.querySelectorAll(".comment-card")].map((c) => [...c.querySelectorAll(".comment-text")].map((t) => t.value).join(" / "))`);
  check("its comments are shown beside the page", JSON.stringify(await cards()) === JSON.stringify(["Is this the right figure?", "Please cite the source.\nAny review paper will do."]), JSON.stringify(await cards()));
  check("the comments are right of the page, the first level with its text, none overlapping", await app.evaluate(`(() => {
    const page = document.querySelector(".page-sheet").getBoundingClientRect();
    const cards = [...document.querySelectorAll(".comment-card")].map((c) => c.getBoundingClientRect());
    const first = document.querySelector(".doc .comment-range").getBoundingClientRect();
    return cards.every((r, i) => r.left > page.right && (i === 0 || r.top >= cards[i - 1].bottom)) && Math.abs(cards[0].top - first.top) < 20;
  })()`));
  await app.shoot("review");
  // Alt+PageDown from the top: the first change ("was"), then the next.
  await app.key("Home", 36, CTRL);
  await app.key("PageDown", 34, ALT);
  const selected = () => app.evaluate(`(() => { const s = window.__dw.body.state; return s.doc.textBetween(s.selection.from, s.selection.to); })()`);
  check("Alt+PageDown goes to the first change", (await selected()) === "was", await selected());
  await app.key("D", 68, CTRL | ALT | SHIFT);
  check("Ctrl+Alt+Shift+D rejects it (the deletion is undone) and goes on to the next", (await text(0)) === "The rate wasis proportional to the temperature." && (await app.count(".doc del.tc-del")) === 1 && (await selected()) === "is", `${await text(0)} / ${await selected()}`);
  await app.key("D", 68, CTRL | ALT | SHIFT);
  check("rejecting an insertion removes it", (await text(0)) === "The rate was proportional to the temperature.");
  await app.clickOn('.split-arrow[title^="Accept the change"]');
  await app.clickText(".menu .menu-item", "Accept all changes");
  check("Accept all changes accepts the rest", (await app.count(".doc del.tc-del, .doc ins.tc-ins")) === 0 && (await text(1)) === "As shown in figure 3, the curve clearly flattens out." && (await text(2)) === "Earlier work found the same.");

  // A new comment, by keyboard: the first time it asks for your name.
  await app.clickText(".doc p", "As shown");
  await app.key("End", 35);
  await app.key("ArrowLeft", 37, CTRL);
  await app.key("ArrowLeft", 37, CTRL);
  await app.key("M", 77, CTRL | ALT);
  check("Ctrl+Alt+M asks once for your name", await app.waitFor(`!!document.querySelector("dialog[open] input")`, 3000));
  await app.type("Kari Nordmann");
  await app.key("Enter", 13);
  check("then puts the cursor in a new comment on the word", (await app.waitFor(`document.activeElement?.classList.contains("comment-text")`, 3000)) && (await app.count(".comment-card")) === 3, await app.evaluate(`document.activeElement?.outerHTML.slice(0, 120)`));
  await app.type("Which curve is this?");
  await app.key("Escape", 27);
  check("Esc goes back to the text", await app.evaluate(`!!document.activeElement?.closest(".doc")`));
  check("the comment shows who wrote it", await app.evaluate(`[...document.querySelectorAll(".comment-author")].some((a) => a.textContent === "Kari Nordmann")`));
  await app.shoot("new-comment");

  // Tracking changes while typing.
  await app.key("E", 69, CTRL | SHIFT);
  check("Ctrl+Shift+E starts tracking changes", await app.evaluate(`!document.querySelector(".track-status").hidden && document.querySelector('.tb[title^="Track changes"]').classList.contains("active")`));
  await app.clickText(".doc p", "The rate");
  await app.key("End", 35);
  await app.key("Backspace", 8);
  await app.type(" here.");
  check("typing is marked as new and deleted text stays, struck through", (await text(0)) === "The rate was proportional to the temperature here.." && (await app.count(".doc ins.tc-ins")) === 1 && (await app.count(".doc del.tc-del")) === 1, await text(0));
  await app.shoot("tracking");
  await app.evaluate(`(async () => {
    const bytes = await window.__dw.build();
    await window.__TAURI_INTERNALS__.invoke("write_file", bytes, { headers: { "x-path": encodeURIComponent(${JSON.stringify(reviewed)}) } });
  })()`);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

{
  const zip = await JSZip.loadAsync(readFileSync(reviewed));
  const xml = await zip.file("word/document.xml").async("string");
  const comments = (await zip.file("word/comments.xml")?.async("string")) ?? "";
  const settings = (await zip.file("word/settings.xml")?.async("string")) ?? "";
  check("saved for Word: the new comment, the tracked changes by name, and tracking left on", comments.includes("Which curve is this?") && comments.includes('w:author="Kari Nordmann"') && /<w:ins w:id="\d+" w:author="Kari Nordmann"/.test(xml) && /<w:del [^>]*w:author="Kari Nordmann"[^>]*><w:r>(?:<w:rPr>.*?<\/w:rPr>)?<w:delText xml:space="preserve">\.<\/w:delText>/.test(xml) && settings.includes("<w:trackRevisions/>"));
  rmSync(reviewed, { force: true });
}

// ---- save as PDF -------------------------------------------------------------------------------

const pdfPath = join(tmpdir(), "duckwriter-smoke.pdf");
rmSync(pdfPath, { force: true });

await session(sample("report.docx"), 9532, async (app) => {
  await app.debugHandle();
  await app.evaluate(`window.__dw.savePdfTo(${JSON.stringify(pdfPath)})`);
  await app.waitFor(`/Saved duckwriter-smoke.pdf/.test(document.querySelector(".toasts").textContent)`, 20000);
  const pdf = readFileSync(pdfPath).toString("latin1");
  check("a one-page document makes a one-page PDF, without a blank page after it", (pdf.match(/\/Type\s*\/Page(?![a-z])/g) ?? []).length === 1);
});

await session(sample("long.docx"), 9530, async (app) => {
  await app.debugHandle();
  const pages = await app.count(".page-sheet");
  await app.evaluate(`window.__dw.savePdfTo(${JSON.stringify(pdfPath)})`);
  check("Save as PDF writes a PDF", await app.waitFor(`/Saved duckwriter-smoke.pdf/.test(document.querySelector(".toasts").textContent)`, 20000));
  const pdf = readFileSync(pdfPath).toString("latin1");
  const pdfPages = (pdf.match(/\/Type\s*\/Page(?![a-z])/g) ?? []).length;
  const box = /\/MediaBox\s*\[\s*0\s+0\s+([\d.]+)\s+([\d.]+)/.exec(pdf);
  check("the PDF has the same pages as the screen, at the document's page size (A4)", pdf.startsWith("%PDF") && pdfPages === pages && !!box && Math.abs(Number(box[1]) - 595.3) < 2 && Math.abs(Number(box[2]) - 841.9) < 2, `${pdfPages} pages in the PDF, ${pages} on screen, ${box?.[1]} × ${box?.[2]}`);
  const headings = await app.evaluate(`[...document.querySelectorAll('.doc [data-style^="Heading"]')].map((h) => h.textContent).filter(Boolean)`);
  const outline = await (await pdfjs.getDocument({ data: new Uint8Array(readFileSync(pdfPath)) }).promise).getOutline();
  const flat = (items) => (items ?? []).flatMap((i) => [i.title, ...flat(i.items)]);
  check("the PDF has a bookmark for each heading", headings.length > 0 && JSON.stringify(flat(outline)) === JSON.stringify(headings), `${JSON.stringify(flat(outline)).slice(0, 200)} vs ${JSON.stringify(headings).slice(0, 200)}`);
  check("the screen goes back to normal afterwards", (await app.evaluate(`!document.documentElement.classList.contains("printing") && document.querySelectorAll(".page-sheet").length`)) === pages);
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

// Older formats open as new Word documents.
for (const [ext, port] of [["doc", 9436], ["rtf", 9437], ["odt", 9438]]) {
  await session(join(root, "tests", "fixtures", `formats.${ext}`), port, async (app) => {
    check(`a .${ext} file opens as a new Word document, with its text`, await app.waitFor(`document.title.includes("formats.docx") && document.querySelector('.doc p[data-style="Heading1"]')?.textContent === "Introduction" && !!document.querySelector(".doc table")`, 8000), await app.evaluate(`document.title`));
    check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
  });
}

// A PDF opens as a new Word document, worked out from its text.
await session(sample("lecture.pdf"), 9439, async (app) => {
  check("a PDF opens as a new Word document", await app.waitFor(`document.title.includes("lecture.docx") && !!document.querySelector(".doc table")`, 20000), await app.evaluate(`document.title + " " + (document.querySelector(".toast")?.textContent ?? "")`));
  check("with its heading", (await app.evaluate(`document.querySelector('.doc p[data-style="Heading1"]')?.textContent`)) === "Lecture 3: Cells");
  check("its paragraph joined from two lines", (await app.evaluate(`[...document.querySelectorAll(".doc p")].some((p) => p.textContent === "Cells are the building blocks of life. This line is long enough to wrap onto a second line in the same paragraph.")`)) === true);
  check("its bulleted list", (await app.evaluate(`[...document.querySelectorAll(".doc p.list-item")].filter((p) => p.querySelector(".list-label.bullet") && /(nucleus|membrane)$/.test(p.textContent)).length`)) === 2);
  check("and its picture", await app.waitFor(`[...document.querySelectorAll(".doc img")].some((img) => img.naturalWidth === 2)`, 5000));
  check("no errors in the app", app.errors.length === 0, app.errors.join("; "));
});

await session(sample("not-a-docx.docx"), 9435, async (app) => {
  check("a file that isn't a Word document gets a plain message", await app.waitFor(`/isn't a Word document/.test(document.querySelector(".toast.error")?.textContent ?? "")`, 5000));
});

removeFolder(profile);
if (shotsDir) console.log(`\n${shots} screenshots saved to ${shotsDir}`);
console.log(failures ? `\n${failures} check(s) failed` : "\nAll checks passed");
process.exit(failures ? 1 : 0);
