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
    // Tests edit the sample files; auto-save would change them. One test switches it on.
    if (!keepAutoSave) await app.evaluate(`localStorage.setItem("duckwriter.autosave", "off")`);
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
  const otherProfile = mkdtempSync(join(tmpdir(), "duckwriter-smoke-b-"));
  await session(null, ports[0], async (kari) => {
    await kari.evaluate(`localStorage.setItem("duckwriter.author", "Kari")`);
    await kari.debugHandle();
    await kari.type("Hello");
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
      check(`${method}: no errors in either app`, kari.errors.length === 0 && ola.errors.length === 0, [...kari.errors, ...ola.errors].join("; "));
    }, { profileDir: otherProfile });
  });
  rmSync(otherProfile, { recursive: true, force: true });
}
relay.kill();
rmSync(liveFolder, { recursive: true, force: true });

rmSync(profile, { recursive: true, force: true });
if (shotsDir) console.log(`\n${shots} screenshots saved to ${shotsDir}`);
console.log(failures ? `\n${failures} check(s) failed` : "\nAll checks passed");
process.exit(failures ? 1 : 0);
