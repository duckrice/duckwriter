// File access. Inside the desktop app this goes through native dialogs and the
// Rust side; in a plain browser (used during development) it falls back to a
// file picker and a download.

import { invoke } from "@tauri-apps/api/core";
import { getCurrentWebview } from "@tauri-apps/api/webview";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { open, save } from "@tauri-apps/plugin-dialog";

export const isDesktop = "__TAURI_INTERNALS__" in window;

export interface OpenedFile {
  name: string;
  /** Absolute path; missing in the browser, where files can't be written back. */
  path: string | null;
  bytes: Uint8Array;
}

/** What a document can be saved as: the Word formats, each kept as it is. */
const SAVE_FILTERS = [
  { name: "Word document", extensions: ["docx"] },
  { name: "Word macro-enabled document", extensions: ["docm"] },
  { name: "Word template", extensions: ["dotx"] },
  { name: "Word macro-enabled template", extensions: ["dotm"] },
  { name: "Word XML document", extensions: ["xml"] },
];
const OPEN_FILTERS = [
  { name: "Documents", extensions: ["docx", "docm", "dotx", "dotm", "xml", "doc", "rtf", "odt", "md", "markdown", "html", "htm", "txt"] },
  { name: "Word document", extensions: ["docx", "docm", "dotx", "dotm", "xml", "doc"] },
  { name: "Rich text or OpenDocument", extensions: ["rtf", "odt"] },
  { name: "Markdown, web page or text", extensions: ["md", "markdown", "html", "htm", "txt"] },
];
export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export function baseName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path;
}

/** Windows' spell checker, in the languages of the user's Windows language list. */
export const windowsSpelling = {
  languages: () => invoke<string[]>("spell_languages"),
  check: (words: string[]) => invoke<boolean[]>("spell_check", { words }),
  suggest: (word: string) => invoke<string[]>("spell_suggest", { word }),
};

/** Items in the Zotero library on this computer matching `query`, as CSL-JSON (Zotero must be running). */
export const searchZotero = (query: string) => invoke<string>("zotero_search", { query });

export async function readPath(path: string): Promise<OpenedFile> {
  const buffer = await invoke<ArrayBuffer>("read_file", { path });
  return { name: baseName(path), path, bytes: new Uint8Array(buffer) };
}

export async function pickDocument(): Promise<OpenedFile | null> {
  if (isDesktop) {
    const path = await open({ multiple: false, directory: false, filters: OPEN_FILTERS });
    return path ? readPath(path) : null;
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = `${DOCX_MIME},.docx,.docm,.dotx,.dotm,.xml,.doc,.rtf,.odt,.md,.markdown,.html,.htm,.txt`;
    input.onchange = async () => {
      const file = input.files?.[0];
      resolve(file ? { name: file.name, path: null, bytes: new Uint8Array(await file.arrayBuffer()) } : null);
    };
    input.oncancel = () => resolve(null);
    input.click();
  });
}

/** Asks for a picture file. */
export async function pickImage(): Promise<{ bytes: Uint8Array; ext: string } | null> {
  const extensions = ["png", "jpg", "jpeg", "gif", "bmp", "webp"];
  if (isDesktop) {
    const path = await open({ multiple: false, directory: false, filters: [{ name: "Pictures", extensions }] });
    if (!path) return null;
    const file = await readPath(path);
    return { bytes: file.bytes, ext: (path.split(".").pop() ?? "png").toLowerCase() };
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.onchange = async () => {
      const file = input.files?.[0];
      resolve(file ? { bytes: new Uint8Array(await file.arrayBuffer()), ext: (file.name.split(".").pop() ?? "png").toLowerCase() } : null);
    };
    input.oncancel = () => resolve(null);
    input.click();
  });
}

/** Asks where to save a PDF; returns null if cancelled. */
export async function pickPdfPath(suggested: string): Promise<string | null> {
  if (!isDesktop) return null;
  const path = await save({ defaultPath: suggested, filters: [{ name: "PDF", extensions: ["pdf"] }] });
  if (!path) return null;
  return /\.pdf$/i.test(path) ? path : `${path}.pdf`;
}

/** Asks where to save a copy in another format; returns null if cancelled. */
export async function pickExportPath(suggested: string, name: string, ext: string): Promise<string | null> {
  if (!isDesktop) return suggested;
  const path = await save({ defaultPath: suggested, filters: [{ name, extensions: [ext] }] });
  if (!path) return null;
  return path.toLowerCase().endsWith(`.${ext}`) ? path : `${path}.${ext}`;
}

/** Asks where to save; returns null if cancelled. In the browser there is nothing to ask. */
export async function pickSavePath(suggested: string): Promise<string | null> {
  if (!isDesktop) return suggested;
  // The document's own kind first, so saving keeps it unless another is chosen.
  const ext = /\.(docx|docm|dotx|dotm|xml)$/i.exec(suggested)?.[1]?.toLowerCase() ?? "docx";
  const filters = [...SAVE_FILTERS].sort((a, b) => Number(b.extensions[0] === ext) - Number(a.extensions[0] === ext));
  const path = await save({ defaultPath: suggested, filters });
  if (!path) return null;
  return /\.(docx|docm|dotx|dotm|xml)$/i.test(path) ? path : `${path}.${ext}`;
}

export async function writePath(path: string, bytes: Uint8Array): Promise<void> {
  if (isDesktop) {
    await invoke("write_file", bytes, { headers: { "x-path": encodeURIComponent(path) } });
    return;
  }
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type: DOCX_MIME }));
  const link = document.createElement("a");
  link.href = url;
  link.download = baseName(path);
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** The documents DuckWriter was started with. */
/** Opens a document (or a new one) in a window of its own. */
export async function openWindow(path: string | null): Promise<void> {
  if (isDesktop) await invoke("open_window", { path });
  else window.open(location.href, "_blank");
}

/** What this window was opened for: a document's path, "" for a new one; undefined for the first window. */
export function windowOpening(): string | undefined {
  return (window as { __DUCKWRITER_OPEN?: string }).__DUCKWRITER_OPEN;
}

export async function startupFiles(): Promise<string[]> {
  return isDesktop ? invoke<string[]>("startup_files") : [];
}

/** Full screen for focus mode (in a browser: the page's own full screen). */
export async function setFullscreen(on: boolean) {
  // A test window stays off-screen.
  if ((window as { __DUCKWRITER_TEST?: boolean }).__DUCKWRITER_TEST) return;
  try {
    if (isDesktop) await getCurrentWindow().setFullscreen(on);
    else if (on) await document.documentElement.requestFullscreen();
    else if (document.fullscreenElement) await document.exitFullscreen();
  } catch {
    // Without full screen, focus mode still hides the rest.
  }
}

export function setTitle(title: string) {
  document.title = title;
  if (isDesktop) void getCurrentWindow().setTitle(title);
}

/** Calls `handler` with the paths of files dropped onto the window. */
export function onPathsDropped(handler: (paths: string[]) => void) {
  if (!isDesktop) return;
  void getCurrentWebview().onDragDropEvent((event) => {
    if (event.payload.type === "drop") handler(event.payload.paths);
  });
}

/**
 * Lets `confirm` veto closing the window (e.g. to offer saving first).
 * It resolves to true when the window may close.
 */
export function onCloseRequested(confirm: () => Promise<boolean>) {
  if (!isDesktop) return;
  const appWindow = getCurrentWindow();
  void appWindow.onCloseRequested(async (event) => {
    event.preventDefault();
    if (await confirm()) await appWindow.destroy();
  });
}
