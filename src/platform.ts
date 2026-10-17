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

const DOCX_FILTER = [{ name: "Word document", extensions: ["docx"] }];
export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

export function baseName(path: string): string {
  return path.split(/[\\/]/).pop() ?? path;
}

export async function readPath(path: string): Promise<OpenedFile> {
  const buffer = await invoke<ArrayBuffer>("read_file", { path });
  return { name: baseName(path), path, bytes: new Uint8Array(buffer) };
}

export async function pickDocument(): Promise<OpenedFile | null> {
  if (isDesktop) {
    const path = await open({ multiple: false, directory: false, filters: DOCX_FILTER });
    return path ? readPath(path) : null;
  }
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = `${DOCX_MIME},.docx`;
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

/** Asks where to save; returns null if cancelled. In the browser there is nothing to ask. */
export async function pickSavePath(suggested: string): Promise<string | null> {
  if (!isDesktop) return suggested;
  const path = await save({ defaultPath: suggested, filters: DOCX_FILTER });
  if (!path) return null;
  return /\.docx$/i.test(path) ? path : `${path}.docx`;
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
export async function startupFiles(): Promise<string[]> {
  return isDesktop ? invoke<string[]>("startup_files") : [];
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
