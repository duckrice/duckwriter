// Auto-save and recovery, all on this computer.
//
// - Auto-save: a few seconds after you stop typing, a document that has a
//   file is saved to that file. It can be switched off.
// - Recovery: while there are unsaved changes, a copy is kept in a private
//   folder (%LOCALAPPDATA%). It is deleted when the work is saved or the
//   window is closed normally, so a copy that is still there when DuckWriter
//   starts means it closed unexpectedly, and it is offered back.

import { invoke } from "@tauri-apps/api/core";
import { isDesktop, readPath, writePath } from "./platform";

/** How long after the last change auto-save waits. */
const QUIET = 4000;
/** How often a recovery copy is refreshed while there are unsaved changes. */
const RECOVERY_EVERY = 60_000;
/** A recovery copy not refreshed for this long belongs to a DuckWriter that is no longer running. */
const ABANDONED_AFTER = 3 * 60_000;

export interface AutoSaveHost {
  readonly dirty: boolean;
  /** The file the document is saved in, or null for a new one. */
  readonly path: string | null;
  readonly name: string;
  /** The document as .docx bytes. */
  build(): Promise<Uint8Array>;
  /** Saves to the document's file without asking anything. Resolves to false if it failed. */
  saveQuietly(): Promise<boolean>;
  /** Shows how saving is going: "Saved 14:02", "Auto-save off"… */
  status(text: string, title: string): void;
}

export interface Abandoned {
  name: string;
  path: string | null;
  time: number;
  docx: string;
  info: string;
}

export function autoSaveOn(): boolean {
  try {
    return localStorage.getItem("duckwriter.autosave") !== "off";
  } catch {
    return true;
  }
}

function setAutoSaveOn(on: boolean) {
  try {
    localStorage.setItem("duckwriter.autosave", on ? "on" : "off");
  } catch {
    // Not kept, but still applies now.
  }
}

const time = (date = new Date()) => date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });

export class AutoSave {
  /** This window's own recovery files are named after it. */
  private session = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  private quietTimer = 0;
  private recoveryTimer = 0;
  private busy = false;
  private hasRecovery = false;
  private lastSaved: Date | null = null;
  private failed = false;

  constructor(private host: AutoSaveHost) {
    this.recoveryTimer = window.setInterval(() => void this.keepRecovery(), RECOVERY_EVERY);
    this.show();
  }

  get enabled(): boolean {
    return autoSaveOn();
  }

  toggle() {
    setAutoSaveOn(!this.enabled);
    this.show();
    if (this.enabled) this.changed();
  }

  /** Called after every change: auto-save once typing pauses. */
  changed() {
    clearTimeout(this.quietTimer);
    if (!this.host.dirty) {
      this.show();
      return;
    }
    this.quietTimer = window.setTimeout(() => void this.run(), QUIET);
    this.show();
  }

  /** The document was saved (by hand or automatically). */
  saved() {
    this.lastSaved = new Date();
    this.failed = false;
    void this.clearRecovery();
    this.show();
  }

  private async run() {
    if (this.busy || !this.host.dirty) return;
    if (!this.enabled || !this.host.path) {
      await this.keepRecovery();
      return;
    }
    this.busy = true;
    try {
      this.failed = !(await this.host.saveQuietly());
      // If saving to the file failed (it may be open in Word), the work is still kept safe.
      if (this.failed) await this.keepRecovery();
    } finally {
      this.busy = false;
      this.show();
    }
  }

  private show() {
    if (this.failed) {
      this.host.status("Couldn't auto-save", "The file couldn't be written (is it open in another program?). Your work is kept in a recovery copy. Click to switch auto-save off.");
    } else if (!this.enabled) {
      this.host.status("Auto-save off", "Click to switch auto-save on. A recovery copy of unsaved work is still kept on this computer.");
    } else if (!this.host.path) {
      this.host.status("Auto-save on", "Auto-save starts once the document has been saved once (Ctrl+S). Until then a recovery copy is kept on this computer. Click to switch it off.");
    } else if (this.host.dirty) {
      this.host.status("Auto-save on", "Saves to the file a few seconds after you stop typing. Click to switch it off.");
    } else {
      this.host.status(this.lastSaved ? `Saved ${time(this.lastSaved)}` : "Auto-save on", "Saves to the file a few seconds after you stop typing. Click to switch it off.");
    }
  }

  // ---- recovery copies --------------------------------------------------------------

  private async files(): Promise<{ docx: string; info: string } | null> {
    if (!isDesktop) return null;
    const dir = await invoke<string>("recovery_dir");
    return { docx: `${dir}\\${this.session}.docx`, info: `${dir}\\${this.session}.json` };
  }

  /** Keeps a copy of unsaved work in the recovery folder. */
  async keepRecovery() {
    if (!this.host.dirty) return;
    const files = await this.files();
    if (!files) return;
    try {
      await writePath(files.docx, await this.host.build());
      const info = { name: this.host.name, path: this.host.path, time: Date.now() };
      await writePath(files.info, new TextEncoder().encode(JSON.stringify(info)));
      this.hasRecovery = true;
    } catch (error) {
      console.warn("Could not keep a recovery copy", error);
    }
  }

  /** Removes this window's recovery copy: the work was saved, or deliberately left. */
  async clearRecovery() {
    if (!this.hasRecovery) return;
    const files = await this.files();
    if (!files) return;
    await invoke("delete_recovery", { path: files.docx }).catch(() => {});
    await invoke("delete_recovery", { path: files.info }).catch(() => {});
    this.hasRecovery = false;
  }

  dispose() {
    clearTimeout(this.quietTimer);
    clearInterval(this.recoveryTimer);
  }

  /** Recovery copies left behind by a DuckWriter that closed unexpectedly. */
  static async abandoned(): Promise<Abandoned[]> {
    if (!isDesktop) return [];
    let paths: string[];
    try {
      paths = await invoke<string[]>("recovery_files");
    } catch {
      return [];
    }
    const out: Abandoned[] = [];
    for (const info of paths.filter((p) => p.endsWith(".json"))) {
      const docx = info.replace(/\.json$/, ".docx");
      if (!paths.includes(docx)) continue;
      try {
        const data = JSON.parse(new TextDecoder().decode((await readPath(info)).bytes)) as { name: string; path: string | null; time: number };
        // Still being refreshed: another DuckWriter window is using it.
        if (Date.now() - data.time < ABANDONED_AFTER) continue;
        out.push({ ...data, docx, info });
      } catch {
        // A damaged leftover; nothing to offer.
      }
    }
    return out.sort((a, b) => b.time - a.time);
  }

  static async discard(found: Abandoned) {
    await invoke("delete_recovery", { path: found.docx }).catch(() => {});
    await invoke("delete_recovery", { path: found.info }).catch(() => {});
  }
}
