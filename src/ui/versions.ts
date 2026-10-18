// Earlier versions of the open document, kept on this computer each time it
// is saved over (Alt+Shift+V): look at the list, and bring one back.

import { invoke } from "@tauri-apps/api/core";
import { h } from "../dom";

export interface Version {
  path: string;
  time: number;
  size: number;
}

export interface VersionsHost {
  readonly path: string | null;
  readonly name: string;
  restoreVersion(version: Version): void;
  compareVersion(version: Version): void;
  toast(text: string, kind?: "info" | "error"): void;
}

const when = (ms: number) => new Date(ms).toLocaleString(undefined, { weekday: "short", day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" });
const size = (bytes: number) => (bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1024 / 1024).toFixed(1)} MB`);

export async function showVersions(host: VersionsHost) {
  if (!host.path) return host.toast("Earlier versions are kept once the document has been saved.");
  let versions: Version[];
  try {
    versions = await invoke<Version[]>("versions_list", { path: host.path });
  } catch (error) {
    return host.toast(String(error), "error");
  }
  const list = h(
    "div",
    { class: "source-list" },
    ...(versions.length
      ? versions.map((v) => {
          const restore = h("button", { type: "button", class: "btn small" }, "Restore");
          restore.addEventListener("click", () => {
            dialog.close();
            host.restoreVersion(v);
          });
          const compare = h("button", { type: "button", class: "btn small", title: "Show what changed since then, as tracked changes" }, "Compare");
          compare.addEventListener("click", () => {
            dialog.close();
            host.compareVersion(v);
          });
          return h("div", { class: "source-item" }, h("span", { class: "source-label" }, when(v.time)), h("span", { class: "source-tag" }, size(v.size)), compare, restore);
        })
      : [h("div", { class: "nav-empty" }, "None yet: a version is kept each time you save over the document.")]),
  );
  const dialog = h(
    "dialog",
    { class: "dialog wide" },
    h("h2", {}, `Earlier versions of ${host.name}`),
    h("p", { class: "dialog-hint" }, "Each time the document is saved, the file as it was before is kept here, on this computer (the last 20). Restoring one shows it as the document; save to keep it, or close without saving to go back."),
    list,
    h("div", { class: "dialog-footer" }, h("button", { type: "button", class: "btn primary", onclick: () => dialog.close() }, "Close")),
  ) as HTMLDialogElement;
  dialog.addEventListener("close", () => dialog.remove());
  document.body.append(dialog);
  dialog.showModal();
  (dialog.querySelector(".btn") as HTMLElement).focus();
}
