// Adding pictures: from a file, pasted, or dragged in. A new picture sits in
// line with the text, as wide as it is (but no wider than the page).

import { NodeSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { schema } from "../editor/schema";
import { pickImage } from "../platform";

/** Whether pictures snap to the grid (on unless switched off). */
export function snapOn(): boolean {
  try {
    return localStorage.getItem("duckwriter.snap") !== "0";
  } catch {
    return true;
  }
}

export function setSnap(on: boolean) {
  try {
    localStorage.setItem("duckwriter.snap", on ? "1" : "0");
  } catch {
    // Not kept, but still applies now.
  }
}

/** A picture's size in pixels. */
function naturalSize(url: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => reject(new Error("This picture can't be shown."));
    img.src = url;
  });
}

/** Adds a picture at `pos` (or the cursor). */
export async function insertPicture(view: EditorView, ctx: DocContext, bytes: Uint8Array, ext: string, pos: number | null = null) {
  const type = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : `image/${ext}`;
  const src = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const size = await naturalSize(src);
  const page = ctx.page;
  const maxWidth = page.width - page.marginLeft - page.marginRight;
  // Pixels at 96 per inch are three quarters of a point.
  let width = size.width * 0.75;
  let height = size.height * 0.75;
  if (width > maxWidth) {
    height *= maxWidth / width;
    width = maxWidth;
  }
  const key = `m${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  ctx.media.set(key, { bytes, ext: ext === "jpeg" ? "jpg" : ext, rid: null });
  const node = schema.nodes.image.create({ media: key, src, width: Math.round(width * 100) / 100, height: Math.round(height * 100) / 100 });
  const tr = pos === null ? view.state.tr.replaceSelectionWith(node, false) : view.state.tr.insert(pos, node);
  const at = pos ?? tr.selection.from - 1;
  try {
    tr.setSelection(NodeSelection.create(tr.doc, at));
  } catch {
    // Selecting it is only a convenience.
  }
  view.dispatch(tr.scrollIntoView());
  view.focus();
}

/** Asks for a picture file and adds it. */
export async function choosePicture(view: EditorView, ctx: DocContext) {
  const file = await pickImage();
  if (!file) return;
  await insertPicture(view, ctx, file.bytes, file.ext);
}

/** Pictures pasted or dropped in. */
export async function insertPictureFiles(files: File[], view: EditorView, ctx: DocContext, pos: number | null) {
  for (const file of files) {
    const ext = (file.type.split("/")[1] ?? "png").replace("svg+xml", "svg");
    await insertPicture(view, ctx, new Uint8Array(await file.arrayBuffer()), ext, pos);
  }
}
