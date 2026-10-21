// Copying pictures so they paste anywhere. A picture's address only works in
// the window it is in, so copying puts the picture itself on the clipboard:
// one picture on its own as a PNG (which every program takes), pictures with
// text as HTML with the pictures written into it. Pasting that HTML back gives
// the pictures their own copies (see pastedPictures).

import { Fragment, type Node } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";
import type { DocContext } from "../docx/model";
import { schema } from "./schema";

function dataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

async function png(blob: Blob): Promise<Blob> {
  if (blob.type === "image/png") return blob;
  const bitmap = await createImageBitmap(blob);
  const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
  canvas.getContext("2d")!.drawImage(bitmap, 0, 0);
  return canvas.convertToBlob({ type: "image/png" });
}

/** Called on copy and cut: when pictures are copied, puts them on the clipboard properly. */
export function copyPictures(view: EditorView) {
  const slice = view.state.selection.content();
  const srcs = new Set<string>();
  slice.content.descendants((node) => {
    if (node.type === schema.nodes.image && node.attrs.src) srcs.add(node.attrs.src);
  });
  if (!srcs.size || !navigator.clipboard?.write) return;
  // Taken now: a cut removes the pictures straight after.
  const { dom, text } = view.serializeForClipboard(slice);
  void (async () => {
    const blobs = new Map(await Promise.all([...srcs].map(async (src) => [src, await (await fetch(src)).blob()] as const)));
    let items: Record<string, Blob>;
    if (srcs.size === 1 && !text.trim()) items = { "image/png": await png([...blobs.values()][0]) };
    else {
      for (const img of dom.querySelectorAll("img")) {
        const blob = blobs.get(img.getAttribute("src") ?? "");
        if (blob) img.setAttribute("src", await dataUrl(blob));
      }
      items = { "text/html": new Blob([dom.innerHTML], { type: "text/html" }), "text/plain": new Blob([text], { type: "text/plain" }) };
    }
    await navigator.clipboard.write([new ClipboardItem(items)]);
  })().catch((error) => console.warn("The pictures could not be put on the clipboard", error));
}

/** Pictures in pasted HTML get their own copy in this document; ones that can't be read are left out. */
export function pastedPictures(fragment: Fragment, ctx: DocContext): Fragment {
  const out: Node[] = [];
  fragment.forEach((node) => {
    if (node.type !== schema.nodes.image) {
      out.push(node.isLeaf ? node : node.copy(pastedPictures(node.content, ctx)));
      return;
    }
    const m = /^data:image\/(png|jpeg|gif|bmp|webp);base64,(.+)$/.exec(node.attrs.src);
    if (!m) return;
    const bytes = Uint8Array.from(atob(m[2]), (ch) => ch.charCodeAt(0));
    const ext = m[1] === "jpeg" ? "jpg" : m[1];
    const key = `m${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
    ctx.media.set(key, { bytes, ext, rid: null });
    out.push(schema.nodes.image.create({ ...node.attrs, media: key, rid: null, xml: null, src: URL.createObjectURL(new Blob([bytes as BlobPart], { type: `image/${m[1]}` })) }));
  });
  return Fragment.from(out);
}
