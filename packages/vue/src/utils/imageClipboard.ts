/**
 * Image clipboard + replace helpers. Pure DOM/PM operations — every
 * function takes the view as a parameter so they're testable and don't
 * close over Vue refs.
 *
 * `copyImageToClipboard` emits both `text/html` (so a subsequent paste
 * re-creates a PM image node via `pasteFromClipboard`) and a `text/plain`
 * fallback. `pasteFromClipboard` walks the clipboard items looking for
 * an image blob first, then an HTML payload carrying our custom data
 * attributes, then plain text.
 *
 * Vue port of the upstream docx-editor util. `makeRevisionInfo` resolves
 * to our fork's suggestion-mode plugin module (the plugins barrel does not
 * re-export it).
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { makeRevisionInfo } from "@stll/folio-core/prosemirror/plugins/suggestionMode";

/**
 * Apply the `insertion` mark to the just-replaced image when suggesting
 * mode is active, so clipboard-pasted images round-trip as tracked
 * additions. Run after replaceSelectionWith on the SAME tr.
 */
function tagPastedImageAsInsertion(view: EditorView, tr: Transaction, imageNode: PMNode): void {
  const info = makeRevisionInfo(view.state);
  const insertionType = view.state.schema.marks["insertion"];
  if (!info || !insertionType) return;
  // After replaceSelectionWith, the image lives at the prior selection's
  // start position (mapped). The cursor sits just after it.
  const to = tr.selection.from;
  const from = to - imageNode.nodeSize;
  if (from < 0) return;
  tr.addMark(
    from,
    to,
    insertionType.create({
      revisionId: info.id,
      author: info.author,
      date: info.date,
    }),
  );
}

export function copyImageToClipboard(view: EditorView, pmPos: number): void {
  const node = view.state.doc.nodeAt(pmPos);
  if (!node || node.type.name !== "image") return;

  const src = node.attrs["src"];
  if (typeof src !== "string") return;

  const width = String(node.attrs["width"] ?? "");
  const height = String(node.attrs["height"] ?? "");
  const wrapType = String(node.attrs["wrapType"] ?? "");
  const displayMode = String(node.attrs["displayMode"] ?? "");
  const rId = String(node.attrs["rId"] ?? "");
  const imgHtml = `<img src="${src}" data-pm-image="true" data-width="${width}" data-height="${height}" data-wrap-type="${wrapType}" data-display-mode="${displayMode}" data-rid="${rId}" />`;

  const clipboardItem = new ClipboardItem({
    "text/html": new Blob([imgHtml], { type: "text/html" }),
    "text/plain": new Blob(["[image]"], { type: "text/plain" }),
  });
  navigator.clipboard.write([clipboardItem]).catch(() => {
    // Fallback: at least copy as HTML
    const ta = document.createElement("textarea");
    ta.value = imgHtml;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  });
}

export function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      if (typeof reader.result === "string") resolve(reader.result);
      else reject(new Error("Failed to read blob as data URL"));
    };
    reader.onerror = () => reject(reader.error ?? new Error("Failed to read blob"));
    reader.readAsDataURL(blob);
  });
}

export function loadImageDimensions(src: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth, height: img.naturalHeight });
    img.onerror = () => resolve({ width: 200, height: 200 });
    img.src = src;
  });
}

function insertPastedImage(view: EditorView, src: string, width: number, height: number): void {
  const imageType = view.state.schema.nodes["image"];
  if (!imageType) return;
  const imageNode = imageType.create({
    src,
    width,
    height,
    rId: `rId_img_${Date.now()}`,
    wrapType: "inline",
    displayMode: "inline",
  });
  const tr = view.state.tr.replaceSelectionWith(imageNode);
  tagPastedImageAsInsertion(view, tr, imageNode);
  view.dispatch(tr);
}

export async function pasteFromClipboard(view: EditorView): Promise<void> {
  try {
    const items = await navigator.clipboard.read();
    for (const item of items) {
      const imageType = item.types.find((t) => t.startsWith("image/"));
      if (imageType) {
        const blob = await item.getType(imageType);
        const dataUrl = await blobToDataUrl(blob);
        const dims = await loadImageDimensions(dataUrl);
        const maxW = 612;
        let w = dims.width;
        let h = dims.height;
        if (w > maxW) {
          h = Math.round(h * (maxW / w));
          w = maxW;
        }
        insertPastedImage(view, dataUrl, w, h);
        return;
      }

      if (item.types.includes("text/html")) {
        const htmlBlob = await item.getType("text/html");
        const html = await htmlBlob.text();
        const match = /<img[^>]+src="(?<src>[^"]+)"[^>]*>/iu.exec(html);
        const src = match?.groups?.["src"];
        if (src) {
          const widthMatch = /data-width="(?<w>\d+)"/u.exec(html)?.groups?.["w"];
          const heightMatch = /data-height="(?<h>\d+)"/u.exec(html)?.groups?.["h"];
          const w = widthMatch ? Number(widthMatch) : 200;
          const h = heightMatch ? Number(heightMatch) : 200;
          insertPastedImage(view, src, w || 200, h || 200);
          return;
        }
      }

      if (item.types.includes("text/plain")) {
        const textBlob = await item.getType("text/plain");
        const text = await textBlob.text();
        if (text && text !== "[image]") {
          const { from } = view.state.selection;
          view.dispatch(view.state.tr.insertText(text, from));
        }
        return;
      }
    }
  } catch {
    // Fallback for browsers without the async clipboard read API.
    const text = await navigator.clipboard?.readText();
    if (text) {
      const { from } = view.state.selection;
      view.dispatch(view.state.tr.insertText(text, from));
    }
  }
}

export function triggerReplaceImage(view: EditorView, pmPos: number): void {
  const node = view.state.doc.nodeAt(pmPos);
  if (!node || node.type.name !== "image") return;

  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/*";
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;
    const dataUrl = await blobToDataUrl(file);
    const dims = await loadImageDimensions(dataUrl);

    // Keep existing dimensions unless the aspect ratio is wildly different;
    // scale the new image to fit within the old bounding box.
    const oldWRaw = node.attrs["width"];
    const oldHRaw = node.attrs["height"];
    const oldW = typeof oldWRaw === "number" ? oldWRaw : dims.width;
    const oldH = typeof oldHRaw === "number" ? oldHRaw : dims.height;
    const scale = Math.min(oldW / dims.width, oldH / dims.height);
    const newW = Math.round(dims.width * scale);
    const newH = Math.round(dims.height * scale);

    try {
      const tr = view.state.tr.setNodeMarkup(pmPos, undefined, {
        ...node.attrs,
        src: dataUrl,
        width: newW,
        height: newH,
        rId: `rId_img_${Date.now()}`,
      });
      view.dispatch(tr);
    } catch {
      // Position may have changed between picker open and file selection.
    }
  };
  input.click();
}
