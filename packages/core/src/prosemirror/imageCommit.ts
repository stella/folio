/**
 * Image resize / drag-move PM commits shared by the React and Vue adapters.
 *
 * The float-vs-inline fork is identical across adapters; only the DOM lookups
 * differ (which `.layout-page-content` the drop landed in; how the inline drop
 * position is hit-tested). Those stay in each adapter, which passes the
 * resolved EMU offsets (float) or drop position (inline) into the pure commits
 * here. Each commit dispatches and returns the PM position the caller should
 * re-select as a NodeSelection (or null on no-op / failure).
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";

/**
 * An image resize handle: the 4 corners ('nw'/'ne'/'se'/'sw') resize both axes
 * and keep the aspect ratio; the 4 edge midpoints ('n'/'s'/'e'/'w') resize a
 * single axis so the user can deliberately stretch the image (break aspect).
 */
export type ImageResizeHandle = "nw" | "ne" | "se" | "sw" | "n" | "s" | "e" | "w";

const MIN_IMAGE_PX = 20;
const MAX_IMAGE_PX = 2000;

/**
 * New image dimensions for a resize drag, shared by the React and Vue overlays
 * (issue #266). Corner handles drive both axes (aspect-locked unless
 * `lockAspect` is false, e.g. Shift held); edge handles drive one axis and
 * never lock. The non-driven axis is returned unchanged. Driven axes are
 * clamped to a sane pixel range.
 */
export function calculateResizedImageDimensions(
  handle: ImageResizeHandle,
  deltaX: number,
  deltaY: number,
  startWidth: number,
  startHeight: number,
  lockAspect: boolean,
): { width: number; height: number } {
  const drivesWidth = handle.includes("w") || handle.includes("e");
  const drivesHeight = handle.includes("n") || handle.includes("s");
  const isCorner = drivesWidth && drivesHeight;

  const signX = handle.includes("w") ? -1 : 1;
  const signY = handle.includes("n") ? -1 : 1;

  let newWidth = drivesWidth ? startWidth + deltaX * signX : startWidth;
  let newHeight = drivesHeight ? startHeight + deltaY * signY : startHeight;

  if (isCorner && lockAspect) {
    const scale = Math.max(newWidth / startWidth, newHeight / startHeight);
    newWidth = startWidth * scale;
    newHeight = startHeight * scale;
  }

  const clamp = (n: number) => Math.max(MIN_IMAGE_PX, Math.min(MAX_IMAGE_PX, n));
  return {
    width: drivesWidth ? clamp(newWidth) : startWidth,
    height: drivesHeight ? clamp(newHeight) : startHeight,
  };
}

/** True when the image is floating (anchored) rather than inline. */
export function isFloatingImage(node: PMNode): boolean {
  const wrapType = node.attrs["wrapType"] as string | undefined;
  return (
    node.attrs["displayMode"] === "float" ||
    (wrapType ? ["square", "tight", "through"].includes(wrapType) : false)
  );
}

/** Resolve the image node at `pmPos`, or null if it isn't an image. */
function imageNodeAt(view: EditorView, pmPos: number): PMNode | null {
  const node = view.state.doc.nodeAt(pmPos);
  if (!node || node.type.name !== "image") return null;
  return node;
}

/**
 * Resize commit: set the image node's `width`/`height`. Returns `pmPos` to
 * re-select, or null if the position no longer holds an image.
 */
export function commitImageResize(
  view: EditorView,
  pmPos: number,
  newWidth: number,
  newHeight: number,
): number | null {
  try {
    const node = imageNodeAt(view, pmPos);
    if (!node) return null;
    view.dispatch(
      view.state.tr.setNodeMarkup(pmPos, undefined, {
        ...node.attrs,
        width: newWidth,
        height: newHeight,
      }),
    );
    return pmPos;
  } catch {
    // Position may have shifted during resize.
    return null;
  }
}

/**
 * Floating drag commit: rewrite the anchor's margin-relative `position`
 * offsets (in EMU) so the image lands at the drop point while staying
 * floating. Returns `pmPos` to re-select, or null on no-op / failure.
 */
export function commitImageFloatMove(
  view: EditorView,
  pmPos: number,
  hOffsetEmu: number,
  vOffsetEmu: number,
): number | null {
  try {
    const node = imageNodeAt(view, pmPos);
    if (!node) return null;
    const newPosition = {
      horizontal: { posOffset: hOffsetEmu, relativeTo: "margin" },
      vertical: { posOffset: vOffsetEmu, relativeTo: "margin" },
    };
    view.dispatch(
      view.state.tr.setNodeMarkup(pmPos, undefined, { ...node.attrs, position: newPosition }),
    );
    return pmPos;
  } catch {
    return null;
  }
}

/**
 * Inline drag commit: move the image node to `dropPos` via a delete + insert
 * pair. Returns the PM position to re-select, or null when the drop is a
 * no-op (same slot) or fails.
 */
export function commitImageInlineMove(
  view: EditorView,
  pmPos: number,
  dropPos: number,
): number | null {
  try {
    const node = imageNodeAt(view, pmPos);
    if (!node) return null;
    if (dropPos === pmPos || dropPos === pmPos + 1) return null;

    let tr = view.state.tr;
    if (dropPos <= pmPos) {
      tr = tr.delete(pmPos, pmPos + node.nodeSize);
      tr = tr.insert(dropPos, node);
      view.dispatch(tr);
      return dropPos;
    }
    tr = tr.delete(pmPos, pmPos + node.nodeSize);
    const adjusted = Math.min(dropPos - node.nodeSize, tr.doc.content.size);
    tr = tr.insert(adjusted, node);
    view.dispatch(tr);
    return Math.min(adjusted, view.state.doc.content.size - 1);
  } catch {
    // Position may have shifted between the drag's frames.
    return null;
  }
}
