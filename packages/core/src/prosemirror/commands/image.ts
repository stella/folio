/**
 * Image commands
 *
 * Framework-agnostic ProseMirror operations for editing and inserting images,
 * extracted from the React `useImageHandlers` hook. The pure helpers (wrap-mode
 * resolution, transform math, position normalization, size constraint) carry no
 * editor state; the `apply*` operations mutate the node at a known position; and
 * `insertImageFromFile` reads a file and inserts an image node.
 *
 * The adapter keeps only React glue: the dialog open/close state, the file
 * input, and refocusing the editor after each operation.
 */

import type { EditorView } from "prosemirror-view";

import {
  IMAGE_HORIZONTAL_ALIGNMENT_VALUES,
  IMAGE_HORIZONTAL_RELATIVE_TO_VALUES,
  IMAGE_VERTICAL_ALIGNMENT_VALUES,
  IMAGE_VERTICAL_RELATIVE_TO_VALUES,
} from "../../types/documentEnumValues";
import { isSafeImageFile } from "../../utils/imageValidation";
import { sanitizeImageSrc } from "../../utils/sanitizeImageSrc";
import { expectImageAttrs, mergeImageAttrs } from "../attrs";
import type { ImageAttrs, ImagePositionAttrs } from "../schema/nodes";

// ============================================================================
// INPUT TYPES (structural mirrors of the adapter's dialog payloads)
// ============================================================================

export type ImagePositionInput = {
  horizontal?: {
    relativeTo?: string;
    posOffset?: number;
    align?: string;
  };
  vertical?: {
    relativeTo?: string;
    posOffset?: number;
    align?: string;
  };
  distTop?: number;
  distBottom?: number;
  distLeft?: number;
  distRight?: number;
};

export type ImagePropertiesInput = {
  alt?: string;
  borderWidth?: number;
  borderColor?: string;
  borderStyle?: string;
};

export type ImageTransformAction = "rotateCW" | "rotateCCW" | "flipH" | "flipV";

// ============================================================================
// PURE HELPERS
// ============================================================================

type ResolvedImageWrap = {
  wrapType: ImageAttrs["wrapType"];
  displayMode: ImageAttrs["displayMode"];
  cssFloat: ImageAttrs["cssFloat"];
};

/** Map a toolbar wrap selection to the image's wrapType/displayMode/cssFloat, or null when unknown. */
export const resolveImageWrap = (wrapType: string): ResolvedImageWrap | null => {
  switch (wrapType) {
    case "inline":
      return { wrapType: "inline", displayMode: "inline", cssFloat: undefined };
    case "square":
    case "tight":
    case "through":
      return { wrapType, displayMode: "float", cssFloat: "left" };
    case "topAndBottom":
      return { wrapType: "topAndBottom", displayMode: "block", cssFloat: undefined };
    case "behind":
    case "inFront":
      return { wrapType, displayMode: "float", cssFloat: "none" };
    case "wrapLeft":
      return { wrapType: "square", displayMode: "float", cssFloat: "right" };
    case "wrapRight":
      return { wrapType: "square", displayMode: "float", cssFloat: "left" };
    default:
      return null;
  }
};

/** The CSS transform after applying a rotate/flip action to the current transform string. */
export const computeImageTransform = (
  currentTransform: string,
  action: ImageTransformAction,
): string | undefined => {
  const rotateMatch = /rotate\((?<degrees>-?\d+(?:\.\d+)?)deg\)/u.exec(currentTransform);
  // SAFETY: `degrees` group always present when regex matches
  let rotation = rotateMatch ? Number.parseFloat(rotateMatch.groups!["degrees"]!) : 0;
  let hasFlipH = currentTransform.includes("scaleX(-1)");
  let hasFlipV = currentTransform.includes("scaleY(-1)");

  switch (action) {
    case "rotateCW":
      rotation = (rotation + 90) % 360;
      break;
    case "rotateCCW":
      rotation = (rotation - 90 + 360) % 360;
      break;
    case "flipH":
      hasFlipH = !hasFlipH;
      break;
    case "flipV":
      hasFlipV = !hasFlipV;
      break;
    default:
      break;
  }

  const parts: string[] = [];
  if (rotation !== 0) {
    parts.push(`rotate(${rotation}deg)`);
  }
  if (hasFlipH) {
    parts.push("scaleX(-1)");
  }
  if (hasFlipV) {
    parts.push("scaleY(-1)");
  }
  return parts.length > 0 ? parts.join(" ") : undefined;
};

const isOneOf = <T extends string>(value: string | undefined, values: readonly T[]): value is T =>
  value !== undefined && values.some((allowed) => allowed === value);

export const normalizeImageHorizontalPosition = (
  data: ImagePositionInput["horizontal"],
): ImagePositionAttrs["horizontal"] | undefined => {
  if (!data) {
    return undefined;
  }
  if (!isOneOf(data.relativeTo, IMAGE_HORIZONTAL_RELATIVE_TO_VALUES)) {
    return undefined;
  }

  const align = isOneOf(data.align, IMAGE_HORIZONTAL_ALIGNMENT_VALUES) ? data.align : undefined;
  return {
    relativeTo: data.relativeTo,
    ...(typeof data.posOffset === "number" ? { posOffset: data.posOffset } : {}),
    ...(align ? { align } : {}),
  };
};

export const normalizeImageVerticalPosition = (
  data: ImagePositionInput["vertical"],
): ImagePositionAttrs["vertical"] | undefined => {
  if (!data) {
    return undefined;
  }
  if (!isOneOf(data.relativeTo, IMAGE_VERTICAL_RELATIVE_TO_VALUES)) {
    return undefined;
  }

  const align = isOneOf(data.align, IMAGE_VERTICAL_ALIGNMENT_VALUES) ? data.align : undefined;
  return {
    relativeTo: data.relativeTo,
    ...(typeof data.posOffset === "number" ? { posOffset: data.posOffset } : {}),
    ...(align ? { align } : {}),
  };
};

/** Clamp an image's intrinsic size to the page text width, preserving aspect ratio. */
export const constrainImageSize = (
  width: number,
  height: number,
): { width: number; height: number } => {
  const maxWidth = 612; // ~6.375 inches
  if (width <= maxWidth) {
    return { width, height };
  }

  const scale = maxWidth / width;
  return {
    width: maxWidth,
    height: Math.max(1, Math.round(height * scale)),
  };
};

// ============================================================================
// VIEW OPERATIONS
// ============================================================================

const imageNodeAt = (view: EditorView, pos: number) => {
  const node = view.state.doc.nodeAt(pos);
  return node && node.type.name === "image" ? node : null;
};

/** Change the wrap mode of the image at `pos`. Returns whether it applied. */
export const applyImageWrapType = (view: EditorView, pos: number, wrapType: string): boolean => {
  const node = imageNodeAt(view, pos);
  if (!node) {
    return false;
  }

  const resolved = resolveImageWrap(wrapType);
  if (!resolved) {
    return false;
  }

  const tr = view.state.tr.setNodeMarkup(
    pos,
    undefined,
    mergeImageAttrs(node, {
      wrapType: resolved.wrapType,
      displayMode: resolved.displayMode,
      cssFloat: resolved.cssFloat,
    }),
  );
  view.dispatch(tr.scrollIntoView());
  return true;
};

/** Rotate or flip the image at `pos`. Returns whether it applied. */
export const applyImageTransform = (
  view: EditorView,
  pos: number,
  action: ImageTransformAction,
): boolean => {
  const node = imageNodeAt(view, pos);
  if (!node) {
    return false;
  }

  const currentTransform = expectImageAttrs(node).transform ?? "";
  const transform = computeImageTransform(currentTransform, action);

  const tr = view.state.tr.setNodeMarkup(pos, undefined, mergeImageAttrs(node, { transform }));
  view.dispatch(tr.scrollIntoView());
  return true;
};

/** Apply position + text-wrap distances to the image at `pos`. Returns whether it applied. */
export const applyImagePosition = (
  view: EditorView,
  pos: number,
  data: ImagePositionInput,
): boolean => {
  const node = imageNodeAt(view, pos);
  if (!node) {
    return false;
  }

  const attrs = expectImageAttrs(node);
  const horizontal = normalizeImageHorizontalPosition(data.horizontal);
  const vertical = normalizeImageVerticalPosition(data.vertical);
  const tr = view.state.tr.setNodeMarkup(
    pos,
    undefined,
    mergeImageAttrs(node, {
      position:
        horizontal || vertical
          ? {
              ...(horizontal ? { horizontal } : {}),
              ...(vertical ? { vertical } : {}),
            }
          : undefined,
      distTop: data.distTop ?? attrs.distTop,
      distBottom: data.distBottom ?? attrs.distBottom,
      distLeft: data.distLeft ?? attrs.distLeft,
      distRight: data.distRight ?? attrs.distRight,
    }),
  );
  view.dispatch(tr.scrollIntoView());
  return true;
};

/** Apply alt text + border properties to the image at `pos`. Returns whether it applied. */
export const applyImageProperties = (
  view: EditorView,
  pos: number,
  data: ImagePropertiesInput,
): boolean => {
  const node = imageNodeAt(view, pos);
  if (!node) {
    return false;
  }

  const tr = view.state.tr.setNodeMarkup(
    pos,
    undefined,
    mergeImageAttrs(node, {
      alt: data.alt,
      borderWidth: data.borderWidth,
      borderColor: data.borderColor,
      borderStyle: data.borderStyle,
    }),
  );
  view.dispatch(tr.scrollIntoView());
  return true;
};

/**
 * Read an image file, then insert a constrained inline image node at the
 * current selection. `onInserted` runs after the image is dispatched. No-ops
 * when the file fails validation or the schema has no image node.
 */
export const insertImageFromFile = async (
  view: EditorView,
  file: File,
  onInserted: () => void,
): Promise<void> => {
  if (!(await isSafeImageFile(file))) {
    return;
  }

  const dataUrl = await readFileAsDataUrl(file);
  const { width, height } = await loadImageDimensions(dataUrl);
  const constrained = constrainImageSize(width, height);
  const imageType = view.state.schema.nodes["image"];
  if (!imageType) {
    return;
  }

  const imageNode = imageType.create({
    src: dataUrl,
    alt: file.name,
    width: constrained.width,
    height: constrained.height,
    rId: `rId_img_${Date.now()}`,
    wrapType: "inline",
    displayMode: "inline",
  });

  const { from } = view.state.selection;
  const tr = view.state.tr.insert(from, imageNode);
  view.dispatch(tr.scrollIntoView());
  onInserted();
};

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => {
      resolve(typeof reader.result === "string" ? reader.result : "");
    });
    reader.addEventListener("error", () => {
      reject(reader.error ?? new Error("Failed to read image file"));
    });
    reader.readAsDataURL(file);
  });
}

function loadImageDimensions(src: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const safeSrc = sanitizeImageSrc(src);
    if (!safeSrc) {
      reject(new Error("Rejected unsafe image source"));
      return;
    }
    const img = new Image();
    img.addEventListener("load", () => {
      resolve({
        width: img.naturalWidth || 1,
        height: img.naturalHeight || 1,
      });
    });
    img.addEventListener("error", () => {
      reject(new Error("Failed to load image"));
    });
    img.src = safeSrc;
  });
}
