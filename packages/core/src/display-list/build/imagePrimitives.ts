/**
 * Images: the source table, and the primitives an image draws.
 *
 * The display list carries raw bytes plus the format they are already in, so a
 * backend never owns a codec. This module therefore decodes only far enough to
 * fill `pixelWidth`/`pixelHeight` and to know PNG from JPEG. It transcodes
 * nothing: a WMF, EMF, TIFF, SVG, GIF or BMP source is reported through
 * `unsupported` and paints nothing, because a silently missing picture is worse
 * than a named one.
 */

import type { ImageBlock, ImageFragment } from "../../layout-engine/types";
import { parseRotationDegrees } from "../../utils/rotationBoundingBox";
import { sanitizeImageSrc } from "../../utils/sanitizeImageSrc";
import type {
  DisplayImagePrimitive,
  DisplayImageRef,
  DisplayImageSource,
  DisplayPrimitive,
  DisplayRect,
} from "../types";
import type { BuildContext } from "./buildContext";
import { resolveBorderStroke } from "./strokes";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

const DATA_URL_PATTERN = /^data:([^;,]*)(;base64)?,(.*)$/su;

type DecodedImage = {
  readonly format: "png" | "jpeg";
  readonly bytes: Uint8Array;
  readonly pixelWidth: number;
  readonly pixelHeight: number;
};

type DecodeFailure = { readonly reason: string };

const decodeBase64 = (encoded: string): Uint8Array | undefined => {
  const binary = atob(encoded);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
};

const readUint32BE = (bytes: Uint8Array, offset: number): number =>
  // SAFETY: every caller checks `bytes.length` covers `offset + 4` first.
  (bytes[offset]! << 24) |
  (bytes[offset + 1]! << 16) |
  (bytes[offset + 2]! << 8) |
  bytes[offset + 3]!;

const readUint16BE = (bytes: Uint8Array, offset: number): number =>
  // SAFETY: every caller checks `bytes.length` covers `offset + 2` first.
  (bytes[offset]! << 8) | bytes[offset + 1]!;

const startsWith = (bytes: Uint8Array, signature: readonly number[]): boolean =>
  bytes.length >= signature.length && signature.every((byte, index) => bytes[index] === byte);

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a] as const;
const JPEG_SIGNATURE = [0xff, 0xd8, 0xff] as const;

/** PNG's IHDR is always the first chunk: 8-byte signature, 8-byte header, then w/h. */
const PNG_WIDTH_OFFSET = 16;

const decodePng = (bytes: Uint8Array): DecodedImage | DecodeFailure => {
  if (bytes.length < PNG_WIDTH_OFFSET + 8) {
    return { reason: "PNG shorter than its IHDR chunk" };
  }
  return {
    format: "png",
    bytes,
    pixelWidth: readUint32BE(bytes, PNG_WIDTH_OFFSET),
    pixelHeight: readUint32BE(bytes, PNG_WIDTH_OFFSET + 4),
  };
};

/**
 * Frame markers that carry the image dimensions. Every other `SOFn` code in the
 * `0xC0..0xCF` block is a table or restart marker, not a frame header.
 */
const JPEG_FRAME_MARKERS = new Set([
  0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf,
]);

const decodeJpeg = (bytes: Uint8Array): DecodedImage | DecodeFailure => {
  let offset = 2;
  while (offset + 9 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    // SAFETY: `offset + 9 < bytes.length` guarantees these reads are in range.
    const marker = bytes[offset + 1]!;
    if (JPEG_FRAME_MARKERS.has(marker)) {
      return {
        format: "jpeg",
        bytes,
        pixelHeight: readUint16BE(bytes, offset + 5),
        pixelWidth: readUint16BE(bytes, offset + 7),
      };
    }
    const segmentLength = readUint16BE(bytes, offset + 2);
    if (segmentLength < 2) {
      return { reason: "JPEG segment length is degenerate" };
    }
    offset += 2 + segmentLength;
  }
  return { reason: "JPEG carries no frame header" };
};

const sniffOtherFormat = (bytes: Uint8Array): string | undefined => {
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) {
    return "GIF";
  }
  if (startsWith(bytes, [0x42, 0x4d])) {
    return "BMP";
  }
  if (startsWith(bytes, [0x49, 0x49, 0x2a, 0x00]) || startsWith(bytes, [0x4d, 0x4d, 0x00, 0x2a])) {
    return "TIFF";
  }
  if (startsWith(bytes, [0xd7, 0xcd, 0xc6, 0x9a]) || startsWith(bytes, [0x01, 0x00, 0x00, 0x00])) {
    return "WMF/EMF";
  }
  if (startsWith(bytes, [0x3c, 0x3f, 0x78, 0x6d]) || startsWith(bytes, [0x3c, 0x73, 0x76, 0x67])) {
    return "SVG";
  }
  return undefined;
};

const decodeImageBytes = (bytes: Uint8Array): DecodedImage | DecodeFailure => {
  if (startsWith(bytes, PNG_SIGNATURE)) {
    return decodePng(bytes);
  }
  if (startsWith(bytes, JPEG_SIGNATURE)) {
    return decodeJpeg(bytes);
  }
  const other = sniffOtherFormat(bytes);
  return {
    reason: other === undefined ? "unrecognised image bytes" : `${other} is not PNG or JPEG`,
  };
};

/**
 * Sources interned in first-use order, so two builds of one layout produce the
 * same `DisplayImageRef` for the same picture.
 */
export class ImageTable {
  private readonly sources: DisplayImageSource[] = [];
  private readonly refBySrc = new Map<string, DisplayImageRef | null>();
  private readonly failureBySrc = new Map<string, string>();

  /** `undefined` means "do not paint"; the reason is available via {@link failureFor}. */
  intern(src: string): DisplayImageRef | undefined {
    const cached = this.refBySrc.get(src);
    if (cached !== undefined) {
      return cached ?? undefined;
    }

    const fail = (reason: string): undefined => {
      this.refBySrc.set(src, null);
      this.failureBySrc.set(src, reason);
      return undefined;
    };

    // Only the schemes the painter itself paints (`data:` / `blob:`). A `blob:`
    // URL has no bytes without a fetch, and the builder does no I/O.
    const safeSrc = sanitizeImageSrc(src);
    if (safeSrc === undefined) {
      return fail("image source is not a data: or blob: URL");
    }
    const match = DATA_URL_PATTERN.exec(safeSrc);
    if (!match) {
      return fail("image bytes are not inline (blob: URLs need a fetch the builder cannot do)");
    }
    // SAFETY: groups 2 and 3 exist whenever the pattern matched.
    if (match[2] === undefined) {
      return fail("only base64 data: URLs carry decodable bytes");
    }
    const bytes = decodeBase64(match[3]!);
    if (bytes === undefined || bytes.length === 0) {
      return fail("image data: URL decoded to no bytes");
    }

    const decoded = decodeImageBytes(bytes);
    if ("reason" in decoded) {
      return fail(decoded.reason);
    }

    const ref = this.sources.length;
    this.sources.push(decoded);
    this.refBySrc.set(src, ref);
    return ref;
  }

  failureFor(src: string): string | undefined {
    return this.failureBySrc.get(src);
  }

  snapshot(): readonly DisplayImageSource[] {
    return this.sources;
  }
}

type ImageVisualSource = Pick<
  ImageBlock,
  | "src"
  | "opacity"
  | "cropTop"
  | "cropRight"
  | "cropBottom"
  | "cropLeft"
  | "borderWidth"
  | "borderColor"
  | "borderStyle"
  | "transform"
>;

const cropOf = (source: ImageVisualSource): DisplayImagePrimitive["crop"] => {
  const l = source.cropLeft ?? 0;
  const t = source.cropTop ?? 0;
  const r = source.cropRight ?? 0;
  const b = source.cropBottom ?? 0;
  if (!(l || t || r || b)) {
    return undefined;
  }
  return { l, t, r, b };
};

export type PaintImageOptions = {
  readonly source: ImageVisualSource;
  readonly rect: DisplayRect;
  readonly context: BuildContext;
  /** Names the construct in an `unsupported` entry, e.g. `"anchored image"`. */
  readonly label: string;
};

/**
 * The primitives one picture draws: the bitmap (rotated and/or washed as the
 * source asks) followed by its border. Empty when the bytes are unusable, with
 * the reason already reported.
 */
export const paintImage = ({
  source,
  rect,
  context,
  label,
}: PaintImageOptions): readonly DisplayPrimitive[] => {
  const primitives: DisplayPrimitive[] = [];
  const ref = context.images.intern(source.src);

  if (ref === undefined) {
    const reason = context.images.failureFor(source.src) ?? "image could not be decoded";
    const construct = reason.includes("not PNG or JPEG")
      ? UNSUPPORTED_CONSTRUCT.imageFormat
      : UNSUPPORTED_CONSTRUCT.imageSource;
    context.unsupported.report(construct, context.pageIndex, `${label}: ${reason}`);
  } else {
    const crop = cropOf(source);
    const image: DisplayImagePrimitive = {
      kind: "image",
      image: ref,
      rect,
      ...(crop === undefined ? {} : { crop }),
      opacity: source.opacity == null ? 1 : Math.min(1, Math.max(0, source.opacity)),
    };

    const degrees = parseRotationDegrees(source.transform);
    if (degrees === 0) {
      primitives.push(image);
    } else {
      // Word rotates a picture about its geometric centre, and so does the CSS
      // the painter emits (`transform-origin: center center`).
      primitives.push({
        kind: "rotateGroup",
        degrees,
        originXPx: rect.xPx + rect.widthPx / 2,
        originYPx: rect.yPx + rect.heightPx / 2,
        children: [image],
      });
    }
  }

  if (source.borderWidth != null && source.borderWidth > 0) {
    const { stroke, unresolvedColor } = resolveBorderStroke({
      width: source.borderWidth,
      style: source.borderStyle ?? "solid",
      color: source.borderColor ?? "#000000",
    });
    if (unresolvedColor !== undefined) {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.unresolvedColor,
        context.pageIndex,
        `${label} border colour ${unresolvedColor}`,
      );
    }
    if (stroke) {
      // `applyImageBorder` sets `box-sizing: border-box`, so the border paints
      // inside the authored box: the stroke's centre line is inset by half its
      // thickness.
      const inset = stroke.thicknessPx / 2;
      primitives.push({
        kind: "rect",
        rect: {
          xPx: rect.xPx + inset,
          yPx: rect.yPx + inset,
          widthPx: Math.max(0, rect.widthPx - stroke.thicknessPx),
          heightPx: Math.max(0, rect.heightPx - stroke.thicknessPx),
        },
        stroke,
      });
    }
  }

  return primitives;
};

/** An anchored or block image fragment placed by the layout engine. */
export const paintImageFragment = (
  fragment: ImageFragment,
  block: ImageBlock,
  context: BuildContext,
): readonly DisplayPrimitive[] =>
  paintImage({
    source: block,
    rect: {
      xPx: fragment.x,
      yPx: fragment.y,
      widthPx: fragment.width,
      heightPx: fragment.height,
    },
    context,
    label: `image block ${String(fragment.blockId)}`,
  });
