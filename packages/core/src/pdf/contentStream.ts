/**
 * Content-stream construction: the graphics operators, and the resource
 * names they refer to.
 *
 * A stream is built as a list of parts rather than a string, because a
 * resource reference (`/F1`, `/Im1`, `/GS1`) is written as a *use* of a font,
 * image or alpha, not as a name. Names are assigned once the whole stream is
 * known, in sorted key order, so nothing about the file depends on the order
 * in which the painter happened to reach a resource. Emitting names as the
 * painter goes would make the bytes depend on traversal order, which is
 * exactly the kind of drift the determinism requirement forbids.
 */

import { panic } from "better-result";
import type { DisplayColor } from "../display-list/types";
import { formatNumber } from "./objects";
import type { PdfMatrix } from "./pageSpace";

export type PdfResourceUse =
  | { readonly kind: "font"; readonly resourceIndex: number }
  | { readonly kind: "image"; readonly imageIndex: number }
  | { readonly kind: "extGState"; readonly alpha: number };

export type ContentPart = { readonly kind: "literal"; readonly text: string } | PdfResourceUse;

/** A glyph plus the correction that follows it in a `TJ` array. */
export type PositionedGlyph = {
  readonly glyphId: number;
  /** Thousandths of text space, subtracted from the pen after the glyph. */
  readonly adjustment: number;
};

/**
 * Text rendering modes. `fillThenStroke` is what an outlined run (`w:outline`)
 * paints.
 */
export const TEXT_RENDER_MODE = { fill: 0, stroke: 1, fillThenStroke: 2 } as const;
export type PdfTextRenderMode = (typeof TEXT_RENDER_MODE)[keyof typeof TEXT_RENDER_MODE];

const COLOR_MAX = 255;

const colorComponents = (color: DisplayColor): string =>
  `${formatNumber(color.r / COLOR_MAX)} ${formatNumber(color.g / COLOR_MAX)} ${formatNumber(color.b / COLOR_MAX)}`;

const glyphHex = (glyphId: number): string => {
  if (!Number.isInteger(glyphId) || glyphId < 0 || glyphId > 0xffff) {
    panic(`glyph id outside Identity-H's two-byte range: ${String(glyphId)}`);
  }
  return glyphId.toString(16).padStart(4, "0").toUpperCase();
};

export type ContentStream = {
  readonly save: () => void;
  readonly restore: () => void;
  readonly concat: (matrix: PdfMatrix) => void;
  readonly setFillColor: (color: DisplayColor) => void;
  readonly setStrokeColor: (color: DisplayColor) => void;
  readonly setLineWidth: (widthPx: number) => void;
  readonly setDash: (pattern: readonly number[], phasePx: number) => void;
  /** Selects an `ExtGState` carrying `/ca` and `/CA` set to `alpha`. */
  readonly setAlpha: (alpha: number) => void;
  readonly appendRect: (xPx: number, yPx: number, widthPx: number, heightPx: number) => void;
  readonly moveTo: (xPx: number, yPx: number) => void;
  readonly lineTo: (xPx: number, yPx: number) => void;
  readonly fill: () => void;
  readonly stroke: () => void;
  readonly clipToCurrentPath: () => void;
  readonly beginText: () => void;
  readonly endText: () => void;
  readonly setFont: (resourceIndex: number, sizePx: number) => void;
  readonly setTextRenderMode: (mode: PdfTextRenderMode) => void;
  readonly setTextMatrix: (matrix: PdfMatrix) => void;
  readonly showGlyphs: (glyphs: readonly PositionedGlyph[]) => void;
  /** Shows one already-encoded byte string, for a simple (base-14) font. */
  readonly showBytes: (bytes: readonly number[]) => void;
  readonly drawImage: (imageIndex: number) => void;
  readonly parts: () => readonly ContentPart[];
};

export const createContentStream = (): ContentStream => {
  const parts: ContentPart[] = [];
  const line = (text: string) => {
    parts.push({ kind: "literal", text: `${text}\n` });
  };
  const use = (resource: PdfResourceUse, suffix: string) => {
    parts.push(resource, { kind: "literal", text: `${suffix}\n` });
  };

  return {
    save: () => line("q"),
    restore: () => line("Q"),
    concat: (matrix) => line(`${matrix.map(formatNumber).join(" ")} cm`),
    setFillColor: (color) => line(`${colorComponents(color)} rg`),
    setStrokeColor: (color) => line(`${colorComponents(color)} RG`),
    setLineWidth: (widthPx) => line(`${formatNumber(widthPx)} w`),
    setDash: (pattern, phasePx) =>
      line(`[${pattern.map(formatNumber).join(" ")}] ${formatNumber(phasePx)} d`),
    setAlpha: (alpha) => use({ kind: "extGState", alpha }, " gs"),
    appendRect: (xPx, yPx, widthPx, heightPx) =>
      line(
        `${formatNumber(xPx)} ${formatNumber(yPx)} ${formatNumber(widthPx)} ${formatNumber(heightPx)} re`,
      ),
    moveTo: (xPx, yPx) => line(`${formatNumber(xPx)} ${formatNumber(yPx)} m`),
    lineTo: (xPx, yPx) => line(`${formatNumber(xPx)} ${formatNumber(yPx)} l`),
    fill: () => line("f"),
    stroke: () => line("S"),
    clipToCurrentPath: () => line("W n"),
    beginText: () => line("BT"),
    endText: () => line("ET"),
    setFont: (resourceIndex, sizePx) =>
      use({ kind: "font", resourceIndex }, ` ${formatNumber(sizePx)} Tf`),
    setTextRenderMode: (mode) => line(`${String(mode)} Tr`),
    setTextMatrix: (matrix) => line(`${matrix.map(formatNumber).join(" ")} Tm`),
    showGlyphs: (glyphs) => {
      let body = "";
      let pending = "";
      for (const { glyphId, adjustment } of glyphs) {
        pending += glyphHex(glyphId);
        const formatted = formatNumber(adjustment);
        if (formatted !== "0") {
          body += `<${pending}>${formatted}`;
          pending = "";
        }
      }
      if (pending !== "") {
        body += `<${pending}>`;
      }
      line(`[${body}] TJ`);
    },
    showBytes: (bytes) => {
      let text = "";
      for (const byte of bytes) {
        if (byte === 0x28 || byte === 0x29 || byte === 0x5c) {
          text += `\\${String.fromCharCode(byte)}`;
        } else if (byte < 0x20 || byte > 0x7e) {
          text += `\\${byte.toString(8).padStart(3, "0")}`;
        } else {
          text += String.fromCharCode(byte);
        }
      }
      line(`(${text}) Tj`);
    },
    drawImage: (imageIndex) => use({ kind: "image", imageIndex }, " Do"),
    parts: () => parts,
  };
};

/** Resource names for one page, assigned from sorted keys. */
export type ContentResources = {
  /** Keyed by font resource index: one face can hold several of them. */
  readonly fontNames: ReadonlyMap<number, string>;
  readonly imageNames: ReadonlyMap<number, string>;
  /** Keyed by the formatted alpha, so two alphas that print alike share a state. */
  readonly extGStateNames: ReadonlyMap<string, string>;
};

export const collectResources = (parts: readonly ContentPart[]): ContentResources => {
  const fonts = new Set<number>();
  const images = new Set<number>();
  const alphas = new Set<string>();
  for (const part of parts) {
    switch (part.kind) {
      case "literal":
        break;
      case "font":
        fonts.add(part.resourceIndex);
        break;
      case "image":
        images.add(part.imageIndex);
        break;
      case "extGState":
        alphas.add(formatNumber(part.alpha));
        break;
      default: {
        const unreachable: never = part;
        panic(`unhandled content part: ${JSON.stringify(unreachable)}`);
      }
    }
  }

  const numberNames = (keys: ReadonlySet<number>, prefix: string) =>
    new Map(
      [...keys]
        .sort((left, right) => left - right)
        .map((key, index) => [key, `${prefix}${String(index + 1)}`] as const),
    );

  return {
    fontNames: numberNames(fonts, "F"),
    imageNames: numberNames(images, "Im"),
    extGStateNames: new Map(
      [...alphas]
        .sort((left, right) => Number(left) - Number(right))
        .map((key, index) => [key, `GS${String(index + 1)}`] as const),
    ),
  };
};

export const renderContentStream = (
  parts: readonly ContentPart[],
  resources: ContentResources,
): string => {
  let out = "";
  for (const part of parts) {
    switch (part.kind) {
      case "literal":
        out += part.text;
        break;
      case "font":
        out += `/${resources.fontNames.get(part.resourceIndex) ?? panic("unnamed font resource")}`;
        break;
      case "image":
        out += `/${resources.imageNames.get(part.imageIndex) ?? panic("unnamed image resource")}`;
        break;
      case "extGState":
        out += `/${resources.extGStateNames.get(formatNumber(part.alpha)) ?? panic("unnamed graphics state")}`;
        break;
      default: {
        const unreachable: never = part;
        panic(`unhandled content part: ${JSON.stringify(unreachable)}`);
      }
    }
  }
  return out;
};
