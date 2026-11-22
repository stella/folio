import type { CellBorderSpec } from "../types";

const TABLE_BORDER_COLOR = /^#[\da-f]{6}$/iu;
const BLACK = "#000000";
const LARGE_BEVEL_WIDTH_PX = 4;
const MAX_CHANNEL = 255;
const LIGHTNESS_STEPS = 240;
const MIDDLE_TONE_DROP = 70;
const DARK_TONE_DROP = 80;

const THIN_COMPOUND_STROKE_PX = 1;
const THICK_COMPOUND_STROKE_PX = 2;

export type TableBorderSide = "top" | "right" | "bottom" | "left";

function beveledEdge(border: CellBorderSpec | undefined) {
  // The OOXML bridge maps threeDEmboss to ridge and threeDEngrave to groove.
  if (border?.style !== "ridge" && border?.style !== "groove") return undefined;
  const width = border.width ?? 1;
  const color = border.color ?? BLACK;
  if (
    !(width > 0) ||
    !Number.isFinite(width) ||
    !TABLE_BORDER_COLOR.test(color)
  )
    return undefined;
  const shoulder = width >= LARGE_BEVEL_WIDTH_PX ? 2 : 1;
  return { width, shoulder, color, recessed: border.style === "groove" };
}

function largeGapCompoundEdge(border: CellBorderSpec | undefined) {
  if (!border) return undefined;
  const style = border.authoredStyle;
  if (style !== "thinThickLargeGap" && style !== "thickThinLargeGap")
    return undefined;
  const gapPx = border.width ?? 1;
  const color = border.color ?? BLACK;
  if (
    !(gapPx > 0) ||
    !Number.isFinite(gapPx) ||
    !TABLE_BORDER_COLOR.test(color)
  )
    return undefined;

  // These two styles swap the physical strokes. The gap takes w:sz.
  const firstIsThick = style === "thinThickLargeGap";
  const firstPx = firstIsThick
    ? THICK_COMPOUND_STROKE_PX
    : THIN_COMPOUND_STROKE_PX;
  const lastPx = firstIsThick
    ? THIN_COMPOUND_STROKE_PX
    : THICK_COMPOUND_STROKE_PX;
  return { firstPx, lastPx, gapPx, color, widthPx: firstPx + gapPx + lastPx };
}

export function tableCellBorderWidth(
  border: CellBorderSpec | undefined,
): number {
  const compound = largeGapCompoundEdge(border);
  if (compound) return compound.widthPx;
  const edge = beveledEdge(border);
  return edge ? edge.width + edge.shoulder * 2 : (border?.width ?? 0);
}

type Tone = { r: number; g: number; b: number; a: 1 };

function gray(channel: number): Tone {
  return { r: channel, g: channel, b: channel, a: 1 };
}

function rgbTone(hex: string): Tone {
  return {
    r: Number.parseInt(hex.slice(0, 2), 16),
    g: Number.parseInt(hex.slice(2, 4), 16),
    b: Number.parseInt(hex.slice(4, 6), 16),
    a: 1,
  };
}

function embossTones(color: string): readonly [Tone, Tone, Tone] {
  const hex = color.slice(1).toLowerCase();
  if (hex === "000000") return [gray(192), gray(96), gray(0)];
  if (hex === "ffffff") return [gray(255), gray(160), gray(128)];

  const { r, g, b } = rgbTone(hex);
  const low = Math.min(r, g, b);
  const high = Math.max(r, g, b);
  const chroma = high - low;
  const lightness = (low + high) / (2 * MAX_CHANNEL);
  const sourceLevel = Math.round(lightness * LIGHTNESS_STEPS);
  const light = Math.min(
    LIGHTNESS_STEPS,
    sourceLevel >= LIGHTNESS_STEPS / 2
      ? sourceLevel
      : sourceLevel + (sourceLevel >= LIGHTNESS_STEPS / 4 ? 40 : 80),
  );
  const chromaAt = (level: number) => 1 - Math.abs(2 * level - 1);
  const toneAt = (lightnessStep: number): Tone => {
    const level = lightnessStep / LIGHTNESS_STEPS;
    if (chroma === 0) return gray(Math.floor(level * MAX_CHANNEL));
    const scale = chromaAt(level) / chromaAt(lightness);
    const base = MAX_CHANNEL * level - (chroma * scale) / 2;
    const channel = (value: number) => Math.round((value - low) * scale + base);
    return { r: channel(r), g: channel(g), b: channel(b), a: 1 };
  };
  return [
    toneAt(light),
    toneAt(Math.max(0, light - MIDDLE_TONE_DROP)),
    toneAt(Math.max(0, light - DARK_TONE_DROP)),
  ];
}

export function tableCellBorderBands(border: CellBorderSpec | undefined) {
  const compound = largeGapCompoundEdge(border);
  if (compound) {
    const { firstPx, gapPx, lastPx, widthPx } = compound;
    const color = rgbTone(compound.color.slice(1));
    return {
      widthPx,
      bands: [
        { offsetPx: 0, widthPx: firstPx, color },
        { offsetPx: firstPx + gapPx, widthPx: lastPx, color },
      ],
    };
  }
  const edge = beveledEdge(border);
  if (!edge) return undefined;
  const { width, shoulder, color, recessed } = edge;
  const [light, middle, dark] = embossTones(color);
  return {
    widthPx: width + 2 * shoulder,
    bands: [
      { offsetPx: 0, widthPx: shoulder, color: recessed ? dark : light },
      { offsetPx: shoulder, widthPx: width, color: middle },
      {
        offsetPx: shoulder + width,
        widthPx: shoulder,
        color: recessed ? light : dark,
      },
    ],
  };
}
