import type { StrokeDashCycle } from "../../display-list/types";
import { pointsToPixels } from "../../utils/units";
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
const DOT_DASH_DASH_WIDTH_FACTOR = 7;
const DOT_DASH_DOT_WIDTH_FACTOR = 3;
const DASH_WIDTH_FACTOR = 4;
const DOT_DOT_DASH_DASH_WIDTH_FACTOR = 6;
const DOT_DOT_DASH_DOT_WIDTH_FACTOR = 2;
const MIN_SMALL_GAP_STROKE_PX = pointsToPixels(0.25);

type CompoundStyle = Extract<
  CellBorderSpec["authoredStyle"],
  `${"thinThick" | "thickThin"}${string}`
>;
type CompoundSegment = number | "authored" | "authoredStroke";

const COMPOUND_SEGMENTS = {
  thinThickSmallGap: [
    "authoredStroke",
    THIN_COMPOUND_STROKE_PX,
    THIN_COMPOUND_STROKE_PX,
  ],
  thickThinSmallGap: [
    THIN_COMPOUND_STROKE_PX,
    THIN_COMPOUND_STROKE_PX,
    "authoredStroke",
  ],
  thinThickLargeGap: [
    THICK_COMPOUND_STROKE_PX,
    "authored",
    THIN_COMPOUND_STROKE_PX,
  ],
  thickThinLargeGap: [
    THIN_COMPOUND_STROKE_PX,
    "authored",
    THICK_COMPOUND_STROKE_PX,
  ],
} as const satisfies Record<
  CompoundStyle,
  readonly [CompoundSegment, CompoundSegment, CompoundSegment]
>;

export function hasCompoundTableBorderBands(
  style: unknown,
): style is CompoundStyle {
  return typeof style === "string" && Object.hasOwn(COMPOUND_SEGMENTS, style);
}

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

function compoundEdge(border: CellBorderSpec | undefined) {
  const style = border?.authoredStyle;
  if (!border || !hasCompoundTableBorderBands(style)) return undefined;

  const width = border.width ?? 1;
  const color = border.color ?? BLACK;
  if (
    !(width > 0) ||
    !Number.isFinite(width) ||
    !TABLE_BORDER_COLOR.test(color)
  )
    return undefined;

  const segmentWidth = (segment: CompoundSegment): number => {
    if (segment === "authored") return width;
    if (segment === "authoredStroke")
      return Math.max(width, MIN_SMALL_GAP_STROKE_PX);
    return segment;
  };
  const [first, gap, last] = COMPOUND_SEGMENTS[style];
  const firstPx = segmentWidth(first);
  const gapPx = segmentWidth(gap);
  const lastPx = segmentWidth(last);
  return { firstPx, lastPx, gapPx, color, widthPx: firstPx + gapPx + lastPx };
}

type TableCellBorderDash = {
  widthPx: number;
  dashArrayPx: StrokeDashCycle;
};

export function tableCellBorderDash(
  border: CellBorderSpec | undefined,
): TableCellBorderDash | undefined {
  if (!border) return undefined;
  const widthPx = border.width ?? 1;
  if (!(widthPx > 0) || !Number.isFinite(widthPx)) return undefined;

  switch (border.authoredStyle) {
    case "dashed": {
      const length = widthPx * DASH_WIDTH_FACTOR;
      return { widthPx, dashArrayPx: [length, length] };
    }
    case "dashSmallGap":
      return {
        widthPx,
        dashArrayPx: [widthPx * DASH_WIDTH_FACTOR, widthPx],
      };
    case "dotDash": {
      const shortPx = widthPx * DOT_DASH_DOT_WIDTH_FACTOR;
      return {
        widthPx,
        dashArrayPx: [
          widthPx * DOT_DASH_DASH_WIDTH_FACTOR,
          shortPx,
          shortPx,
          shortPx,
        ],
      };
    }
    case "dotDotDash": {
      const shortPx = widthPx * DOT_DOT_DASH_DOT_WIDTH_FACTOR;
      return {
        widthPx,
        dashArrayPx: [
          widthPx * DOT_DOT_DASH_DASH_WIDTH_FACTOR,
          shortPx,
          shortPx,
          shortPx,
          shortPx,
          shortPx,
        ],
      };
    }
    default:
      return undefined;
  }
}

export function tableCellBorderWidth(
  border: CellBorderSpec | undefined,
): number {
  const compound = compoundEdge(border);
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
  const compound = compoundEdge(border);
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
