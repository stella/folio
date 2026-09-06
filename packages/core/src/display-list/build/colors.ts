/**
 * CSS colour → {@link DisplayColor}.
 *
 * The display list's contract is that colours are already resolved: a backend
 * never sees a custom property, a named OOXML highlight, or `currentColor`.
 * Every such value is collapsed here, at the one place that still knows what
 * the editor's stylesheet means.
 */

import { BLACK, WHITE } from "../primitives";
import type { DisplayColor } from "../types";

/**
 * The editor's document custom properties, materialised. A PDF has no theme
 * and no stylesheet: `--doc-canvas` is the paper, `--doc-canvas-text` the ink,
 * `--suggestion-color` the AI-proposal hue (`renderParagraph.ts`'s
 * `SUGGESTION_COLOR_CSS` fallback).
 */
export const DOC_CANVAS = WHITE;
export const DOC_CANVAS_TEXT = BLACK;
export const SUGGESTION_COLOR: DisplayColor = { r: 0x6d, g: 0x3b, b: 0xd6, a: 1 };

export const TRANSPARENT: DisplayColor = { r: 0, g: 0, b: 0, a: 0 };

/** Word's automatic-colour keywords (`w:color w:val="auto"`). */
const AUTOMATIC_COLOR_VALUES = new Set(["auto", "windowtext", "currentcolor", "inherit"]);

/**
 * OOXML's named highlight palette (§17.18.40) plus the CSS basic colours the
 * bridge can pass through. Kept small on purpose: the bridge resolves named
 * highlights to hex before layout, so this is a safety net for the values that
 * still arrive as names.
 */
const NAMED_COLORS = {
  black: "#000000",
  blue: "#0000ff",
  cyan: "#00ffff",
  darkblue: "#000080",
  darkcyan: "#008080",
  darkgray: "#808080",
  darkgrey: "#808080",
  darkgreen: "#008000",
  darkmagenta: "#800080",
  darkred: "#800000",
  darkyellow: "#808000",
  gray: "#808080",
  grey: "#808080",
  green: "#00ff00",
  lightgray: "#c0c0c0",
  lightgrey: "#c0c0c0",
  magenta: "#ff00ff",
  red: "#ff0000",
  silver: "#c0c0c0",
  white: "#ffffff",
  yellow: "#ffff00",
} as const;

const DOC_CUSTOM_PROPERTIES: Record<string, DisplayColor> = {
  "--doc-canvas": DOC_CANVAS,
  "--doc-canvas-text": DOC_CANVAS_TEXT,
  "--suggestion-color": SUGGESTION_COLOR,
};

const HEX_PATTERN = /^#?(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/iu;
const RGB_PATTERN =
  /^rgba?\(\s*([\d.]+%?)\s*[, ]\s*([\d.]+%?)\s*[, ]\s*([\d.]+%?)\s*(?:[,/]\s*([\d.]+%?)\s*)?\)$/iu;
const VAR_PATTERN = /^var\(\s*(--[\w-]+)\s*(?:,\s*(.+))?\)$/iu;

const clampChannel = (value: number): number => Math.min(255, Math.max(0, Math.round(value)));

const clampAlpha = (value: number): number => Math.min(1, Math.max(0, value));

const parseHex = (raw: string): DisplayColor | undefined => {
  const hex = raw.replace(/^#/u, "");
  const expand = (part: string): number => Number.parseInt(part.repeat(2), 16);
  if (hex.length === 3 || hex.length === 4) {
    // SAFETY: length checked above, so every index below exists.
    const r = expand(hex[0]!);
    const g = expand(hex[1]!);
    const b = expand(hex[2]!);
    const a = hex.length === 4 ? expand(hex[3]!) / 255 : 1;
    return { r, g, b, a };
  }
  if (hex.length === 6 || hex.length === 8) {
    const r = Number.parseInt(hex.slice(0, 2), 16);
    const g = Number.parseInt(hex.slice(2, 4), 16);
    const b = Number.parseInt(hex.slice(4, 6), 16);
    const a = hex.length === 8 ? Number.parseInt(hex.slice(6, 8), 16) / 255 : 1;
    return { r, g, b, a };
  }
  return undefined;
};

const parseComponent = (raw: string, scale: number): number =>
  raw.endsWith("%") ? (Number.parseFloat(raw) / 100) * scale : Number.parseFloat(raw);

const parseRgb = (raw: string): DisplayColor | undefined => {
  const match = RGB_PATTERN.exec(raw);
  if (!match) {
    return undefined;
  }
  // SAFETY: groups 1-3 are non-optional in the pattern.
  const r = clampChannel(parseComponent(match[1]!, 255));
  const g = clampChannel(parseComponent(match[2]!, 255));
  const b = clampChannel(parseComponent(match[3]!, 255));
  const alphaRaw = match[4];
  const a = alphaRaw === undefined ? 1 : clampAlpha(parseComponent(alphaRaw, 1));
  return { r, g, b, a };
};

/**
 * Resolve an authored CSS colour. Returns `undefined` for a value this module
 * cannot turn into concrete channels (`color-mix()`, `oklch()`, an unknown
 * name) so the caller can report it rather than paint a guess.
 */
export const parseDisplayColor = (color: string | undefined): DisplayColor | undefined => {
  if (color === undefined) {
    return undefined;
  }
  const trimmed = color.trim();
  if (trimmed.length === 0) {
    return undefined;
  }
  const lower = trimmed.toLowerCase();

  if (lower === "transparent" || lower === "none") {
    return TRANSPARENT;
  }
  if (AUTOMATIC_COLOR_VALUES.has(lower)) {
    return DOC_CANVAS_TEXT;
  }

  const varMatch = VAR_PATTERN.exec(trimmed);
  if (varMatch) {
    // SAFETY: group 1 is non-optional in the pattern.
    const known = DOC_CUSTOM_PROPERTIES[varMatch[1]!.toLowerCase()];
    if (known) {
      return known;
    }
    return parseDisplayColor(varMatch[2]);
  }

  const named = NAMED_COLORS[lower as keyof typeof NAMED_COLORS];
  if (named !== undefined) {
    return parseHex(named);
  }

  if (HEX_PATTERN.test(trimmed)) {
    return parseHex(trimmed);
  }

  return parseRgb(lower);
};
