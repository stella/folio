/**
 * Inline object footprints on a line: rotated block-layout images and math
 * zones.
 */

import type { ImageRun, MathRun } from "../types";
import { ptToPx } from "./measureHelpers";
import { DEFAULT_FONT_SIZE } from "./paragraphMeasureShared";

// Local copies of the painter's rotation helpers (eigenpal #424). Kept in
// sync with `renderParagraph.parseRotationDegrees` /
// `rotatedBoundingBox`; will dedupe once PR #518 + PR #521 land.
function parseRotationDegrees(transform: string | undefined): number {
  if (!transform) {
    return 0;
  }
  const match = /rotate\(\s*(?<degrees>[-\d.]+)\s*deg\s*\)/u.exec(transform);
  if (!match) {
    return 0;
  }
  const raw = Number.parseFloat(match.groups!["degrees"]!);
  if (!Number.isFinite(raw)) {
    return 0;
  }
  return ((raw % 360) + 360) % 360;
}

export function rotatedBlockImageHeight(run: ImageRun): number {
  const deg = parseRotationDegrees(run.transform);
  if (deg === 0 || deg === 180) {
    return run.height;
  }
  if (deg === 90 || deg === 270) {
    return run.width;
  }
  const rad = (deg * Math.PI) / 180;
  const sinA = Math.abs(Math.sin(rad));
  const cosA = Math.abs(Math.cos(rad));
  return run.width * sinA + run.height * cosA;
}

function isStackedMathLocalName(localName: string): boolean {
  switch (localName) {
    case "bar":
    case "d":
    case "eqArr":
    case "f":
    case "groupChr":
    case "limLow":
    case "limUpp":
    case "m":
    case "nary":
    case "rad":
    case "sPre":
    case "sSubSup":
      return true;
    default:
      return false;
  }
}

function hasStackedMathLayout(run: MathRun): boolean {
  if (run.display === "block") {
    return true;
  }
  const tagPattern = /<(?<tag>[A-Za-z_][\w.-]*(?::[A-Za-z_][\w.-]*)?)(?:\s|\/|>)/gu;
  for (const match of run.ommlXml.matchAll(tagPattern)) {
    const tagName = match.groups?.["tag"];
    if (!tagName) {
      continue;
    }
    const colonIndex = tagName.indexOf(":");
    const localName = colonIndex === -1 ? tagName : tagName.slice(colonIndex + 1);
    if (isStackedMathLocalName(localName)) {
      return true;
    }
  }
  return false;
}

export function estimateMathFootprintPx(run: MathRun): number {
  const fontSizePx = ptToPx(run.fontSize ?? DEFAULT_FONT_SIZE);
  return fontSizePx * (hasStackedMathLayout(run) ? 2.4 : 1.25);
}
