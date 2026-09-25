/**
 * Tab measurement on a line: the width of content following a tab, decimal
 * tab prefixes, neighbouring tabs, and right-edge clamping.
 */

import { inlineImageBoundingBox } from "../../utils/rotationBoundingBox";
import type { Run, TextRun, FieldRun, MathRun } from "../types";
import { isFloatingImageRun } from "../types";
import { measureTextWidth } from "./measureProvider";
import {
  WIDTH_TOLERANCE,
  runToFontStyle,
  isTextRun,
  isTabRun,
  isImageRun,
  isLineBreakRun,
  isFieldRun,
  fieldMeasureText,
  isMathRun,
  isBlockLayoutImageRun,
} from "./paragraphMeasureShared";

export function measureInlineWidthAfterTab(
  runs: Run[],
  tabIndex: number,
  fieldValues?: ReadonlyMap<number, string>,
): number {
  let width = 0;
  for (let i = tabIndex + 1; i < runs.length; i++) {
    const next = runs[i];
    if (!next || isTabRun(next) || isLineBreakRun(next)) {
      break;
    }
    if (isImageRun(next)) {
      if (isBlockLayoutImageRun(next)) {
        break;
      }
      if (!isFloatingImageRun(next)) {
        width += inlineImageBoundingBox(next).width || 0;
      }
      continue;
    }
    if (isTextRun(next)) {
      width += measureTextWidth(next.text || "", runToFontStyle(next));
    } else if (isFieldRun(next)) {
      width += measureTextWidth(fieldMeasureText(next, fieldValues), runToFontStyle(next));
    } else if (isMathRun(next)) {
      width += measureTextWidth(next.plainText || "[equation]", runToFontStyle(next));
    }
  }
  return width;
}

export function hasFollowingTabOnLine(runs: Run[], tabIndex: number): boolean {
  for (let i = tabIndex + 1; i < runs.length; i++) {
    const next = runs[i];
    if (!next || isLineBreakRun(next)) {
      break;
    }
    if (isImageRun(next) && isBlockLayoutImageRun(next)) {
      break;
    }
    if (isTabRun(next)) {
      return true;
    }
  }
  return false;
}

export function hasPriorTabOnLine(runs: Run[], tabIndex: number): boolean {
  for (let i = tabIndex - 1; i >= 0; i--) {
    const prior = runs[i];
    if (!prior || isLineBreakRun(prior) || (isImageRun(prior) && isBlockLayoutImageRun(prior))) {
      break;
    }
    if (isTabRun(prior)) {
      return true;
    }
  }
  return false;
}

export function canClampTabToRightEdge(
  alignment: string,
  currentLineWidth: number,
  hasPriorTab: boolean,
  followingWidth: number,
  availableWidth: number,
): boolean {
  if (alignment === "start" || alignment === "default") {
    return currentLineWidth > WIDTH_TOLERANCE && (hasPriorTab || followingWidth <= availableWidth);
  }
  return true;
}

/**
 * Width of the inline content preceding the first `.` in the runs that follow
 * a tab, used to anchor `decimal` tab stops. Mirrors `getTextAfterTab` +
 * decimal-prefix measurement in the painter (`renderParagraph.ts`) so the
 * measurer and painter agree on tab advance for decimal stops.
 *
 * Returns 0 when no decimal separator appears before the next tab / line
 * break — `calculateTabWidth` treats that as "no anchor adjustment".
 */
export function measureDecimalPrefixWidthAfterTab(
  runs: Run[],
  tabIndex: number,
  fieldValues?: ReadonlyMap<number, string>,
): number {
  let text = "";
  let firstRun: TextRun | FieldRun | MathRun | undefined;
  for (let i = tabIndex + 1; i < runs.length; i++) {
    const next = runs[i];
    if (!next || isTabRun(next) || isLineBreakRun(next)) {
      break;
    }
    if (isTextRun(next)) {
      text += next.text || "";
      firstRun ??= next;
    } else if (isFieldRun(next)) {
      text += fieldMeasureText(next, fieldValues);
      firstRun ??= next;
    } else if (isMathRun(next)) {
      text += next.plainText || "[equation]";
      firstRun ??= next;
    }
  }
  const decimalIndex = text.indexOf(".");
  if (decimalIndex === -1 || !firstRun) {
    return 0;
  }
  return measureTextWidth(text.slice(0, decimalIndex), runToFontStyle(firstRun));
}
