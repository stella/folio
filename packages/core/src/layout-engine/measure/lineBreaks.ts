import { getLineBreakProvider } from "./lineBreakProvider";
import type { LineBreakPolicy } from "./lineBreakProvider";
import { isCjkCodePoint } from "../../utils/scriptSegments";

/**
 * Indices in `text` where a line may end. Latin text breaks after
 * whitespace or hyphen; CJK text can break after every ideograph so a
 * partial line is filled instead of pushing the whole run to the next line.
 */
export const findWordBreaks = (text: string, policy?: LineBreakPolicy): number[] =>
  getLineBreakProvider().findBreaks(text, policy);

export const findGraphemeBreaks = (text: string, policy?: LineBreakPolicy): number[] =>
  getLineBreakProvider().findGraphemeBreaks(text, policy);

export function isBreakChar(char: string | undefined): boolean {
  if (char === undefined) {
    return false;
  }
  if (/\s/u.test(char) || char === "-" || char === "\u00AD" || char === "\u200B") {
    return true;
  }
  if (char >= "\uDC00" && char <= "\uDFFF") {
    return true;
  }
  const codePoint = char.codePointAt(0);
  return codePoint !== undefined && isCjkCodePoint(codePoint);
}
