import type { Paragraph } from "../model/document";
import { paragraphLogicalText } from "./offsets";
import { defaultInsertionGap, zeroWidthLeavesAt, type Gap } from "./leaves";
import { DOCUMENT_OP_REFUSAL_REASONS, type DocumentOpRefusalReason } from "./refusal";
import type { TextPosition } from "./types";

const isHighSurrogate = (code: number): boolean => code >= 0xd8_00 && code <= 0xdb_ff;
const isLowSurrogate = (code: number): boolean => code >= 0xdc_00 && code <= 0xdf_ff;

/** What a position that states no `zeroWidthBefore` assumes. */
export type ZeroWidthDefault = "afterAll" | "beforeAll" | "insertion";

/** The gap a position names, or why it names none in this paragraph. */
type ResolveGapOptions = {
  paragraph: Paragraph;
  position: TextPosition;
  fallback: ZeroWidthDefault;
};
export const resolveGap = ({
  paragraph,
  position,
  fallback,
}: ResolveGapOptions): Gap | DocumentOpRefusalReason => {
  const text = paragraphLogicalText(paragraph);
  const { offset, zeroWidthBefore } = position;
  if (!Number.isInteger(offset) || offset < 0 || offset > text.length) {
    return DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET;
  }
  if (
    offset > 0 &&
    offset < text.length &&
    isHighSurrogate(text.charCodeAt(offset - 1)) &&
    isLowSurrogate(text.charCodeAt(offset))
  ) {
    return DOCUMENT_OP_REFUSAL_REASONS.SPLITS_SURROGATE_PAIR;
  }
  const available = zeroWidthLeavesAt(paragraph.content, offset).length;
  if (zeroWidthBefore !== undefined) {
    return Number.isInteger(zeroWidthBefore) && zeroWidthBefore >= 0 && zeroWidthBefore <= available
      ? { offset, zeroWidthBefore }
      : DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET;
  }
  switch (fallback) {
    case "afterAll":
      return { offset, zeroWidthBefore: available };
    case "beforeAll":
      return { offset, zeroWidthBefore: 0 };
    case "insertion":
      return defaultInsertionGap(paragraph.content, offset);
    default: {
      const unreachable: never = fallback;
      return unreachable;
    }
  }
};
