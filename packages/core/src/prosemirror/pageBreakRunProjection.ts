import type { Node as PMNode } from "prosemirror-model";

import { expectParagraphAttrs } from "./attrs";
import type { ParagraphAttrs } from "./schema/nodes";

const PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES = {
  borders: "A bordered paragraph containing an explicit page-break run cannot be projected",
  frame: "A framed paragraph containing an explicit page-break run cannot be projected",
  outline: "An outline paragraph containing an explicit page-break run cannot be projected",
  textBoxAnchor:
    "A paragraph containing both an explicit page-break run and a text-box anchor cannot be projected",
} as const;

export type PageBreakRunParagraphProjectionReason =
  keyof typeof PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES;

export type PageBreakRunParagraphProjectionDisposition =
  | { status: "supported" }
  | {
      status: "unsupported";
      reason: PageBreakRunParagraphProjectionReason;
      message: (typeof PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES)[keyof typeof PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES];
    };

type PageBreakRunParagraphFeatures = {
  attrs: ParagraphAttrs;
  hasTextBoxAnchor: boolean;
};

/** Keep source-import and ProseMirror-layout ownership decisions on one predicate. */
export const pageBreakRunParagraphProjectionDispositionForFeatures = ({
  attrs,
  hasTextBoxAnchor,
}: PageBreakRunParagraphFeatures): PageBreakRunParagraphProjectionDisposition => {
  const frame = attrs._originalFormatting?.frame;
  if (frame !== undefined && frame.dropCap !== "drop" && frame.dropCap !== "margin") {
    return {
      status: "unsupported",
      reason: "frame",
      message: PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES.frame,
    };
  }
  if (attrs.outlineLevel !== undefined) {
    return {
      status: "unsupported",
      reason: "outline",
      message: PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES.outline,
    };
  }
  if (
    attrs.borders !== undefined &&
    Object.values(attrs.borders).some((border) => border !== undefined)
  ) {
    return {
      status: "unsupported",
      reason: "borders",
      message: PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES.borders,
    };
  }
  if (hasTextBoxAnchor) {
    return {
      status: "unsupported",
      reason: "textBoxAnchor",
      message: PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES.textBoxAnchor,
    };
  }
  return { status: "supported" };
};

export const pageBreakRunParagraphProjectionDisposition = (
  paragraph: PMNode,
): PageBreakRunParagraphProjectionDisposition => {
  const attrs = expectParagraphAttrs(paragraph);
  let hasTextBoxAnchor = false;
  paragraph.descendants((descendant) => {
    hasTextBoxAnchor ||= descendant.type.name === "textBoxAnchor";
    return !hasTextBoxAnchor;
  });
  return pageBreakRunParagraphProjectionDispositionForFeatures({ attrs, hasTextBoxAnchor });
};
