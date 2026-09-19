import type { Node as PMNode } from "prosemirror-model";

import type { ParagraphFormatting } from "../types/document";
import { expectParagraphAttrs } from "./attrs";
import type { ParagraphAttrs } from "./schema/nodes";

/**
 * Paragraph shapes whose explicit page-break run the editor lays out
 * approximately.
 *
 * Layout splits a paragraph into fragments at its breaks, and a paragraph
 * property that describes the paragraph as a whole then applies to each
 * fragment. That is an approximation, not a loss: the run is projected, it is
 * saved, and the document opens. Folio used to refuse these instead, which
 * meant a file Word opens could not be opened, laid out or exported at all.
 */
const PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES = {
  borders:
    "A bordered paragraph is split at its explicit page-break run, so its border is drawn around each part",
  frame:
    "A framed paragraph is split at its explicit page-break run, so its frame applies to each part",
  outline:
    "An outline paragraph is split at its explicit page-break run, so each part carries the outline level",
  textBoxAnchor:
    "A text-box anchor following an explicit page-break run is hosted by the paragraph's first part",
} as const;

export type PageBreakRunParagraphProjectionReason =
  keyof typeof PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES;

export type PageBreakRunParagraphProjectionDisposition =
  | { status: "exact" }
  | {
      status: "approximate";
      reason: PageBreakRunParagraphProjectionReason;
      message: (typeof PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES)[keyof typeof PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES];
    };

type PageBreakRunParagraphFeatures = {
  attrs: ParagraphAttrs;
  effectiveFrame: ParagraphFormatting["frame"];
  /**
   * A text-box anchor that the paragraph's first page break precedes.
   *
   * Layout splits such a paragraph into fragments at its breaks, and only the
   * first fragment keeps the paragraph's block id, which is what an anchor
   * resolves its host through. An anchor before the first break is therefore
   * laid out exactly; one after it falls back to the first fragment.
   */
  textBoxAnchorAfterPageBreak: boolean;
};

/** Keep source-import and ProseMirror-layout ownership decisions on one predicate. */
export const pageBreakRunParagraphProjectionDispositionForFeatures = ({
  attrs,
  effectiveFrame,
  textBoxAnchorAfterPageBreak,
}: PageBreakRunParagraphFeatures): PageBreakRunParagraphProjectionDisposition => {
  if (
    effectiveFrame !== undefined &&
    effectiveFrame.dropCap !== "drop" &&
    effectiveFrame.dropCap !== "margin"
  ) {
    return {
      status: "approximate",
      reason: "frame",
      message: PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES.frame,
    };
  }
  if (attrs.outlineLevel !== undefined) {
    return {
      status: "approximate",
      reason: "outline",
      message: PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES.outline,
    };
  }
  if (
    attrs.borders !== undefined &&
    Object.values(attrs.borders).some((border) => border !== undefined)
  ) {
    return {
      status: "approximate",
      reason: "borders",
      message: PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES.borders,
    };
  }
  if (textBoxAnchorAfterPageBreak) {
    return {
      status: "approximate",
      reason: "textBoxAnchor",
      message: PAGE_BREAK_RUN_PARAGRAPH_PROJECTION_MESSAGES.textBoxAnchor,
    };
  }
  return { status: "exact" };
};

export const pageBreakRunParagraphProjectionDisposition = (
  paragraph: PMNode,
): PageBreakRunParagraphProjectionDisposition => {
  const attrs = expectParagraphAttrs(paragraph);
  let firstPageBreakPos: number | undefined;
  let textBoxAnchorAfterPageBreak = false;
  paragraph.descendants((descendant, pos) => {
    if (descendant.type.name === "pageBreakRun") {
      firstPageBreakPos ??= pos;
      return false;
    }
    if (descendant.type.name === "textBoxAnchor" && firstPageBreakPos !== undefined) {
      textBoxAnchorAfterPageBreak = true;
      return false;
    }
    return !textBoxAnchorAfterPageBreak;
  });
  return pageBreakRunParagraphProjectionDispositionForFeatures({
    attrs,
    effectiveFrame: attrs._originalFormatting?.frame,
    textBoxAnchorAfterPageBreak,
  });
};
