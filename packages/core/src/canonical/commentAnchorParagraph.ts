import {
  locateParagraphThroughTextBoxes,
  type OpStory,
  paragraphsThroughTextBoxes,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";

type CommentAnchorParagraphOptions = {
  document: Document;
  story: OpStory;
  blockId: string | undefined;
};

/** Resolve the first anchor with this exact paragraph ID spelling. */
export const findCommentAnchorParagraph = ({
  document,
  story,
  blockId,
}: CommentAnchorParagraphOptions) => {
  if (blockId !== undefined && blockId !== "") {
    const paragraph = locateParagraphThroughTextBoxes(
      document,
      story,
      blockId,
    )?.paragraph;
    // The locator walks the same story and box order, but ignores case.
    // An exact first hit or a true miss also settles the stricter search.
    if (paragraph === undefined || paragraph.paraId === blockId)
      return paragraph;
  }
  // A case-equivalent predecessor or an unnamed paragraph cannot establish
  // the first exact match.
  return paragraphsThroughTextBoxes(document, story).find(
    ({ paragraph }) => paragraph.paraId === blockId,
  )?.paragraph;
};
