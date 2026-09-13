/** Character-aligned inline-formatting comparison of two text-equal blocks. */

import { canonicalInlinePresentationSegments } from "../internal/compare/inline-presentation";
import type { FolioContentBlock, FolioContentInlineComparisonResult } from "./content-types";

type InlineFormattingSegmentsOptions = {
  baseBlock: FolioContentBlock;
  targetBlock: FolioContentBlock;
  maxSegments: number;
};

/** Malformed run streams take precedence over the output budget. */
export const inlineFormattingSegments = ({
  baseBlock,
  targetBlock,
  maxSegments,
}: InlineFormattingSegmentsOptions): FolioContentInlineComparisonResult =>
  canonicalInlinePresentationSegments({
    baseBlock,
    revisedBlock: targetBlock,
    maxSegments,
  });
