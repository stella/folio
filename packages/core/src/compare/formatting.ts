/** Character-aligned inline-formatting comparison of two text-equal blocks. */

import { canonicalInlinePresentationSegments } from "../internal/compare/inline-presentation";
import type {
  FolioContentBlockIdentity,
  FolioContentInlineComparisonResult,
} from "./content-types";

type InlineFormattingSegmentsOptions = {
  baseBlock: FolioContentBlockIdentity;
  targetBlock: FolioContentBlockIdentity;
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
    // Paragraph styles supply inherited run values; only authored values need
    // a run operation when the style itself changes.
    presentationBasis: baseBlock.styleId === targetBlock.styleId ? "full" : "authored",
  });
