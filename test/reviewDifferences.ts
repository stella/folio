import type { Document } from "../packages/docx-core/src/model/document";

import {
  describePackageDifferences,
  type PackageDifferences,
} from "../scripts/lib/corpus-invariants/model-equality";
import { canonicalReviewBlocks } from "./reviewProjection";

/** Compare authored main-story content; ids and paragraph survivor identity stay exact. */
export const reviewDifferences = (before: Document, after: Document): PackageDifferences => {
  const contentOnly = (document: Document): Document => ({
    package: { document: { content: canonicalReviewBlocks(document.package.document.content) } },
  });
  return describePackageDifferences(contentOnly(before), contentOnly(after));
};
