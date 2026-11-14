/** Documents a headless reviewer reads, for properties that compare its engines. */
import fc from "fast-check";

import type { Document } from "../../packages/docx-core/src/model/document";
import { reviewDocumentArbitrary } from "../../packages/docx-core/src/ops/__tests__/documentArbitraries";
import { packageDocumentArbitrary } from "./packageOperationArbitraries";

/**
 * Every story a reviewer lists (headers, footers, notes, a comment thread,
 * content controls and a table), or a body with review marks of every shape
 * the review properties draw.
 */
export const reviewerDocumentArbitrary: fc.Arbitrary<Document> = fc.oneof(
  packageDocumentArbitrary,
  reviewDocumentArbitrary,
);

/**
 * Mostly bodies with review marks of every shape, for properties about the
 * body's edits and resolutions; sometimes a package with every story, whose
 * own edits and resolutions the story properties hold to their laws.
 */
export const bodyReviewerDocumentArbitrary: fc.Arbitrary<Document> = fc.oneof(
  { weight: 3, arbitrary: reviewDocumentArbitrary },
  { weight: 1, arbitrary: packageDocumentArbitrary },
);
