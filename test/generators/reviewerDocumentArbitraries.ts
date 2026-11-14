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
