import type { Document } from "@stll/docx-core";

import {
  type ListNormalizationCache,
  normalizeListModelRendering,
} from "../docx/listModelNormalization";
import type { DocumentOp } from "./canonicalOperations";

/** Canonical edits share list model normalization with the headless projection. */
export const normalizeCanonicalListRendering = (
  document: Document,
  cache?: ListNormalizationCache,
): {
  document: Document;
  inverse: DocumentOp[];
} =>
  normalizeListModelRendering(document, cache === undefined ? {} : { cache });
