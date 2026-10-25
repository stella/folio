import type { Document } from "@stll/docx-core";

import { normalizeListModelRendering } from "../docx/listModelNormalization";
import type { DocumentOp } from "./canonicalOperations";

/** Canonical edits share list model normalization with the headless projection. */
export const normalizeCanonicalListRendering: (document: Document) => {
  document: Document;
  inverse: DocumentOp[];
} = normalizeListModelRendering;
