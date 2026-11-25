import type { Node as PMNode } from "prosemirror-model";

import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import type { Document } from "../../types/document";

/** Project the reference editor's content into its loaded package. */
export const projectReferenceDocument = (
  document: Document,
  doc: PMNode,
): Document =>
  fromProseDoc(doc, document, { stylesheetSource: { type: "package" } });
