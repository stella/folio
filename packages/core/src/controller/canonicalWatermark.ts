import { type DocumentOp, storyPartsOpOf } from "@stll/docx-core/ops";

import { withCanonicalParagraphIds } from "../canonical/canonicalOperations";
import type { Document, HeaderFooter, Watermark } from "../types/document";
import { setDocumentWatermark } from "../watermark/index";

/**
 * The canonical operation that sets (or clears) the document's watermark: the
 * header parts and section references {@link setDocumentWatermark} changes,
 * stale-checked against `doc`, so the edit journals and undoes like any other.
 * A header it creates holds one empty paragraph with a fresh id, as every
 * canonical story does.
 */
export const watermarkOperations = (
  doc: Document,
  watermark: Watermark | undefined,
): DocumentOp[] => {
  const marked = setDocumentWatermark(doc, watermark);
  const headers = new Map<string, HeaderFooter>(marked.package.headers);
  const after: Document = {
    ...marked,
    package: { ...marked.package, headers },
  };
  for (const [rId, header] of headers) {
    if (header.content.length > 0) continue;
    headers.set(rId, {
      ...header,
      content: withCanonicalParagraphIds(
        [{ type: "paragraph", content: [] }],
        after,
      ),
    });
  }
  return [storyPartsOpOf({ before: doc, after })];
};
