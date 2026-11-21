import { expect, spyOn, test } from "bun:test";

import * as ops from "@stll/docx-core/ops";

import type { Document, Paragraph } from "../types/document";
import {
  insertedCommentRange,
  plainInsertedParagraphIds,
} from "./canonicalInsertedCommentRange";

const paragraph = (id: number): Paragraph => ({
  type: "paragraph",
  paraId: ops.formatParaId(id),
  content: [{ type: "run", content: [{ type: "text", text: String(id) }] }],
});

test.each([16, 64])(
  "comment endpoints inspect the inserted run in a story of %i blocks",
  (width) => {
    const before: Document = {
      package: {
        document: {
          content: Array.from({ length: width }, (_, index) =>
            paragraph(index + 1),
          ),
        },
      },
    };
    const insertion = {
      type: ops.DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: ops.OP_STORIES.MAIN,
      at: { type: "after", blockId: ops.formatParaId(3) },
      // The authored order deliberately differs from allocation order.
      blocks: [paragraph(300), paragraph(100), paragraph(200)],
    } as const satisfies ops.DocumentOp;
    const earlier = {
      ...insertion,
      at: { type: "before", blockId: ops.formatParaId(1) },
      blocks: [paragraph(400)],
    } as const satisfies ops.DocumentOp;
    const single = ops.applyDocumentOps(before, [insertion]).unwrap().document;
    const compound = ops
      .applyDocumentOps(before, [insertion, earlier])
      .unwrap().document;
    const walk = spyOn(ops, "paragraphsThroughTextBoxes");
    const point = spyOn(ops, "locateParagraphThroughTextBoxes");
    try {
      const range = insertedCommentRange({
        before,
        after: single,
        story: ops.OP_STORIES.MAIN,
        ops: [insertion],
      });
      expect(range.firstAdded).toBe(single.package.document.content[3]);
      expect(range.lastAdded).toBe(single.package.document.content[5]);
      expect(plainInsertedParagraphIds([insertion])).toEqual([100, 200, 300]);
      expect(walk).toHaveBeenCalledTimes(0);
      expect(point).toHaveBeenCalledTimes(1);
      // A compound journal can add the earlier paragraph last.
      const fallback = insertedCommentRange({
        before,
        after: compound,
        story: ops.OP_STORIES.MAIN,
        ops: [insertion, earlier],
      });
      expect(fallback.firstAdded?.paraId).toBe(ops.formatParaId(400));
      expect(fallback.lastAdded?.paraId).toBe(ops.formatParaId(200));
      expect(walk).toHaveBeenCalledTimes(2);
      expect(point).toHaveBeenCalledTimes(1);
    } finally {
      walk.mockRestore();
      point.mockRestore();
    }
  },
);
