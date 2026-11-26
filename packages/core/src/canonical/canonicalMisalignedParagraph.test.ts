/**
 * A session opening a story aligns a paragraph's text with its editor node
 * when the paragraph is first addressed. A paragraph that does not align is
 * then refused, as it would have refused the whole story at open: an
 * operation on it, or on a neighbour that reaches it, is skipped, never a
 * failure of the batch.
 */
import { Result } from "better-result";
import { describe, expect, spyOn, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioDocumentOperation } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import {
  InlineProjectionCache,
  InlineProjectionError,
} from "./canonicalInlineProjection";

const packageBytes = async () =>
  (
    await prepareCanonicalDocxInput(
      await createDocx(fromMarkdown("First.\n\nSecond.\n\nThird.\n\nFourth.")),
    )
  )
    .unwrap()
    .slice().buffer;

const open = (bytes: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(bytes.slice(0), { session: "canonical" });

/** The ids of the package's blocks, read without misaligning any. */
const blockIds = async (bytes: ArrayBuffer) =>
  (await open(bytes)).snapshot().blocks.map(({ id }) => id);

describe("a paragraph that does not align when first addressed", () => {
  test("refuses each operation that reaches it", async () => {
    const bytes = await packageBytes();
    const [first, second, third] = await blockIds(bytes);
    if (first === undefined || second === undefined || third === undefined)
      throw new Error("Missing fixture");
    const project = InlineProjectionCache.prototype.project;
    const misaligned = spyOn(
      InlineProjectionCache.prototype,
      "project",
    ).mockImplementation(function (this: InlineProjectionCache, args) {
      return args.source.paraId === second
        ? Result.err(new InlineProjectionError({ message: "Misaligned." }))
        : project.call(this, args);
    });
    try {
      const operations: FolioDocumentOperation[] = [
        { id: "merge-into", type: "mergeBlockWithNext", blockId: first },
        { id: "merge-from", type: "mergeBlockWithNext", blockId: second },
        { id: "replace", type: "replaceBlock", blockId: second, text: "New." },
        {
          id: "before",
          type: "insertBeforeBlock",
          blockId: second,
          text: "Before.",
        },
        {
          id: "after",
          type: "insertAfterBlock",
          blockId: second,
          text: "After.",
        },
      ];
      for (const operation of operations) {
        // oxlint-disable-next-line no-await-in-loop -- one session per operation
        const reviewer = await open(bytes);
        const result = reviewer.applyDocumentOperations(
          {
            version: 1,
            mode: "direct",
            atomic: false,
            operations: [operation],
          },
          { undefinedReferences: "refuse" },
        );
        expect(result.applied).toEqual([]);
        expect(result.skipped.map(({ id }) => id)).toEqual([operation.id]);
      }
      // A paragraph the misaligned one does not reach still takes edits.
      const reviewer = await open(bytes);
      const result = reviewer.applyDocumentOperations(
        {
          version: 1,
          mode: "direct",
          atomic: false,
          operations: [
            {
              id: "replace",
              type: "replaceBlock",
              blockId: third,
              text: "Ok.",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
    } finally {
      misaligned.mockRestore();
    }
  });
});
