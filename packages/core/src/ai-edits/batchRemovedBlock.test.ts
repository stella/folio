/**
 * An operation whose block an earlier operation of the same batch removed is
 * refused with a reason (`overlappingOperation`), not applied to whatever took
 * the block's place and not thrown out of the batch.
 */

import { describe, expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { getFolioDocumentOperationRemovedBlockIds } from "./removed-blocks";

const buildReviewer = async (): Promise<FolioDocxReviewer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("First clause."),
    paragraph("Second clause."),
    paragraph("Third clause."),
  ];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return FolioDocxReviewer.fromBuffer(
    docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer,
  );
};

describe("a batch that removes a block and then formats it", () => {
  test("refuses the formatting and leaves the next paragraph alone", async () => {
    const reviewer = await buildReviewer();
    const second = reviewer.getContent().find(({ text }) => text === "Second clause.");
    if (!second) {
      throw new Error("the fixture paragraph is missing");
    }
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [
          { id: "delete", type: "deleteBlock", blockId: second.id },
          {
            id: "center",
            type: "setBlockParagraphProperties",
            blockId: second.id,
            properties: { alignment: "center" },
          },
        ],
      },
      { undefinedReferences: "refuse" },
    );

    expect(result.applied.map(({ id }) => id)).toEqual(["delete"]);
    expect(result.skipped).toEqual([
      {
        id: "center",
        reason: "overlappingOperation",
        message: 'operation "delete", earlier in this batch, already claims its target.',
      },
    ]);
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Third clause.",
    ]);
    expect(
      reviewer.getContent().some((block) => {
        switch (block.kind) {
          case "paragraph":
          case "heading":
          case "listItem":
            return block.directAlignment !== undefined;
          case "diagnostic":
            return false;
          default: {
            const unreachable: never = block;
            return unreachable;
          }
        }
      }),
    ).toBe(false);
    await reviewer.toBuffer();
  });

  test.each(["direct", "tracked-changes"] as const)(
    "refuses an edit of the block an earlier merge removes (%s)",
    async (mode) => {
      const reviewer = await buildReviewer();
      const [first, second] = reviewer.getContent();
      if (!first || !second) throw new Error("the fixture paragraphs are missing");

      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            { id: "merge", type: "mergeBlockWithNext", blockId: first.id, separator: " " },
            {
              id: "format-removed",
              type: "setBlockParagraphProperties",
              blockId: second.id,
              properties: { alignment: "center" },
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );

      expect(result.applied.map(({ id }) => id)).toEqual(["merge"]);
      expect(result.skipped).toEqual([
        {
          id: "format-removed",
          reason: "overlappingOperation",
          message: 'operation "merge", earlier in this batch, already claims its target.',
        },
      ]);
    },
  );
});

describe("getFolioDocumentOperationRemovedBlockIds", () => {
  const blocks = [{ id: "A" }, { id: "B" }, { id: "C" }];

  test("a merge removes the block after its anchor", () => {
    expect([
      ...getFolioDocumentOperationRemovedBlockIds(blocks, {
        type: "mergeBlockWithNext",
        blockId: "A",
      } as never),
    ]).toEqual(["B"]);
  });

  test("a merge does not remove its anchor", () => {
    expect(
      getFolioDocumentOperationRemovedBlockIds(blocks, {
        type: "mergeBlockWithNext",
        blockId: "A",
      } as never).has("A"),
    ).toBe(false);
  });

  test("a merge with an unknown anchor removes nothing", () => {
    expect([
      ...getFolioDocumentOperationRemovedBlockIds(blocks, {
        type: "mergeBlockWithNext",
        blockId: "missing",
      } as never),
    ]).toEqual([]);
  });
});
