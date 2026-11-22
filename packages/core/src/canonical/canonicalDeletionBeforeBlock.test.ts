/**
 * A direct deletion of a container's first paragraph, with range boundaries in
 * it and a block after it that offers no paragraph to run into: the boundaries
 * go to the first paragraph after the block, as accepting the tracked deletion
 * leaves them.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Paragraph, Run } from "../types/document";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const paragraph = (
  paraId: string,
  content: Paragraph["content"],
): Paragraph => ({ type: "paragraph", paraId, content });

const COMMENTED: Paragraph["content"] = [
  run("a"),
  { type: "commentRangeStart", id: 8 },
  { type: "commentRangeEnd", id: 8 },
];

const BOOKMARKED: Paragraph["content"] = [
  { type: "bookmarkStart", id: 3, name: "mark" },
  run("a"),
  { type: "bookmarkEnd", id: 3 },
];

const BLOCKS: Record<string, BlockContent> = {
  "an unsupported block": { type: "preservedBlock", xml: "<w:altChunk/>" },
  "a control opening with a table": {
    type: "blockSdt",
    properties: { sdtType: "richText", id: 1, tag: "clause" },
    content: [
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [paragraph("00000004", [run("cell")])],
              },
            ],
          },
        ],
      },
      paragraph("00000005", [run("after the table")]),
    ],
  },
};

const opened = async (
  words: Paragraph["content"],
  block: BlockContent,
): Promise<FolioDocxReviewer> => {
  const bytes = await createDocx({
    package: {
      document: {
        content: [
          paragraph("00000001", words),
          block,
          paragraph("00000002", [run("z")]),
        ],
        comments: [
          { id: 8, author: "A", content: [paragraph("7FFFFFF0", [])] },
        ],
      },
    },
  });
  return FolioDocxReviewer.fromBuffer(bytes, {
    author: "Editor",
    session: "canonical",
  });
};

const remove = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
) => {
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [{ id: "op", type: "deleteBlock", blockId: "00000001" }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  return result;
};

const body = (reviewer: FolioDocxReviewer): BlockContent[] =>
  reviewer.toDocument().package.document.content;

for (const [kind, words] of Object.entries({
  "a comment range": COMMENTED,
  "a bookmark": BOOKMARKED,
}))
  for (const [shape, block] of Object.entries(BLOCKS))
    test(`a direct deletion of a paragraph with ${kind} before ${shape} leaves what accepting the tracked one does`, async () => {
      const tracked = await opened(words, block);
      remove(tracked, "tracked-changes");
      tracked.acceptAll();

      const direct = await opened(words, block);
      const before = body(direct);
      const { undoHandle } = remove(direct, "direct");
      expect(body(direct)).toEqual(body(tracked));

      if (undoHandle === null)
        throw new TypeError("A direct deletion can be undone.");
      direct.undoDocumentOperations(undoHandle);
      expect(body(direct)).toEqual(before);
    });
