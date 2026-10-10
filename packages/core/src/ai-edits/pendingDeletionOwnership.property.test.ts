import { expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import { paragraph, table } from "../docx/server/build";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import type { Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle, sourceDocumentOf } from "./snapshot";

type OperationContext = { blockId: string; text: string };
const operations = {
  deleteBlock: ({ blockId }: OperationContext) => ({ id: "delete", type: "deleteBlock", blockId }),
  replaceBlock: ({ blockId }: OperationContext) => ({
    id: "replace",
    type: "replaceBlock",
    blockId,
    text: "final",
  }),
  replaceInBlock: ({ blockId, text }: OperationContext) => ({
    id: "replace",
    type: "replaceInBlock",
    blockId,
    find: text,
    replace: "final",
  }),
  replaceRange: ({ blockId, text }: OperationContext) => {
    const range = createFolioAITextRangeHandle({
      blockId,
      text,
      startOffset: 0,
      endOffset: text.length,
    });
    if (!range) throw new Error("fixture range missing");
    return { id: "replace", type: "replaceRange", range, replace: "final" };
  },
  splitBlock: ({ blockId, text }: OperationContext) => ({
    id: "split",
    type: "splitBlock",
    blockId,
    offset: text.length - 2,
  }),
  mergeBlockWithNext: ({ blockId }: OperationContext) => ({
    id: "merge",
    type: "mergeBlockWithNext",
    blockId,
  }),
  deleteTable: ({ blockId }: OperationContext) => ({ id: "delete", type: "deleteTable", blockId }),
  deleteTableRow: ({ blockId }: OperationContext) => ({
    id: "delete",
    type: "deleteTableRow",
    blockId,
  }),
  deleteTableColumn: ({ blockId }: OperationContext) => ({
    id: "delete",
    type: "deleteTableColumn",
    blockId,
  }),
} as const satisfies Record<string, (context: OperationContext) => FolioDocumentOperation>;

const hiddenMarks = (doc: PMNode) => {
  const marks: unknown[] = [];
  doc.descendants((node) => {
    if (node.text === "hidden") marks.push(node.marks.map((mark) => mark.toJSON()));
  });
  return marks;
};
const content = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ kind, text }) => ({ kind, text }));

// The former sequence generator did not cross a save with nested run revisions
// before a later deletion spanned both visible and already deleted content.
for (const [kind, operation] of Object.entries(operations)) {
  test(
    `${kind} preserves pending deletion ownership across saves`,
    async () => {
      await assertProperty(
        fc.asyncProperty(
          fc.record({
            prefix: fc.integer({ min: 1, max: 30 }),
            suffix: fc.integer({ min: 1, max: 30 }),
            ancestor: fc.constantFrom("insertion", "deletion"),
            author: fc.constantFrom("Reviewer", "Other Reviewer"),
          }),
          async ({ prefix, suffix, ancestor, author }) => {
            const p = {
              type: "paragraph",
              content: [
                { type: "run", content: [{ type: "text", text: "a".repeat(prefix) }] },
                {
                  type: ancestor,
                  info: { id: 11, author: "First Reviewer", date: "2026-01-02T03:04:05Z" },
                  content: [
                    {
                      type: "deletion",
                      info: { id: 12, author: "Second Reviewer", date: "2026-01-03T03:04:05Z" },
                      content: [{ type: "run", content: [{ type: "text", text: "hidden" }] }],
                    },
                  ],
                },
                { type: "run", content: [{ type: "text", text: "z".repeat(suffix) + " end" }] },
              ],
            } as const satisfies Paragraph;
            const document = createEmptyDocument();
            if (kind.startsWith("deleteTable")) {
              const grid = table({
                rows: [
                  ["cell", "other"],
                  ["lower", "last"],
                ],
              });
              const cell = grid.rows.at(0)?.cells.at(0);
              expect(cell).toBeDefined();
              if (!cell) throw new Error("fixture cell missing");
              cell.content = [p];
              document.package.document.content = [grid, paragraph("following")];
            } else document.package.document.content = [p, paragraph("following")];
            const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(document), {
              author,
            });
            const before = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
            before.rejectAll();
            const block = reviewer.getContent().at(0);
            if (!block) throw new Error("fixture block missing");
            const originalMarks = hiddenMarks(sourceDocumentOf(reviewer.snapshot()));
            expect(originalMarks).toHaveLength(1);
            const result = reviewer.applyDocumentOperations({
              version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
              mode: "tracked-changes",
              operations: [operation({ blockId: block.id, text: block.text })],
            });
            expect(result.applied).toHaveLength(1);
            expect(hiddenMarks(sourceDocumentOf(reviewer.snapshot()))).toEqual(originalMarks);
            const saved = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
            saved.rejectAll();
            expect(content(saved)).toEqual(content(before));
          },
        ),
        { numRuns: 20, id: "${kind} preserves pending deletion ownership across saves" },
      );
    },
    propertyTestTimeout(30_000),
  );
}
