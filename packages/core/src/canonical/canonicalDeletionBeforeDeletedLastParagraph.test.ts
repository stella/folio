/**
 * On the canonical session, tracked deletions of the last paragraph and of
 * the paragraph before it, whose properties the last one's deletion gives it,
 * read alike one at a time and batched, in either order: accepting all
 * equals deleting both directly, rejecting all gives the source back, and
 * each step undoes.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { Paragraph } from "../types/document";

const ITEM = "00000002";
const CLOSING = "00000003";

const paragraph = (
  paraId: string,
  text: string,
  formatting?: Paragraph["formatting"],
): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(formatting === undefined ? {} : { formatting }),
});

const SOURCE = createDocx({
  package: {
    document: {
      content: [
        paragraph("00000001", "Kept."),
        paragraph(ITEM, "Item.", { alignment: "end" }),
        paragraph(CLOSING, "Closing.", { alignment: "center" }),
      ],
    },
  },
});

const open = async (buffer?: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer((buffer ?? (await SOURCE)).slice(0), {
    author: "Editor",
    session: "canonical",
  });

const deleting = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationMode,
  blockIds: readonly string[],
) => {
  const { skipped, undoHandle } = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: blockIds.map((blockId) => ({
        id: `delete-${blockId}`,
        type: "deleteBlock" as const,
        blockId,
      })),
    },
    { undefinedReferences: "refuse" },
  );
  expect(skipped).toEqual([]);
  if (undoHandle === null) throw new Error("The batch has an undo handle.");
  return undoHandle;
};

/** What a reader sees: the words, each paragraph's properties, the changes. */
const view = (reviewer: FolioDocxReviewer) => ({
  content: reviewer.getContent().map(({ text }) => text),
  properties: reviewer.toDocument().package.document.content.map((block) =>
    block.type === "paragraph"
      ? {
          formatting: block.formatting,
          previous: block.propertyChanges?.map(
            ({ previousFormatting }) => previousFormatting,
          ),
          mark: block.pPrMark?.kind,
        }
      : block.type,
  ),
  changes: reviewer
    .getChanges()
    .map(({ type, blockId, text }) => ({ type, blockId, text })),
});

const ORDERS = {
  "one at a time, the last paragraph first": [[CLOSING], [ITEM]],
  "one at a time, the item first": [[ITEM], [CLOSING]],
  "batched, the last paragraph first": [[CLOSING, ITEM]],
  "batched, the item first": [[ITEM, CLOSING]],
} as const satisfies Record<string, readonly (readonly string[])[]>;

test("tracked deletions of a paragraph and the last one after it read alike in every order", async () => {
  const direct = await open();
  deleting(direct, "direct", [ITEM, CLOSING]);
  const source = view(await open());
  const views: unknown[] = [];
  for (const batches of Object.values(ORDERS)) {
    const tracked = await open();
    for (const blockIds of batches) {
      const before = view(tracked);
      const undoHandle = deleting(tracked, "tracked-changes", blockIds);
      const after = view(tracked);
      expect(tracked.undoDocumentOperations(undoHandle)).toMatchObject({
        status: "undone",
      });
      expect(view(tracked)).toEqual(before);
      deleting(tracked, "tracked-changes", blockIds);
      expect(view(tracked)).toEqual(after);
    }
    views.push(view(tracked));
    const accepted = await open(await tracked.toBuffer());
    accepted.acceptAll();
    expect(view(accepted)).toEqual(view(direct));
    tracked.rejectAll();
    expect(view(tracked)).toEqual(source);
  }
  for (const other of views.slice(1)) expect(other).toEqual(views.at(0));
});
