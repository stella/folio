/**
 * Deleting directly the paragraph that opens its container before a table
 * removes it, as accepting its tracked deletion does.
 */
import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document } from "../types/document";

const SESSIONS = ["canonical"] as const;
const MODES = ["direct", "tracked-changes"] as const;
const TABLE = "| A | B |\n|---|---|\n| 1 | 2 |";

const blockText = (
  block: Document["package"]["document"]["content"][number],
) => {
  if (block.type === "table") return "[table]";
  if (block.type !== "paragraph") return block.type;
  return block.content
    .flatMap((node) =>
      node.type === "run"
        ? node.content.flatMap((leaf) =>
            leaf.type === "text" ? [leaf.text] : [],
          )
        : [],
    )
    .join("");
};

const open = async (
  markdown: string,
  session: (typeof SESSIONS)[number],
): Promise<FolioDocxReviewer> => {
  const { docx } = await ensureParaIds(
    await createDocx(fromMarkdown(markdown)),
  );
  return FolioDocxReviewer.fromBuffer(docx, { author: "Editor", session });
};

const idOf = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((entry) => entry.text === text);
  if (!block) throw new Error(`No block reads "${text}"`);
  return block.id;
};

type DeleteAndSettleOptions = {
  reviewer: FolioDocxReviewer;
  mode: (typeof MODES)[number];
  operations: FolioDocumentOperation[];
};

const deleteAndSettle = async ({
  reviewer,
  mode,
  operations,
}: DeleteAndSettleOptions) => {
  const result = reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  reviewer.acceptAll();
  const saved = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  return saved.toDocument().package.document.content.map(blockText);
};

describe.each(SESSIONS)("%s session", (session) => {
  test.each(MODES)(
    "%s deletion of the opening paragraph before a table removes it",
    async (mode) => {
      const reviewer = await open(`Opening.\n\n${TABLE}\n\nClosing.`, session);
      expect(
        await deleteAndSettle({
          reviewer,
          mode,
          operations: [
            {
              id: "delete",
              type: "deleteBlock",
              blockId: idOf(reviewer, "Opening."),
            },
          ],
        }),
      ).toEqual(["[table]", "Closing."]);
    },
  );

  test.each(MODES)(
    "%s deletion of the opening paragraph before a table the batch inserts removes it",
    async (mode) => {
      const reviewer = await open("Opening.\n\nClosing.", session);
      const blockId = idOf(reviewer, "Opening.");
      expect(
        await deleteAndSettle({
          reviewer,
          mode,
          operations: [
            {
              id: "table",
              type: "insertTable",
              blockId,
              rows: [["Term", "Value"]],
            },
            { id: "delete", type: "deleteBlock", blockId },
          ],
        }),
      ).toEqual(["[table]", "Closing."]);
    },
  );
});
