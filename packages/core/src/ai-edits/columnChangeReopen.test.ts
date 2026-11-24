/**
 * A tracked column edit marks one cell per row. The save writes each cell's
 * marker as a revision of its own, so the batch gives each cell its own id:
 * the reviewer lists, resolves, saves and reopens the same changes.
 */

import { describe, expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const MARKDOWN = [
  "# Schedule",
  "| Item | Price |\n|---|---|\n| Term | Value |\n| Period | 12 months |",
  "Closing.",
].join("\n\n");

const TYPES = ["insertTableColumn", "deleteTableColumn"] as const;

const isCellChange = ({ type }: { type: string }) =>
  type === "cellInserted" || type === "cellDeleted";

const cellChanges = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getChanges()
    .filter(isCellChange)
    .map(({ type, author, text, blockId }) => ({
      type,
      author,
      text,
      blockId,
    }));

const SESSIONS = ["canonical"] as const;
type Session = (typeof SESSIONS)[number];

const build = async (
  session: Session,
  type: (typeof TYPES)[number],
): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown(MARKDOWN)),
    { author: "AI", session },
  );
  const blockId = reviewer
    .getContent()
    .find((block) => block.text === "Value")?.id;
  if (blockId === undefined)
    throw new Error("fixture must expose the Value cell");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        type === "insertTableColumn"
          ? { id: "op", type, blockId, position: "after" }
          : { id: "op", type, blockId },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  // Each cell's change has an id of its own among the batch's. A column the
  // canonical session inserts also records the grid and table properties it
  // changes, with ids of their own.
  const cellIds = reviewer
    .getChanges()
    .filter(isCellChange)
    .map(({ id }) => id);
  expect(new Set(cellIds).size).toBe(3);
  expect(result.applied[0]?.revisionIds).toEqual(
    expect.arrayContaining(cellIds),
  );
  return reviewer;
};

const reopen = async (
  session: Session,
  reviewer: FolioDocxReviewer,
): Promise<FolioDocxReviewer> =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), { session });

describe.each(SESSIONS)("tracked column changes (%s)", (session) => {
  test.each(TYPES)(
    "%s lists one change per cell, the same after a save",
    async (type) => {
      const reviewer = await build(session, type);
      const before = cellChanges(reviewer);
      expect(before).toHaveLength(3);
      const reopened = await reopen(session, reviewer);
      expect(cellChanges(reopened)).toEqual(before);
      for (const current of [reviewer, reopened]) {
        const ids = current
          .getChanges()
          .filter(isCellChange)
          .map(({ id }) => id);
        expect(new Set(ids).size).toBe(ids.length);
      }
    },
  );

  test.each(TYPES)(
    "%s: accepting one cell leaves the others pending",
    async (type) => {
      const reviewer = await build(session, type);
      const saved = await reopen(session, reviewer);
      for (const current of [reviewer, saved]) {
        const first = current.getChanges().find(isCellChange);
        if (!first) throw new Error("no cell change");
        expect(current.acceptChange(first)).toBe(true);
        expect(cellChanges(current)).toHaveLength(2);
      }
      expect(cellChanges(saved)).toEqual(cellChanges(reviewer));
    },
  );
});
