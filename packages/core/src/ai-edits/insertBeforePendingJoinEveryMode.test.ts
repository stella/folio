/**
 * A paragraph put before one that a pending join runs words into, or before
 * the paragraph whose words it runs, lands in the same gap in every mode
 * (`insertionGap`): suggested where the tracked insertion goes, before the
 * join, with the joined paragraph's properties, so accepting either leaves
 * what the direct insertion does. The canonical session; the legacy engine,
 * being retired, still reads a suggestion as the document stands.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const MARKDOWN = "Opening words.\n\n1. Joined words.\n\nTarget words.";

type Anchor = "joining" | "target";

type InsertBeforeOptions = {
  session: HeadlessReviewerSession;
  mode: FolioDocumentOperationMode;
  anchor: Anchor;
};

/**
 * The list item's break pending deletion by a tracked merge, so its words
 * run into "Target words.", then a paragraph put before `anchor` in `mode`.
 */
const insertBefore = async ({ session, mode, anchor }: InsertBeforeOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown(MARKDOWN)),
    { author: "Editor", session },
  );
  const [, joining, target] = reviewer.snapshot().blocks;
  if (joining === undefined || target === undefined)
    throw new Error("the fixture's blocks are missing");
  const apply = (
    batchMode: FolioDocumentOperationMode,
    operation: FolioDocumentOperation,
  ) =>
    expect(
      reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: batchMode,
          operations: [operation],
        },
        { undefinedReferences: "refuse" },
      ).skipped,
    ).toEqual([]);
  apply("tracked-changes", {
    id: "merge",
    type: "mergeBlockWithNext",
    blockId: joining.id,
    separator: " ",
  });
  apply(mode, {
    id: "insert",
    type: "insertBeforeBlock",
    blockId: anchor === "joining" ? joining.id : target.id,
    text: "Inserted words.",
  });
  return reviewer;
};

/** The blocks a reader lists, with the properties an insertion copies. */
const blocks = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ text, kind, styleId, listReference }) => ({
    text,
    kind,
    styleId,
    listed: listReference !== undefined,
  }));

const MODES = [
  "direct",
  "tracked-changes",
  "suggested",
] as const satisfies readonly FolioDocumentOperationMode[];

/**
 * "First words." keeps its words and its break goes; "Gone words." goes
 * whole, words and break, so accepting runs "First words." into "Target
 * words.", then a paragraph put before the one going whole in `mode`.
 */
const insertBeforeWholeDeletion = async (mode: FolioDocumentOperationMode) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(
      fromMarkdown(
        "Opening words.\n\nFirst words.\n\nGone words.\n\nTarget words.",
      ),
    ),
    { author: "Editor", session: "canonical" },
  );
  const [, first, gone] = reviewer.snapshot().blocks;
  if (first === undefined || gone === undefined)
    throw new Error("the fixture's blocks are missing");
  for (const [batchMode, operation] of [
    [
      "tracked-changes",
      { id: "delete", type: "deleteBlock", blockId: gone.id },
    ],
    [
      "tracked-changes",
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: first.id,
        separator: " ",
      },
    ],
    [
      mode,
      {
        id: "insert",
        type: "insertBeforeBlock",
        blockId: gone.id,
        text: "Inserted words.",
      },
    ],
  ] as const satisfies readonly (readonly [
    FolioDocumentOperationMode,
    FolioDocumentOperation,
  ])[])
    expect(
      reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: batchMode,
          operations: [operation],
        },
        { undefinedReferences: "refuse" },
      ).skipped,
    ).toEqual([]);
  return reviewer;
};

test.each(MODES)(
  "a paragraph put before one going whole that a join runs words into goes before the join (%s)",
  async (mode) => {
    const reviewer = await insertBeforeWholeDeletion(mode);
    expect(reviewer.snapshot().blocks.at(1)?.text).toBe("Inserted words.");
    reviewer.acceptAll();
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual([
      "Opening words.",
      "Inserted words.",
      "First words. Target words.",
    ]);
  },
);

describe.each(["canonical"] as const)("%s engine", (session) => {
  test.each(["joining", "target"] as const)(
    "a paragraph suggested before the %s paragraph lands where the tracked one does",
    async (anchor) => {
      const tracked = await insertBefore({
        session,
        mode: "tracked-changes",
        anchor,
      });
      const suggested = await insertBefore({
        session,
        mode: "suggested",
        anchor,
      });
      expect(blocks(suggested)).toEqual(blocks(tracked));
      expect(blocks(tracked).at(1)?.text).toBe("Inserted words.");

      const direct = await insertBefore({ session, mode: "direct", anchor });
      for (const reviewer of [tracked, suggested, direct]) reviewer.acceptAll();
      expect(blocks(suggested)).toEqual(blocks(direct));
      expect(blocks(tracked)).toEqual(blocks(direct));
    },
  );
});
