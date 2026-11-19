/**
 * Text put into a paragraph that a comment's range runs through, with the
 * comment's text on both sides across the paragraph's edges, carries the
 * comment: a save writes one range start and one range end, so the package
 * reopens with that text covered, and the live document reads it so too.
 * Removing the comment's first stretch afterwards then leaves the same range
 * whether or not the document was saved in between.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationMode,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const MODES = ["direct", "tracked-changes"] as const;

const open = async (bytes: ArrayBuffer | Uint8Array) =>
  FolioDocxReviewer.fromBuffer(bytes, { author: "Test" });

const apply = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationMode,
  operations: FolioDocumentOperation[],
) => {
  const result = reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
};

/** Each paragraph's words and comment range points, in order (ids renumber on a save). */
const ranges = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content.flatMap((child) => {
          if (child.type === "run")
            return child.content.flatMap((item) =>
              item.type === "text" ? [item.text] : [],
            );
          if (
            child.type === "commentRangeStart" ||
            child.type === "commentRangeEnd"
          )
            return [child.type];
          return [];
        })
      : [block.type],
  );

/**
 * A comment on one paragraph, split in two, with a blank paragraph between
 * the halves: the comment's range runs through the blank one.
 */
const commentedHalves = async (mode: FolioDocumentOperationMode) => {
  const reviewer = await open(
    (
      await ensureParaIds(
        await createDocx(fromMarkdown("Opening.\n\nAlpha beta.\n\nClosing.")),
      )
    ).docx,
  );
  const commented = reviewer.snapshot().blocks[1];
  if (!commented) throw new Error("Missing fixture");
  apply(reviewer, "direct", [
    {
      id: "comment",
      type: "commentOnBlock",
      blockId: commented.id,
      comment: { text: "Note." },
    },
  ]);
  apply(reviewer, mode, [
    { id: "split", type: "splitBlock", blockId: commented.id, offset: 6 },
  ]);
  apply(reviewer, mode, [
    { id: "blank", type: "insertAfterBlock", blockId: commented.id, text: "" },
  ]);
  const [, first, blank] = reviewer.snapshot().blocks;
  if (!first || blank?.text !== "") throw new Error("Missing blank paragraph");
  apply(reviewer, mode, [
    { id: "fill", type: "replaceBlock", blockId: blank.id, text: "Inserted." },
  ]);
  return { reviewer, first: first.id };
};

const removeFirst = (
  reviewer: FolioDocxReviewer,
  mode: FolioDocumentOperationMode,
  first: string,
) => {
  apply(reviewer, mode, [
    { id: "remove", type: "deleteBlock", blockId: first },
  ]);
  return ranges(
    mode === "direct"
      ? reviewer.toDocument()
      : reviewer.toResolvedDocument("accept"),
  );
};

describe("text put inside a comment's range across paragraphs", () => {
  for (const mode of MODES) {
    test(`carries the comment, saved or not (${mode})`, async () => {
      const live = await commentedHalves(mode);
      const saved = await commentedHalves(mode);
      const reopened = await open(await saved.reviewer.toBuffer());
      const expected = removeFirst(reopened, mode, saved.first);
      expect(removeFirst(live.reviewer, mode, live.first)).toEqual(expected);
      expect(expected.flat()).toContain("Inserted.");
      const start = expected.flat().indexOf("commentRangeStart");
      expect(expected.flat()[start + 1]).toBe("Inserted.");
    });
  }
});
