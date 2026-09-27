/**
 * Rejecting a tracked numbering change restores the list rendering of the
 * numbering it restores. An operation records only the numbering in its
 * `w:pPrChange`, so a reject that left the rendering attrs in place kept the
 * marker of the list it undid: the reviewer showed a list item the saved
 * package does not have.
 */

import { describe, expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const reviewerOf = async (markdown: string): Promise<FolioDocxReviewer> => {
  const { docx } = await ensureParaIds(await createDocx(fromMarkdown(markdown)));
  return FolioDocxReviewer.fromBuffer(docx, { author: "Agent" });
};

/** What a reader sees of each block: kind and list reference. */
const kinds = (reviewer: FolioDocxReviewer): string[] =>
  reviewer
    .getContent()
    .map(
      ({ kind, listReference, text }) =>
        `${kind}${listReference ? ` ${listReference.numId}:${listReference.level}` : ""} ${text}`,
    );

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads "${text}"`);
  }
  return block.id;
};

const numberingChanges: Record<string, (reviewer: FolioDocxReviewer) => FolioDocumentOperation> = {
  "a new list": (reviewer) => ({
    id: "1",
    type: "setBlockParagraphProperties",
    blockId: blockId(reviewer, "Closing remarks."),
    properties: { numbering: { start: "new", kind: "numbered" } },
  }),
  "the existing list": (reviewer) => ({
    id: "1",
    type: "setBlockParagraphProperties",
    blockId: blockId(reviewer, "Closing remarks."),
    properties: {
      numbering: {
        numId:
          reviewer.getContent().find(({ text }) => text === "Deposit on signature")?.listReference
            ?.numId ?? 0,
        level: 0,
      },
    },
  }),
  "numbering removed": (reviewer) => ({
    id: "1",
    type: "setBlockParagraphProperties",
    blockId: blockId(reviewer, "Balance on delivery"),
    properties: { numbering: null },
  }),
};

describe("rejecting a tracked numbering change", () => {
  for (const [name, change] of Object.entries(numberingChanges)) {
    test(`into ${name} reads as before, in the reviewer and the saved package`, async () => {
      const reviewer = await reviewerOf(
        "Intro.\n\n1. Deposit on signature\n2. Balance on delivery\n\nClosing remarks.",
      );
      const before = kinds(reviewer);
      const result = reviewer.applyDocumentOperations({
        version: 1,
        mode: "tracked-changes",
        operations: [change(reviewer)],
      });
      expect(result.issues).toEqual([]);
      expect(kinds(reviewer)).not.toEqual(before);

      reviewer.rejectAll();
      expect(kinds(reviewer)).toEqual(before);
      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      expect(kinds(reopened)).toEqual(before);
    });
  }
});
