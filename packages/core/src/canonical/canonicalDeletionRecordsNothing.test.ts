/**
 * Deleting a paragraph that holds no words on the canonical session: an
 * applied deletion changes the document, recording a change where it is
 * tracked or suggested; one with nothing to change or record is skipped.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Document, Paragraph } from "../types/document";

const PARAGRAPH = "00000001";

const documentOf = (content: BlockContent[]): Document => ({
  package: {
    document: {
      content,
      comments: [
        {
          id: 8,
          author: "Earlier",
          content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
        },
      ],
    },
  },
});

const empty = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [],
});

const bookmarked: BlockContent[] = [
  { type: "bookmarkStart", id: 90, name: "_GoBack" },
];

const pending = (content: Paragraph["content"]): Paragraph => ({
  ...empty(PARAGRAPH),
  content,
  pPrMark: { kind: "del", info: { id: 1, author: "Earlier" } },
});

const CASES: Record<string, readonly BlockContent[]> = {
  "the only paragraph": [empty(PARAGRAPH)],
  "the only paragraph, before a bookmark": [empty(PARAGRAPH), ...bookmarked],
  "a paragraph before another": [empty(PARAGRAPH), empty("00000002")],
  "a paragraph whose mark is pending deletion": [
    pending([]),
    empty("00000002"),
    empty("00000003"),
  ],
  "one with a comment's range and a bookmark's end": [
    pending([
      { type: "commentRangeStart", id: 8 },
      { type: "commentRangeEnd", id: 8 },
      { type: "bookmarkEnd", id: 1 },
    ]),
    empty("00000002"),
    empty("00000003"),
  ],
};

const read = (reviewer: FolioDocxReviewer) => ({
  ids: reviewer.getContent().map(({ id }) => id),
  changes: reviewer.getChanges().length,
  suggestions: reviewer.exportPendingSuggestions().length,
});

test.each(
  Object.entries(CASES).flatMap(([place, content]) =>
    (["direct", "tracked-changes", "suggested"] as const).map((mode) => ({
      place,
      content,
      mode,
    })),
  ),
)(
  "deleting $place applies only where it changes something ($mode)",
  async ({ content, mode }) => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(documentOf(structuredClone([...content]))),
      { author: "Editor", session: "canonical" },
    );
    const before = read(reviewer);
    const { applied, skipped } = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [{ id: "op", type: "deleteBlock", blockId: PARAGRAPH }],
      },
      { undefinedReferences: "refuse" },
    );
    const after = read(reviewer);
    if (applied.length === 0) {
      expect(skipped.map(({ id }) => id)).toEqual(["op"]);
      expect(after).toEqual(before);
      return;
    }
    switch (mode) {
      case "direct":
        expect(after.ids).not.toEqual(before.ids);
        return;
      case "tracked-changes":
        expect(after.changes).toBeGreaterThan(before.changes);
        return;
      case "suggested":
        expect(after.suggestions).toBeGreaterThan(before.suggestions);
        return;
      default: {
        const unreachable: never = mode;
        return unreachable;
      }
    }
  },
);

test.each(["direct", "tracked-changes", "suggested"] as const)(
  "deleting the only, empty paragraph before a bookmark is a no-op, as without it (%s)",
  async (mode) => {
    const skips = [];
    for (const content of [
      [empty(PARAGRAPH)],
      [empty(PARAGRAPH), ...bookmarked],
    ]) {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        await createDocx(documentOf(content)),
        { author: "Editor", session: "canonical" },
      );
      const { skipped } = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [{ id: "op", type: "deleteBlock", blockId: PARAGRAPH }],
        },
        { undefinedReferences: "refuse" },
      );
      skips.push(skipped.map(({ reason }) => reason));
    }
    expect(skips).toEqual([["noopOperation"], ["noopOperation"]]);
  },
);
