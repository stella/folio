/**
 * A paragraph put directly after one that ends a section, before a block that
 * is no paragraph, on the canonical session: it opens the section after the
 * break, which stays where it was, in the saved file too, and undoes. Tracked
 * or suggested, accepting gives the direct edit and rejecting the source.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Paragraph } from "../types/document";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const FOLLOWERS: Record<string, BlockContent> = {
  "a content control": {
    type: "blockSdt",
    properties: { sdtType: "richText", id: 4 },
    content: [paragraph("00000003", "Held.")],
  },
  "a table": {
    type: "table",
    columnWidths: [2000],
    rows: [
      {
        type: "tableRow",
        cells: [
          { type: "tableCell", content: [paragraph("00000003", "Held.")] },
        ],
      },
    ],
  },
};

const open = (buffer: ArrayBuffer) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Editor",
    session: "canonical",
  });

/** Each section's page width and the words of its paragraphs outside blocks. */
const sections = (reviewer: FolioDocxReviewer) =>
  (reviewer.toDocument().package.document.sections ?? []).map(
    ({ properties, content }) => ({
      width: properties.pageWidth,
      blocks: content.map((block) =>
        block.type === "paragraph" ? block.paraId : block.type,
      ),
    }),
  );

const source = (follower: BlockContent) =>
  createDocx({
    package: {
      document: {
        content: [
          {
            ...paragraph("00000001", "First."),
            sectionProperties: { pageWidth: 11_906 },
          },
          follower,
          paragraph("00000002", "Last."),
        ],
        finalSectionProperties: { pageWidth: 12_240 },
      },
    },
  });

const insertAfterSectionEnd = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked" | "suggested",
) => {
  const { applied, skipped, undoHandle } = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [
        {
          id: "op",
          type: "insertAfterBlock",
          blockId: "00000001",
          text: "New.",
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(skipped).toEqual([]);
  expect(applied.map(({ id }) => id)).toEqual(["op"]);
  if (undoHandle === null) throw new Error("The batch has an undo handle.");
  return undoHandle;
};

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text }) => text);

test.each(
  Object.entries(FOLLOWERS).map(([kind, follower]) => ({ kind, follower })),
)("opens the section after the break, before $kind", async ({ follower }) => {
  const reviewer = await open(await source(follower));
  const before = sections(reviewer);
  const undoHandle = insertAfterSectionEnd(reviewer, "direct");
  const after = sections(reviewer);
  expect(after.map(({ width }) => width)).toEqual([11_906, 12_240]);
  expect(after.at(0)).toEqual(before.at(0));
  expect(after.at(1)?.blocks.slice(1)).toEqual(before.at(1)?.blocks ?? []);
  expect(texts(reviewer)).toEqual(["First.", "New.", "Held.", "Last."]);
  const reopened = await open(await reviewer.toBuffer());
  expect(sections(reopened)).toEqual(after);
  expect(reviewer.undoDocumentOperations(undoHandle)).toMatchObject({
    status: "undone",
  });
  expect(sections(reviewer)).toEqual(before);
});

const REVIEWS = {
  "tracked-changes": {
    accept: (reviewer: FolioDocxReviewer) => reviewer.acceptAll(),
    reject: (reviewer: FolioDocxReviewer) => reviewer.rejectAll(),
  },
  suggested: {
    // Accepting a suggestion adopts it as ordinary tracked changes.
    accept: (reviewer: FolioDocxReviewer) => {
      reviewer.acceptSuggestion("op");
      reviewer.acceptAll();
    },
    reject: (reviewer: FolioDocxReviewer) => reviewer.rejectSuggestion("op"),
  },
} as const;

const MODES = ["tracked-changes", "suggested"] as const;

test.each(
  Object.entries(FOLLOWERS).flatMap(([kind, follower]) =>
    MODES.map((mode) => ({ kind, follower, mode })),
  ),
)(
  "$mode, opens the section after the break before $kind, accepts to the direct edit and rejects to the source",
  async ({ follower, mode }) => {
    const buffer = await source(follower);
    const direct = await open(buffer);
    insertAfterSectionEnd(direct, "direct");
    const original = sections(direct);
    const reviewed = async (decision: "accept" | "reject") => {
      const reviewer = await open(buffer);
      insertAfterSectionEnd(reviewer, mode);
      // The break stays on its paragraph, with what it ended.
      expect(sections(reviewer).at(0)).toEqual(original.at(0));
      REVIEWS[mode][decision](reviewer);
      expect(reviewer.getChanges()).toEqual([]);
      return reviewer;
    };
    const accepted = await reviewed("accept");
    expect(sections(accepted)).toEqual(sections(direct));
    expect(texts(accepted)).toEqual(texts(direct));
    const rejected = await reviewed("reject");
    const untouched = await open(buffer);
    expect(sections(rejected)).toEqual(sections(untouched));
    expect(texts(rejected)).toEqual(texts(untouched));
    const undone = await open(buffer);
    const undoHandle = insertAfterSectionEnd(undone, mode);
    const reopened = await open(await undone.toBuffer());
    expect(sections(reopened)).toEqual(
      mode === "suggested" ? sections(untouched) : sections(undone),
    );
    expect(undone.undoDocumentOperations(undoHandle)).toMatchObject({
      status: "undone",
    });
    expect(sections(undone)).toEqual(sections(untouched));
    expect(undone.getChanges()).toEqual([]);
  },
);
