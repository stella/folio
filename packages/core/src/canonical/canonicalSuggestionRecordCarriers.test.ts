/**
 * Accept-all on the canonical session resolves ordinary changes only: every
 * record a pending suggestion made stays staged, whichever carrier the review
 * surface holds it on (a revision nested in another, a row or cell marker, a
 * table's property change). Each case pairs the suggestion with an ordinary
 * change, so accept-all has work to do beside it.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { fromMarkdown } from "../markdown/fromMarkdown";

const open = async () => {
  const document = fromMarkdown("Opening words.\n\nClosing words.");
  document.package.document.content.splice(1, 0, {
    type: "table",
    rows: ["Item", "Widget"].map((text) => ({
      type: "tableRow",
      cells: [
        { type: "tableCell", content: [paragraph(text)] },
        { type: "tableCell", content: [paragraph(`${text} price`)] },
      ],
    })),
  });
  return FolioDocxReviewer.fromBuffer(await createDocx(document), {
    author: "Editor",
    session: "canonical",
  });
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "tracked-changes" | "suggested",
  operations: FolioDocumentOperation[],
) => {
  const result = reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  return result;
};

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((row) => row.text === text);
  if (block === undefined) throw new Error(`No block reads "${text}".`);
  return block.id;
};

/** What a save writes, every paragraph's words in order. */
const savedTexts = async (reviewer: FolioDocxReviewer) =>
  (
    await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
      session: "canonical",
    })
  )
    .getContent()
    .map(({ text }) => text);

type Staging = (reviewer: FolioDocxReviewer) => void;

const SUGGESTIONS: Record<string, Staging> = {
  "a deletion of another suggestion's inserted paragraph": (reviewer) => {
    apply(reviewer, "suggested", [
      {
        id: "insert",
        type: "insertBeforeBlock",
        blockId: blockId(reviewer, "Opening words."),
        text: "Proposed clause.",
      },
    ]);
    apply(reviewer, "suggested", [
      {
        id: "delete",
        type: "deleteBlock",
        blockId: blockId(reviewer, "Proposed clause."),
      },
    ]);
  },
  "a table column": (reviewer) => {
    apply(reviewer, "suggested", [
      {
        id: "column",
        type: "insertTableColumn",
        blockId: blockId(reviewer, "Item"),
        position: "after",
      },
    ]);
  },
  "a table row's deletion": (reviewer) => {
    apply(reviewer, "suggested", [
      { id: "row", type: "deleteTableRow", blockId: blockId(reviewer, "Item") },
    ]);
  },
};

for (const [name, stage] of Object.entries(SUGGESTIONS)) {
  test(`accept-all leaves ${name} staged`, async () => {
    const reviewer = await open();
    const ordinary = new Set(
      apply(reviewer, "tracked-changes", [
        {
          id: "ordinary",
          type: "insertAfterBlock",
          blockId: blockId(reviewer, "Closing words."),
          text: "Ordinary words.",
        },
      ]).applied.flatMap(({ revisionIds }) => revisionIds ?? []),
    );
    const before = await savedTexts(reviewer);
    stage(reviewer);
    const staged = reviewer
      .getChanges()
      .map(({ id }) => id)
      .filter((id) => !ordinary.has(id));
    expect(staged.length).toBeGreaterThan(0);
    expect(await savedTexts(reviewer)).toEqual(before);

    reviewer.acceptAll();
    expect(reviewer.getChanges().map(({ id }) => id)).toEqual(staged);
    expect(await savedTexts(reviewer)).toEqual(before);
  });
}
