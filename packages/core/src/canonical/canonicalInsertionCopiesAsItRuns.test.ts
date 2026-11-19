/**
 * A paragraph inserted beside another copies that paragraph's stated
 * properties as they stand when the insertion runs in the batch's execution
 * order: equal coordinates run in reverse input order, an insertion before a
 * paragraph shares its start with an edit of the paragraph's properties, and a
 * merge of the paragraph sits at its end. So an edit the order runs first is
 * in the copy, and one it runs after is not.
 */

import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import type {
  FolioDocumentOperation,
  FolioDocumentOperationMode,
} from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const INSERTED = "Inserted clause.";

const open = async () => {
  const input = (
    await prepareCanonicalDocxInput(
      await createDocx(
        fromMarkdown(
          "# Terms\n\n1. Numbered clause one\n\nUnnumbered body text.\n\nClosing.",
        ),
      ),
    )
  )
    .unwrap()
    .slice().buffer;
  const reviewer = await FolioDocxReviewer.fromBuffer(input, {
    session: "canonical",
  });
  const [heading, numbered, body] = reviewer.snapshot().blocks;
  if (
    heading?.text !== "Terms" ||
    numbered?.listReference === undefined ||
    body?.text !== "Unnumbered body text."
  )
    throw new Error("Missing fixture");
  return { reviewer, heading, numbered, body };
};

const insertion = (
  type: "insertBeforeBlock" | "insertAfterBlock",
  blockId: string,
): FolioDocumentOperation => ({
  id: "insert",
  type,
  blockId,
  text: INSERTED,
});

const inserted = (reviewer: FolioDocxReviewer) => {
  const block = reviewer
    .snapshot()
    .blocks.find(({ text }) => text === INSERTED);
  if (block === undefined) throw new Error("The insertion is missing");
  return block;
};

const MODES = ["direct", "tracked-changes"] as const;

const apply = (
  reviewer: FolioDocxReviewer,
  operations: FolioDocumentOperation[],
  mode: FolioDocumentOperationMode,
) => {
  const result = reviewer.applyDocumentOperations(
    { version: 1, mode, atomic: true, operations },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(operations.length);
};

describe("an inserted paragraph copies its neighbour as the batch order leaves it", () => {
  for (const mode of MODES) {
    test(`before a paragraph the batch then numbers nothing (${mode})`, async () => {
      const { reviewer, numbered } = await open();
      apply(
        reviewer,
        [
          insertion("insertBeforeBlock", numbered.id),
          {
            id: "unnumber",
            type: "setBlockParagraphProperties",
            blockId: numbered.id,
            properties: { numbering: { kind: "none" } },
          },
        ],
        mode,
      );
      expect(inserted(reviewer).listReference).toBeUndefined();
    });

    test(`before a paragraph the batch numbered nothing first, in input order (${mode})`, async () => {
      const { reviewer, numbered } = await open();
      apply(
        reviewer,
        [
          {
            id: "unnumber",
            type: "setBlockParagraphProperties",
            blockId: numbered.id,
            properties: { numbering: { kind: "none" } },
          },
          insertion("insertBeforeBlock", numbered.id),
        ],
        mode,
      );
      expect(inserted(reviewer).listReference).toEqual(numbered.listReference);
    });

    test(`after a paragraph the batch then numbers nothing (${mode})`, async () => {
      const { reviewer, numbered } = await open();
      apply(
        reviewer,
        [
          insertion("insertAfterBlock", numbered.id),
          {
            id: "unnumber",
            type: "setBlockParagraphProperties",
            blockId: numbered.id,
            properties: { numbering: { kind: "none" } },
          },
        ],
        mode,
      );
      expect(inserted(reviewer).listReference).toEqual(numbered.listReference);
    });

    test(`before a paragraph the batch then clears the style of (${mode})`, async () => {
      const { reviewer, heading } = await open();
      apply(
        reviewer,
        [
          insertion("insertBeforeBlock", heading.id),
          {
            id: "unstyle",
            type: "setBlockParagraphProperties",
            blockId: heading.id,
            properties: { styleId: null },
          },
        ],
        mode,
      );
      expect(inserted(reviewer).styleId).toBeUndefined();
    });

    test(`before a paragraph the batch then starts a list with (${mode})`, async () => {
      const { reviewer, body } = await open();
      apply(
        reviewer,
        [
          insertion("insertBeforeBlock", body.id),
          {
            id: "list",
            type: "setBlockParagraphProperties",
            blockId: body.id,
            properties: { numbering: { kind: "newList", format: "bullet" } },
          },
        ],
        mode,
      );
      const listed = reviewer
        .snapshot()
        .blocks.find(({ id }) => id === body.id)?.listReference;
      expect(listed).toBeDefined();
      expect(inserted(reviewer).listReference).toEqual(listed);
    });
  }

  // A tracked merge leaves the paragraph after it, whose properties the
  // merge sets: the paragraph inserted before the joined one copies them.
  test(`before a paragraph the batch merges with new properties (tracked-changes)`, async () => {
    const { reviewer, heading, numbered } = await open();
    const numPr = numbered.statedNumbering;
    if (numPr?.kind !== "reference") throw new Error("Missing numbering");
    apply(
      reviewer,
      [
        {
          id: "merge",
          type: "mergeBlockWithNext",
          blockId: heading.id,
          separator: " ",
          mergedParagraphProperties: { numbering: numPr },
        },
        insertion("insertBeforeBlock", heading.id),
      ],
      "tracked-changes",
    );
    expect(inserted(reviewer).listReference).toEqual(numbered.listReference);
  });
});
