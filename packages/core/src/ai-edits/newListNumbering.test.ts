/**
 * `numbering: { start: "new", kind }` (issue #1092): an operation can start a
 * list in a package with no numbering part, or a second list beside an
 * existing one, and the saved package defines every instance it references.
 */

import { describe, expect, test } from "bun:test";

import { paragraphNumberingReferenceId } from "@stll/docx-core/model";

import {
  parseFolioDocumentOperationBatch,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const reviewerOf = async (markdown: string): Promise<FolioDocxReviewer> => {
  const { docx } = await ensureParaIds(await createDocx(fromMarkdown(markdown)));
  return FolioDocxReviewer.fromBuffer(docx, { author: "Agent" });
};

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) {
    throw new Error(`no block reads "${text}"`);
  }
  return block.id;
};

/** Apply, save, check every reference is defined, and read the labels back. */
const savedLabels = async (
  reviewer: FolioDocxReviewer,
  operations: FolioDocumentOperation[],
): Promise<string[]> => {
  const result = reviewer.applyDocumentOperations({
    version: 1,
    mode: "tracked-changes",
    operations,
  });
  expect(result.status).toBe("committed");
  expect(result.issues).toEqual([]);

  const saved = await reviewer.toBuffer();
  const reparsed = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
  const defined = new Set((reparsed.package.numbering?.nums ?? []).map(({ numId }) => numId));
  for (const block of reparsed.package.document.content) {
    if (block.type !== "paragraph") continue;
    const numId = paragraphNumberingReferenceId(block.formatting?.numPr);
    if (numId !== undefined) expect(defined.has(numId)).toBe(true);
  }
  return (await FolioDocxReviewer.fromBuffer(saved))
    .getContent()
    .map(({ text, displayLabel }) => `${displayLabel ?? ""} ${text}`.trim());
};

describe("an operation that starts a new list", () => {
  test("defines the list in a package without a numbering part", async () => {
    const reviewer = await reviewerOf("Intro.\n\nTail.");
    const labels = await savedLabels(reviewer, [
      {
        id: "list",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Intro."),
        text: "First\nSecond",
        formattingScope: "allParagraphs",
        numbering: { start: "new", kind: "numbered" },
      },
    ]);

    expect(labels).toEqual(["Intro.", "1. First", "2. Second", "Tail."]);
  });

  test("starts a second list at one without renumbering the first", async () => {
    const reviewer = await reviewerOf("1. Alpha\n2. Beta\n\nTail.");
    const labels = await savedLabels(reviewer, [
      {
        id: "list",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Tail."),
        text: "Gamma\nDelta",
        formattingScope: "allParagraphs",
        numbering: { start: "new", kind: "numbered" },
      },
    ]);

    expect(labels).toEqual(["1. Alpha", "2. Beta", "Tail.", "1. Gamma", "2. Delta"]);
  });

  test("starts a bullet list", async () => {
    const reviewer = await reviewerOf("Intro.");
    const labels = await savedLabels(reviewer, [
      {
        id: "list",
        type: "insertAfterBlock",
        blockId: blockId(reviewer, "Intro."),
        text: "First\nSecond",
        formattingScope: "allParagraphs",
        numbering: { start: "new", kind: "bullet" },
      },
    ]);

    expect(labels).toEqual(["Intro.", "• First", "• Second"]);
  });

  test("numbers existing paragraphs, one list per operation", async () => {
    const reviewer = await reviewerOf("One\n\nProse.\n\nTwo");
    const labels = await savedLabels(reviewer, [
      {
        id: "first",
        type: "setBlockParagraphProperties",
        blockId: blockId(reviewer, "One"),
        properties: { numbering: { start: "new", kind: "numbered" } },
      },
      {
        id: "second",
        type: "setBlockParagraphProperties",
        blockId: blockId(reviewer, "Two"),
        properties: { numbering: { start: "new", kind: "numbered", level: 0 } },
      },
    ]);

    expect(labels).toEqual(["1. One", "Prose.", "1. Two"]);
  });

  test.each([
    [{ start: "new", kind: "roman" }, "$.operations[0].properties.numbering.kind"],
    [{ start: "new", kind: "numbered", level: 9 }, "$.operations[0].properties.numbering.level"],
    [{ start: "new", kind: "bullet", numId: 3 }, "$.operations[0].properties.numbering.numId"],
    [{ start: "old", kind: "bullet" }, "$.operations[0].properties.numbering.start"],
  ])("rejects %j", (numbering, path) => {
    expect(() =>
      parseFolioDocumentOperationBatch({
        version: 1,
        operations: [
          { id: "1", type: "setBlockParagraphProperties", blockId: "a", properties: { numbering } },
        ],
      }),
    ).toThrow(expect.objectContaining({ _tag: "InvalidFolioDocumentOperationBatchError", path }));
  });
});
