/**
 * Deleting, tracked, one's own inserted paragraph break right before a table
 * whose rows are all pending insertion, on the canonical session: the
 * deletion is recorded beside the insertion on the mark and takes the
 * original words, so accepting everything removes the paragraph as the direct
 * deletion does, and rejecting everything removes it as the insertion it is.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document } from "../types/document";

const AUTHOR = "Editor";

type Mode = "direct" | "tracked-changes";

const apply = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
) =>
  reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  ).skipped;

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

const shape = (reviewer: FolioDocxReviewer) =>
  reviewer.toDocument().package.document.content.map(blockText);

/** The paragraph with its own break inserted, a pending inserted table after it. */
const open = async (table: "standing" | "none" = "standing") => {
  const document = fromMarkdown("Opening words.\n\nAlpha words.\n\nClosing.");
  const inserted = document.package.document.content[1];
  if (inserted?.type !== "paragraph") throw new Error("no paragraph");
  inserted.pPrMark = {
    kind: "ins",
    info: { id: 7, author: AUTHOR, date: "2026-01-01T00:00:00Z" },
  };
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(document),
    { author: AUTHOR, session: "canonical" },
  );
  const paragraph = reviewer.snapshot().blocks[1];
  if (paragraph === undefined) throw new Error("no paragraph");
  const insertion: FolioDocumentOperation = {
    id: "table",
    type: "insertTable",
    blockId: paragraph.id,
    rows: [["Term", "Value"]],
  };
  if (table === "standing")
    expect(apply(reviewer, "tracked-changes", [insertion])).toEqual([]);
  return { reviewer, paragraph, insertion };
};

test("accepting the tracked deletion removes the paragraph as the direct deletion does", async () => {
  const shapes: string[][] = [];
  for (const mode of ["direct", "tracked-changes"] as const) {
    const { reviewer, paragraph } = await open();
    expect(
      apply(reviewer, mode, [
        { id: "delete", type: "deleteBlock", blockId: paragraph.id },
      ]),
    ).toEqual([]);
    reviewer.acceptAll();
    shapes.push(shape(reviewer));
  }
  expect(shapes).toEqual([
    ["Opening words.", "[table]", "Closing."],
    ["Opening words.", "[table]", "Closing."],
  ]);
});

test("the tracked deletion sits beside the insertion on the mark, and rejecting removes the insertion", async () => {
  const { reviewer, paragraph } = await open();
  expect(
    apply(reviewer, "tracked-changes", [
      { id: "delete", type: "deleteBlock", blockId: paragraph.id },
    ]),
  ).toEqual([]);
  const target = reviewer
    .toDocument()
    .package.document.content.find(
      (block) =>
        block.type === "paragraph" &&
        block.paraId?.toUpperCase() === paragraph.id.toUpperCase(),
    );
  expect(target?.type === "paragraph" && target.pPrMark?.kind).toBe("ins");
  expect(
    target?.type === "paragraph" && target.pPrMark?.deletion?.info.author,
  ).toBe(AUTHOR);
  reviewer.rejectAll();
  expect(shape(reviewer)).toEqual(["Opening words.", "Alpha words.Closing."]);
});

test("a batch that inserts the table and deletes the paragraph reads as one at a time, in either order", async () => {
  const { reviewer: sequential, paragraph } = await open();
  expect(
    apply(sequential, "tracked-changes", [
      { id: "delete", type: "deleteBlock", blockId: paragraph.id },
    ]),
  ).toEqual([]);
  sequential.acceptAll();
  expect(shape(sequential)).toEqual(["Opening words.", "[table]", "Closing."]);
  for (const order of ["table first", "deletion first"] as const) {
    const { reviewer, insertion } = await open("none");
    const deletion: FolioDocumentOperation = {
      id: "delete",
      type: "deleteBlock",
      blockId: paragraph.id,
    };
    expect(
      apply(
        reviewer,
        "tracked-changes",
        order === "table first" ? [insertion, deletion] : [deletion, insertion],
      ),
    ).toEqual([]);
    reviewer.acceptAll();
    expect(shape(reviewer)).toEqual(shape(sequential));
  }
});
