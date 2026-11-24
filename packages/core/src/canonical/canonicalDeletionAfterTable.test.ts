/**
 * Deleting a paragraph a table stands right before, or right after, on the
 * canonical session: the body's last paragraph after a table keeps its mark
 * and loses its words, whether the table stood there already or the same
 * batch puts it there; a paragraph before a table whose own mark is a pending
 * insertion keeps that mark and loses its words; so does a paragraph between
 * two tables, which keeps them apart, unless the same batch deletes the table
 * after it (both sessions).
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { buildBodySequenceDocx } from "../compare/__fixtures__/body-sequence";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document } from "../types/document";

type Mode = "direct" | "tracked-changes";

const batch = (mode: Mode, operations: FolioDocumentOperation[]) => ({
  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  mode,
  operations,
});

const open = async (document: Document) =>
  FolioDocxReviewer.fromBuffer(await createDocx(document), {
    author: "Editor",
    session: "canonical",
  });

const ROWS = [["Term", "Value"]];

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

/** Block texts, a table as one entry. */
const shape = (reviewer: FolioDocxReviewer) =>
  reviewer.toDocument().package.document.content.map(blockText);

test.each(["direct", "tracked-changes"] as const)(
  "the last paragraph after a table loses its words and keeps its mark (%s)",
  async (mode) => {
    const reviewer = await open(
      fromMarkdown("Opening words.\n\nMiddle words.\n\nClosing words."),
    );
    const [, middle, closing] = reviewer.snapshot().blocks;
    if (middle === undefined || closing === undefined)
      throw new Error("the fixture's blocks are missing");
    expect(
      reviewer.applyDocumentOperations(
        batch("direct", [
          { id: "table", type: "insertTable", blockId: middle.id, rows: ROWS },
        ]),
        { undefinedReferences: "refuse" },
      ).skipped,
    ).toEqual([]);
    const deleted = reviewer.applyDocumentOperations(
      batch(mode, [{ id: "delete", type: "deleteBlock", blockId: closing.id }]),
      { undefinedReferences: "refuse" },
    );
    expect(deleted.skipped).toEqual([]);
    reviewer.acceptAll();
    expect(shape(reviewer)).toEqual([
      "Opening words.",
      "Middle words.",
      "[table]",
      "",
    ]);
  },
);

test.each(["direct", "tracked-changes"] as const)(
  "a table the batch puts before the deleted last paragraph keeps the paragraph before it whole (%s)",
  async (mode) => {
    const reviewer = await open(
      fromMarkdown("Opening words.\n\nMiddle words.\n\nClosing words."),
    );
    const [, middle, closing] = reviewer.snapshot().blocks;
    if (middle === undefined || closing === undefined)
      throw new Error("the fixture's blocks are missing");
    const result = reviewer.applyDocumentOperations(
      batch(mode, [
        { id: "table", type: "insertTable", blockId: middle.id, rows: ROWS },
        { id: "delete", type: "deleteBlock", blockId: closing.id },
      ]),
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    reviewer.acceptAll();
    expect(shape(reviewer)).toEqual([
      "Opening words.",
      "Middle words.",
      "[table]",
      "",
    ]);
  },
);

test("a paragraph before a table whose mark is a pending insertion keeps the mark", async () => {
  const document = fromMarkdown(
    "Opening words.\n\nInserted break.\n\nClosing.",
  );
  const inserted = document.package.document.content[1];
  if (inserted?.type !== "paragraph") throw new Error("no paragraph");
  inserted.pPrMark = {
    kind: "ins",
    info: { id: 7, author: "Earlier Reviewer", date: "2026-01-01T00:00:00Z" },
  };
  const reviewer = await open(document);
  const [, paragraph] = reviewer.snapshot().blocks;
  if (paragraph === undefined) throw new Error("no paragraph");
  expect(
    reviewer.applyDocumentOperations(
      batch("direct", [
        { id: "table", type: "insertTable", blockId: paragraph.id, rows: ROWS },
      ]),
      { undefinedReferences: "refuse" },
    ).skipped,
  ).toEqual([]);
  const deleted = reviewer.applyDocumentOperations(
    batch("tracked-changes", [
      { id: "delete", type: "deleteBlock", blockId: paragraph.id },
    ]),
    { undefinedReferences: "refuse" },
  );
  expect(deleted.skipped).toEqual([]);
  const target = reviewer
    .toDocument()
    .package.document.content.find(
      (block) =>
        block.type === "paragraph" &&
        block.paraId?.toUpperCase() === paragraph.id.toUpperCase(),
    );
  expect(target?.type === "paragraph" && target.pPrMark?.kind).toBe("ins");
  reviewer.rejectAll();
  expect(shape(reviewer)).toEqual(["Opening words.", "[table]", "Closing."]);
});

test.each([
  ["direct", "standing"],
  ["tracked-changes", "standing"],
  ["direct", "same batch"],
  ["tracked-changes", "same batch"],
] as const)(
  "a paragraph between two tables keeps its mark and loses its words (%s, table before it %s)",
  async (mode, before) => {
    const reviewer = await open(
      fromMarkdown("Opening words.\n\nBetween words.\n\nClosing."),
    );
    const [opening, between] = reviewer.snapshot().blocks;
    if (opening === undefined || between === undefined)
      throw new Error("the fixture's blocks are missing");
    const after: FolioDocumentOperation = {
      id: "after",
      type: "insertTable",
      blockId: between.id,
      rows: ROWS,
    };
    const preceding: FolioDocumentOperation = {
      id: "before",
      type: "insertTable",
      blockId: opening.id,
      rows: ROWS,
    };
    const standing = before === "standing" ? [after, preceding] : [after];
    expect(
      reviewer.applyDocumentOperations(batch("direct", standing), {
        undefinedReferences: "refuse",
      }).skipped,
    ).toEqual([]);
    const deleted = reviewer.applyDocumentOperations(
      batch(mode, [
        ...(before === "standing" ? [] : [preceding]),
        { id: "delete", type: "deleteBlock", blockId: between.id },
      ]),
      { undefinedReferences: "refuse" },
    );
    expect(deleted.skipped).toEqual([]);
    const pending = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { author: "Editor", session: "canonical" },
    );
    pending.rejectAll();
    let rejected = ["Opening words.", "[table]", "", "[table]", "Closing."];
    if (mode === "tracked-changes")
      rejected =
        before === "standing"
          ? [
              "Opening words.",
              "[table]",
              "Between words.",
              "[table]",
              "Closing.",
            ]
          : ["Opening words.", "Between words.", "[table]", "Closing."];
    expect(shape(pending)).toEqual(rejected);
    reviewer.acceptAll();
    expect(shape(reviewer)).toEqual([
      "Opening words.",
      "[table]",
      "",
      "[table]",
      "Closing.",
    ]);
  },
);

test.each(
  (["canonical"] as const).flatMap((session) =>
    (["direct", "tracked-changes"] as const).flatMap((mode) =>
      (["Between words.", ""] as const).map(
        (text) => [session, mode, text] as const,
      ),
    ),
  ),
)(
  "a paragraph between two tables goes whole when the batch deletes the table after it (%s, %s, text %p)",
  async (session, mode, text) => {
    const source = await buildBodySequenceDocx([
      { kind: "paragraph", text: "Opening words." },
      { kind: "table", rows: [["Before row"]] },
      { kind: "paragraph", text },
      { kind: "table", rows: [["After row"]] },
      { kind: "paragraph", text: "Closing." },
    ]);
    const reviewer = await FolioDocxReviewer.fromBuffer(source, {
      author: "Editor",
      session,
    });
    const blocks = reviewer.snapshot().blocks;
    const between = blocks.find((block) => !block.table && block.text === text);
    const table = blocks.find((block) => block.text === "After row");
    if (between === undefined || table === undefined)
      throw new Error("the fixture's blocks are missing");
    const original = shape(reviewer);
    const originalBody = reviewer.toDocument().package.document.content;
    const kept = ["Opening words.", "[table]", "Closing."];
    const deleted = reviewer.applyDocumentOperations(
      batch(mode, [
        { id: "delete", type: "deleteBlock", blockId: between.id },
        { id: "table", type: "deleteTable", blockId: table.id },
      ]),
      { undefinedReferences: "refuse" },
    );
    expect(deleted.skipped).toEqual([]);
    const saved = await reviewer.toBuffer();
    if (mode === "direct") expect(shape(reviewer)).toEqual(kept);
    else {
      const rejected = await FolioDocxReviewer.fromBuffer(saved, {
        author: "Editor",
        session,
      });
      rejected.rejectAll();
      expect(shape(rejected)).toEqual(original);
      const accepted = await FolioDocxReviewer.fromBuffer(saved, {
        author: "Editor",
        session,
      });
      accepted.acceptAll();
      expect(shape(accepted)).toEqual(kept);
    }
    if (deleted.undoHandle === undefined)
      throw new Error("the batch has no undo handle");
    expect(reviewer.undoDocumentOperations(deleted.undoHandle).status).toBe(
      "undone",
    );
    expect(reviewer.toDocument().package.document.content).toEqual(
      originalBody,
    );
  },
);
