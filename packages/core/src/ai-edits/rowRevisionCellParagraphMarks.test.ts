/**
 * A tracked row insertion or deletion is one change. The canonical session
 * also marks each cell paragraph's mark with the row's revision; those marks
 * belong to the row's entry in the change list, and accepting or rejecting
 * the row resolves them with it, on either engine.
 */
import { expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const ENGINES = ["legacy", "canonical"] as const;

type Engine = (typeof ENGINES)[number];

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  operation: FolioDocumentOperation,
) =>
  expect(
    reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: [operation],
      },
      { undefinedReferences: "refuse" },
    ).skipped,
  ).toEqual([]);

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.snapshot().blocks.find((each) => each.text === text);
  if (block === undefined) throw new Error(`no block reads ${text}`);
  return block.id;
};

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ text }) => text);

const changeList = (reviewer: FolioDocxReviewer) =>
  reviewer.getChanges().map(({ type, text }) => ({ type, text }));

/** A table under "Opening words." whose second row's "months" is struck. */
const openTable = async (engine: Engine) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(fromMarkdown("Opening words.\n\nClosing words.")),
    { author: "Editor", session: engine },
  );
  apply(reviewer, "direct", {
    id: "table",
    type: "insertTable",
    blockId: blockId(reviewer, "Opening words."),
    rows: [
      ["Term", "Value"],
      ["Period", "12 months"],
    ],
  });
  apply(reviewer, "tracked-changes", {
    id: "strike",
    type: "replaceInBlock",
    blockId: blockId(reviewer, "12 months"),
    find: "months",
    replace: "",
  });
  return reviewer;
};

const deleteRow = async (engine: Engine) => {
  const reviewer = await openTable(engine);
  apply(reviewer, "tracked-changes", {
    id: "row",
    type: "deleteTableRow",
    blockId: blockId(reviewer, "Period"),
  });
  const row = reviewer.getChanges().find(({ type }) => type === "rowDeleted");
  if (row === undefined) throw new Error("the row deletion is not listed");
  return { reviewer, row };
};

test.each(ENGINES)(
  "a tracked row deletion lists as one change beside the edit inside it (%s)",
  async (engine) => {
    const { reviewer } = await deleteRow(engine);
    expect(changeList(reviewer)).toEqual([
      { type: "rowDeleted", text: "Period12 months" },
      { type: "deletion", text: "months" },
    ]);
  },
);

test.each(ENGINES)(
  "rejecting a row deletion restores the row and leaves the edit inside it (%s)",
  async (engine) => {
    const { reviewer, row } = await deleteRow(engine);
    expect(reviewer.rejectChange(row)).toBe(true);
    expect(changeList(reviewer)).toEqual([
      { type: "deletion", text: "months" },
    ]);
    expect(texts(reviewer)).toEqual([
      "Opening words.",
      "Term",
      "Value",
      "Period",
      "12 ",
      "Closing words.",
    ]);
  },
);

test.each(ENGINES)(
  "accepting a row deletion removes the row and every mark in it (%s)",
  async (engine) => {
    const { reviewer, row } = await deleteRow(engine);
    expect(reviewer.acceptChange(row)).toBe(true);
    expect(changeList(reviewer)).toEqual([]);
    expect(texts(reviewer)).toEqual([
      "Opening words.",
      "Term",
      "Value",
      "Closing words.",
    ]);
  },
);

test.each(ENGINES)(
  "a tracked row insertion lists and resolves as one change (%s)",
  async (engine) => {
    const reviewer = await openTable(engine);
    apply(reviewer, "tracked-changes", {
      id: "row",
      type: "insertTableRow",
      blockId: blockId(reviewer, "Term"),
      position: "after",
      cellTexts: ["Notice", "30 days"],
    });
    expect(changeList(reviewer)).toEqual([
      { type: "rowInserted", text: "Notice30 days" },
      { type: "deletion", text: "months" },
    ]);
    const row = reviewer.getChanges().at(0);
    if (row === undefined) throw new Error("the row insertion is not listed");
    expect(reviewer.rejectChange(row)).toBe(true);
    expect(changeList(reviewer)).toEqual([
      { type: "deletion", text: "months" },
    ]);
    expect(texts(reviewer)).toEqual([
      "Opening words.",
      "Term",
      "Value",
      "Period",
      "12 ",
      "Closing words.",
    ]);
  },
);
