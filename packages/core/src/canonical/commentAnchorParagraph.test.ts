import {
  afterAll,
  beforeAll,
  expect,
  setSystemTime,
  spyOn,
  test,
} from "bun:test";

import * as nativeOps from "@stll/docx-core/ops";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import { paragraph as buildParagraph } from "../docx/server/build";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import * as anchors from "./commentAnchorParagraph";

const DATE = new Date("2026-01-01T10:00:00.000Z");
beforeAll(() => setSystemTime(DATE));
afterAll(() => setSystemTime());

const box = (...content: Paragraph[]) =>
  ({
    type: "shape",
    shape: {
      type: "shape",
      shapeType: "textBox",
      size: { width: 2_000_000, height: 800_000 },
      textBody: { content },
    },
  }) as const;
const named = (paraId: string, value: string) => ({
  ...buildParagraph(value),
  paraId,
});
const unnamed = buildParagraph("Unnamed");
const emptyId = named("", "Explicit empty ID");
const firstOwn = named("1000000a", "First own");
const secondOwn = named("1000000a", "Second own");
const firstBox = named("1000000A", "First exact-case box");
const laterBox = named("1000000A", "Later exact-case box");
const nestedOnly = named("1000000B", "Nested first");
const laterOnly = named("1000000B", "Later outer box");
const lowercaseOnly = named("1000000d", "Lowercase only");

test("an anchor uses the first exact ID in the story and nested-box order", () => {
  const host = named("10000001", "Host");
  const nestedHost = {
    ...named("10000002", "Nested host"),
    content: [{ type: "run", content: [box(nestedOnly)] }],
  } satisfies Paragraph;
  host.content = [
    {
      type: "run",
      content: [box(firstBox, nestedHost), box(laterBox, laterOnly)],
    },
  ];
  const document = {
    package: {
      document: {
        content: [unnamed, emptyId, firstOwn, secondOwn, lowercaseOnly, host],
      },
    },
  } satisfies Document;
  const find = (blockId: string | undefined) =>
    anchors.findCommentAnchorParagraph({
      document,
      story: nativeOps.OP_STORIES.MAIN,
      blockId,
    });
  const walk = spyOn(nativeOps, "paragraphsThroughTextBoxes");
  try {
    expect(find("1000000a")).toBe(firstOwn);
    expect(find("1000000B")).toBe(nestedOnly);
    expect(find("1000000C")).toBeUndefined();
    expect(walk).toHaveBeenCalledTimes(0);

    expect(find("1000000A")).toBe(firstBox);
    expect(walk).toHaveBeenCalledTimes(1);
    expect(find("1000000D")).toBeUndefined();
    expect(walk).toHaveBeenCalledTimes(2);
    expect(find("")).toBe(emptyId);
    expect(find(undefined)).toBe(unnamed);
    expect(walk).toHaveBeenCalledTimes(4);
  } finally {
    walk.mockRestore();
  }
});

type Mode = "direct" | "tracked-changes";
type EditFixture = { width: number; mode: Mode };

const fixture = async ({ width, mode }: EditFixture) => {
  const idAt = (index: number) => nativeOps.formatParaId(0x10000000 + index);
  const document = createEmptyDocument();
  document.package.document.content = Array.from(
    { length: width },
    (_, index) => named(idAt(index), `Clause ${index} stays.`),
  );
  const source = await createDocx(document);
  const operations: FolioDocumentOperation[] = [
    width - 5,
    width - 4,
    width - 3,
  ].map((index) => ({
    id: `edit-${index}`,
    type: "replaceInBlock",
    blockId: idAt(index),
    find: "Clause",
    replace: "Item",
    comment: { text: `Check clause ${index}.` },
  }));
  if (mode === "direct")
    operations.push({
      id: "remove-anchor",
      type: "deleteBlock",
      blockId: idAt(width - 2),
      comment: { text: "Cannot anchor." },
    });
  else
    operations.push({
      id: "missing-quote",
      type: "replaceInBlock",
      blockId: idAt(width - 2),
      find: "missing quote",
      replace: "Different",
      comment: { text: "Cannot anchor." },
    });
  return { source, operations };
};

const applyEdits = async (
  source: ArrayBuffer,
  options: { mode: Mode; operations: FolioDocumentOperation[] },
) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source, {
    session: "canonical",
  });
  const original = reviewer.toDocument();
  const { undoHandle, ...receipt } = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      ...options,
    },
    { undefinedReferences: "refuse" },
  );
  const bytes = new Uint8Array(await reviewer.toBuffer());
  const reopened = await FolioDocxReviewer.fromBuffer(bytes, {
    session: "canonical",
  });
  const contents = (current: FolioDocxReviewer) =>
    current
      .getComments()
      .map(({ id, text, anchoredText }) => ({ id, text, anchoredText }));
  const live = contents(reviewer);
  // The handle's UUID is specific to this reviewer instance. Verify that
  // each handle actually restores its input.
  if (undoHandle === null) throw new Error("Missing undo handle");
  expect(reviewer.undoDocumentOperations(undoHandle)).toEqual({
    status: "undone",
    undoHandle,
  });
  expect(reviewer.toDocument()).toEqual(original);
  return { receipt, bytes, anchors: live, reopened: contents(reopened) };
};

const findWithFullWalk: typeof anchors.findCommentAnchorParagraph = ({
  document,
  story,
  blockId,
}) =>
  nativeOps
    .paragraphsThroughTextBoxes(document, story)
    .find(({ paragraph }) => paragraph.paraId === blockId)?.paragraph;

for (const { width, mode } of [
  { width: 16, mode: "direct" },
  { width: 64, mode: "direct" },
  { width: 16, mode: "tracked-changes" },
] as const) {
  test(`public ${mode} comments at width ${width} avoid full anchor walks`, async () => {
    const { source, operations } = await fixture({ width, mode });
    const fullWalk = nativeOps.paragraphsThroughTextBoxes;
    let enumerated = 0;
    const census = spyOn(
      nativeOps,
      "paragraphsThroughTextBoxes",
    ).mockImplementation((document, story) => {
      const result = fullWalk(document, story);
      enumerated += result.length;
      return result;
    });
    try {
      const actual = await applyEdits(source, { mode, operations });
      expect(actual.receipt.applied).toHaveLength(3);
      expect(actual.receipt.skipped).toHaveLength(1);
      if (mode === "direct")
        expect(actual.receipt.skipped.at(0)).toMatchObject({
          id: "remove-anchor",
          reason: "unsupportedBlock",
        });
      expect(actual.anchors).toHaveLength(3);
      expect(actual.reopened).toHaveLength(3);
      const indexedCalls = census.mock.calls.length;
      const indexedParagraphs = enumerated;
      census.mockClear();
      enumerated = 0;

      const reference = spyOn(
        anchors,
        "findCommentAnchorParagraph",
      ).mockImplementation(findWithFullWalk);
      try {
        expect(await applyEdits(source, { mode, operations })).toEqual(actual);
        const anchorWalks = reference.mock.calls.length;
        expect(anchorWalks).toBe(6);
        expect(census.mock.calls.length).toBe(indexedCalls + anchorWalks);
        expect(enumerated).toBe(indexedParagraphs + width * anchorWalks);
      } finally {
        reference.mockRestore();
      }
    } finally {
      census.mockRestore();
    }
  });
}

test("public edits retain a comment on a paragraph in a text box", async () => {
  const model = createEmptyDocument();
  const target = named("20000001", "Box clause.");
  const host = named("10000001", "Host.");
  host.content.push({ type: "run", content: [box(target)] });
  model.package.document.content = [host, named("10000002", "Body.")];
  const source = await createDocx(model);
  const operations = [
    {
      id: "box",
      type: "replaceInBlock",
      blockId: "20000001",
      find: "clause",
      replace: "term",
      comment: { text: "In the box." },
    },
  ] satisfies FolioDocumentOperation[];
  const actual = await applyEdits(source, { mode: "direct", operations });
  expect(actual.receipt.skipped).toEqual([]);
  expect(actual.receipt.applied).toHaveLength(1);
  expect(actual.anchors).toHaveLength(1);
  expect(actual.anchors.at(0)).toMatchObject({
    text: "In the box.",
    anchoredText: "term",
  });
  expect(actual.reopened).toEqual(actual.anchors);
  const reference = spyOn(
    anchors,
    "findCommentAnchorParagraph",
  ).mockImplementation(findWithFullWalk);
  try {
    expect(await applyEdits(source, { mode: "direct", operations })).toEqual(
      actual,
    );
    expect(reference).toHaveBeenCalledTimes(2);
  } finally {
    reference.mockRestore();
  }
});
