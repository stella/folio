import { expect, test } from "bun:test";

import { paragraphNumberingReference } from "@stll/docx-core/model";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  storyParagraphs,
} from "@stll/docx-core/ops";

import type { Document, Paragraph } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { normalizeListModelRendering } from "../listModelNormalization";
import { changeListFormat } from "./changeListFormat";

const item = (paraId: string, numId: number, ilvl: number): Paragraph => ({
  type: "paragraph",
  paraId,
  formatting: { numPr: paragraphNumberingReference({ numId, ilvl }) },
  content: [{ type: "run", content: [{ type: "text", text: paraId }] }],
});

const fixture = () => {
  const document = createEmptyDocument();
  document.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 0,
        multiLevelType: "multilevel",
        levels: [
          { ilvl: 0, numFmt: "decimal", lvlText: "%1.", start: 1 },
          { ilvl: 1, numFmt: "decimal", lvlText: "%2.", start: 1 },
        ],
      },
    ],
    nums: [
      { numId: 1, abstractNumId: 0 },
      { numId: 2, abstractNumId: 0 },
    ],
  };
  document.package.document.content = [
    item("10000001", 1, 0),
    item("10000002", 1, 0),
    item("10000003", 1, 1),
    item("10000004", 2, 0),
  ];
  return document;
};

const bindings = (document: Document) =>
  storyParagraphs(document.package.document).map(
    ({ paragraph }) => paragraph.formatting?.numPr,
  );

const labels = (document: Document) =>
  storyParagraphs(
    normalizeListModelRendering(document, { inverse: "omit" }).document.package
      .document,
  ).map(({ paragraph }) => paragraph.listRendering?.marker);

test("accepts changed list markers and rejects back to the original bindings", () => {
  const document = fixture();
  const originalBindings = bindings(document);
  const applied = changeListFormat(document, {
    sourceNumId: 1,
    level: 0,
    numFmt: "upperRoman",
    lvlText: "%1)",
    start: 4,
    mode: {
      type: "suggesting",
      revision: { id: 10, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
      newIds: { revision: [11, 12] },
    },
  }).unwrap();
  expect(applied.revisions).toHaveLength(3);
  expect(bindings(document)).toEqual(originalBindings);
  const resolve = (decision: "accept" | "reject") =>
    applyDocumentOps(applied.document, [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: applied.revisions,
        decision,
      },
    ]).unwrap().document;
  const rejected = resolve("reject");
  const accepted = resolve("accept");
  expect(bindings(rejected)).toEqual(originalBindings);
  expect(labels(rejected)).toEqual(labels(document));
  expect(labels(accepted)).toEqual(["IV)", "V)", "1.", "1."]);
  const [first, second, nested, unrelated] = bindings(accepted);
  expect(first).toEqual(second);
  expect(first?.kind).toBe("reference");
  if (first?.kind !== "reference") throw new Error("Expected list membership");
  expect(nested).toEqual(
    paragraphNumberingReference({ numId: first.numId, ilvl: 1 }),
  );
  expect(unrelated).toEqual(originalBindings.at(-1));
});

test("restarts only explicitly selected list members", () => {
  const document = fixture();
  const options = {
    sourceNumId: 1,
    level: 0,
    numFmt: "lowerLetter",
    start: 3,
    mode: { type: "editing" },
  } as const;
  const applied = changeListFormat(document, {
    ...options,
    paragraphIds: ["10000002"],
  }).unwrap();
  expect(labels(applied.document)).toEqual(["1.", "c.", "1.", "1."]);
  const mixed = changeListFormat(document, {
    ...options,
    paragraphIds: ["10000002", "10000004"],
  });
  if (mixed.isOk()) throw new Error("Expected an invalid selection refusal");
  expect(mixed.error.reason).toBe("blockNotFound");
  expect(labels(document)).toEqual(["1.", "2.", "1.", "1."]);
});

test("changes one story without rebinding other stories on the same instance", () => {
  const document = fixture();
  const story = { kind: "header", rId: "rId7" } as const;
  document.package.headers = new Map([
    [
      story.rId,
      {
        type: "header",
        hdrFtrType: "default",
        content: [item("10000005", 1, 0)],
      },
    ],
  ]);
  const changed = changeListFormat(document, {
    story,
    sourceNumId: 1,
    level: 0,
    numFmt: "upperLetter",
    start: 5,
    mode: { type: "editing" },
  }).unwrap().document;
  expect(bindings(changed)).toEqual(bindings(document));
  const normalized = normalizeListModelRendering(changed, {
    inverse: "omit",
  }).document;
  const header = normalized.package.headers?.get(story.rId);
  if (header === undefined) throw new Error("Expected the edited header");
  expect(
    storyParagraphs(header).map(
      ({ paragraph }) => paragraph.listRendering?.marker,
    ),
  ).toEqual(["E."]);
});
