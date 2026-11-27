/**
 * The persisted form of operations in the current schema version.
 *
 * Journaled operations are read back years after they were written, so their
 * JSON is pinned: one envelope per operation kind and inverse specimen,
 * built from fixed documents. Producer decomposition may change while each
 * existing operation keeps its meaning; preserve its historical specimen
 * and prove both it and the current inverse against the actual producer.
 * A change to an operation's fields, or to a model
 * record an operation embeds, shows up here as a diff of the fixture; it
 * needs a new schema version and a migration, not an updated released fixture.
 * The fixture for an unreleased schema records that cutover's final contract;
 * schema 9 adds clipboard package resources over schema 8's semantic table edits,
 * and schema 10 adds comment operations.
 *
 * `bun packages/docx-core/src/ops/__tests__/wireFixtures.ts` generates the current fixture.
 */

import { panic } from "better-result";
import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOpEnvelope } from "../apply";
import { paragraphLogicalText } from "../offsets";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpEnvelope,
  toOpEnvelope,
} from "../types";
import {
  envelopes,
  lifecycleWireCases,
  tableBoundaryWireCases,
  wireFixturePath,
} from "./wireFixtures";

/** Paragraph reviews have one OOXML record; run-property review arrays have a separate contract. */
const expectParagraphReviewCardinality = (value: unknown): void => {
  if (Array.isArray(value)) {
    for (const entry of value) expectParagraphReviewCardinality(entry);
    return;
  }
  if (value === null || typeof value !== "object") return;
  if (
    "type" in value &&
    value.type === "paragraph" &&
    "propertyChanges" in value
  ) {
    expect(Array.isArray(value.propertyChanges)).toBe(true);
    if (Array.isArray(value.propertyChanges))
      expect(value.propertyChanges.length).toBeLessThanOrEqual(1);
  }
  for (const entry of Object.values(value))
    expectParagraphReviewCardinality(entry);
};

test("the wire cardinality guard rejects two property reviews on an identified paragraph", () => {
  const valid = {
    type: "paragraph",
    paraId: "00000001",
    content: [],
    propertyChanges: [
      {
        type: "paragraphPropertyChange",
        info: { id: 7, author: "Original reviewer" },
      },
    ],
  } satisfies Paragraph;
  expect(() => expectParagraphReviewCardinality(valid)).not.toThrow();
  const invalid = {
    ...valid,
    propertyChanges: [
      ...valid.propertyChanges,
      {
        type: "paragraphPropertyChange",
        info: { id: 8, author: "Latest reviewer" },
      },
    ],
  } satisfies Paragraph;
  expect(() => expectParagraphReviewCardinality(invalid)).toThrow();
});

test("every operation kind and its inverse keep their persisted JSON", async () => {
  const written = JSON.stringify(envelopes(), null, 2);
  if (process.env["UPDATE_OP_WIRE_FIXTURE"] === "1") {
    await Bun.write(wireFixturePath, `${written}\n`);
  }
  // The JSON as persisted, compared as data: the fixture's layout is the formatter's.
  const pinned: unknown = await Bun.file(wireFixturePath).json();
  expectParagraphReviewCardinality(pinned);
  const generated = envelopes();
  expectParagraphReviewCardinality(generated);
  for (const { op } of generated) {
    if (op.type !== DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW) continue;
    expect(op.expected.propertyChanges?.length ?? 0).toBeLessThanOrEqual(1);
    expect(op.review.propertyChanges?.length ?? 0).toBeLessThanOrEqual(1);
  }
  expect(JSON.parse(written)).toStrictEqual(pinned);
  const kinds = new Set(envelopes().map(({ op }) => op.type));
  expect([...kinds].toSorted()).toEqual(
    Object.values(DOCUMENT_OP_TYPES).toSorted(),
  );
  expect(
    envelopes().every(({ schema }) => schema === DOCUMENT_OP_SCHEMA_VERSION),
  ).toBe(true);
});

test("the pinned note restore and current producer inverse both decode, undo and redo exactly", () => {
  const cases = lifecycleWireCases().filter(
    ({ op }) => op.type === DOCUMENT_OP_TYPES.REMOVE_NOTE,
  );
  expect(cases).toHaveLength(1);
  for (const { before, edit, wireInverse } of cases) {
    expect(edit.inverse.map(({ type }) => type)).not.toEqual(
      wireInverse.map(({ type }) => type),
    );
    for (const inverse of [wireInverse, edit.inverse]) {
      const wire: DocumentOpEnvelope[] = JSON.parse(
        JSON.stringify(inverse.map(toOpEnvelope)),
      );
      let document = edit.document;
      const redo: (readonly DocumentOp[])[] = [];
      for (const envelope of wire) {
        const applied = applyDocumentOpEnvelope(document, envelope).unwrap();
        document = applied.document;
        redo.unshift(applied.inverse);
      }
      expect(document).toStrictEqual(before);
      const wireRedo: DocumentOpEnvelope[] = JSON.parse(
        JSON.stringify(redo.flat().map(toOpEnvelope)),
      );
      for (const envelope of wireRedo)
        document = applyDocumentOpEnvelope(document, envelope).unwrap()
          .document;
      expect(document).toStrictEqual(edit.document);
    }
  }
});

test("semantic table wire fixtures exercise geometry, ids, patches and tracked payloads", () => {
  const columnTable: Document = {
    package: {
      document: {
        content: [
          {
            type: "table",
            columnWidths: [600, 700, 800],
            rows: [
              {
                type: "tableRow",
                cells: ["00000001", "00000002", "00000003"].map((paraId) => ({
                  type: "tableCell",
                  content: [{ type: "paragraph", paraId, content: [] }],
                })),
              },
            ],
          },
        ],
      },
    },
  };
  const target = {
    story: "main",
    blockId: "00000003",
    column: 0,
    count: 2,
  } as const;
  // Optional counts extend the current decoder while the pinned single-column
  // specimens keep their original bytes.
  const columnRanges = [
    {
      ...target,
      type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
      width: 900,
      newBlockIds: ["00000004", "00000005"],
    },
    { ...target, type: DOCUMENT_OP_TYPES.DELETE_COLUMN },
  ] as const satisfies readonly DocumentOp[];
  for (const op of columnRanges) {
    const encoded: DocumentOpEnvelope = JSON.parse(
      JSON.stringify(toOpEnvelope(op)),
    );
    const applied = applyDocumentOpEnvelope(columnTable, encoded).unwrap();
    const table = applied.document.package.document.content.at(0);
    if (table?.type !== "table")
      throw new TypeError("Expected a retained table.");
    expect(table.columnWidths).toHaveLength(
      op.type === DOCUMENT_OP_TYPES.INSERT_COLUMN ? 5 : 1,
    );
  }
  const fixtures = [...envelopes().map(({ op }) => op), ...columnRanges];
  const fields = {
    insertColumn: {
      type: true,
      story: true,
      blockId: true,
      column: true,
      count: true,
      width: true,
      newBlockIds: true,
      revision: true,
      newIds: true,
    },
    deleteColumn: {
      type: true,
      story: true,
      blockId: true,
      column: true,
      count: true,
      revision: true,
      newIds: true,
    },
    mergeCells: {
      type: true,
      story: true,
      blockId: true,
      top: true,
      bottom: true,
      left: true,
      right: true,
      newBlockIds: true,
      revision: true,
      newIds: true,
    },
    splitCell: {
      type: true,
      story: true,
      blockId: true,
      newBlockIds: true,
      revision: true,
      newIds: true,
    },
    setTableGrid: {
      type: true,
      story: true,
      blockId: true,
      columnWidths: true,
      revision: true,
      newIds: true,
    },
    setCellProps: {
      type: true,
      story: true,
      blockId: true,
      patch: true,
      revision: true,
      newIds: true,
    },
    setRowProps: {
      type: true,
      story: true,
      blockId: true,
      patch: true,
      revision: true,
      newIds: true,
    },
    setTableProps: {
      type: true,
      story: true,
      blockId: true,
      patch: true,
      revision: true,
      newIds: true,
    },
  } as const satisfies {
    [Kind in import("../types").TableEditOp["type"]]: Record<
      keyof Extract<import("../types").TableEditOp, { type: Kind }>,
      true
    >;
  };
  for (const [kind, declared] of Object.entries(fields)) {
    const matching = fixtures.filter((op) => op.type === kind);
    expect(matching.length).toBeGreaterThan(0);
    const exercised = new Set(matching.flatMap((op) => Object.keys(op)));
    expect([...exercised].toSorted()).toEqual(Object.keys(declared).toSorted());
  }
  const restorations = fixtures.filter(
    (op) => op.type === DOCUMENT_OP_TYPES.SET_TABLE,
  );
  expect(restorations.length).toBeGreaterThan(0);
  for (const op of restorations) {
    expect(op.expected.type).toBe("table");
    expect(op.table.type).toBe("table");
  }
});

test("older envelopes are refused structurally and current envelopes apply", async () => {
  const document: Document = {
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000001", content: [] }],
      },
    },
  };
  const op = {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: { story: "main", blockId: "00000001", offset: 0 },
    text: "x",
    runProps: "inherit",
  } as const;
  // Schema 8 adds semantic table edits; schema 9 adds clipboard resources; schema 10 adds comments.
  const older: unknown = await Bun.file(
    new URL("./__fixtures__/ops-v4.json", import.meta.url),
  ).json();
  expect(Array.isArray(older)).toBe(true);
  for (const schema of [
    1,
    2,
    3,
    4,
    5,
    6,
    7,
    8,
    DOCUMENT_OP_SCHEMA_VERSION + 1,
  ]) {
    const refused = applyDocumentOpEnvelope(document, { schema, op });
    expect(refused.isErr()).toBe(true);
    if (refused.isErr()) expect(refused.error.reason).toBe("unsupportedSchema");
    expect(document.package.document.content.at(0)?.content).toEqual([]);
  }
  const current = applyDocumentOpEnvelope(document, {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    op,
  }).unwrap();
  const paragraph = current.document.package.document.content.at(0);
  expect(paragraph?.type).toBe("paragraph");
  if (paragraph?.type === "paragraph")
    expect(paragraphLogicalText(paragraph)).toBe("x");
  expect(current.inverse.length).toBe(1);
});

test("lifecycle wire fixtures exercise each declared own-undefined target", () => {
  const targets = {
    body: true,
    package: true,
    section: true,
  } as const satisfies Record<
    NonNullable<
      import("../types").StoryParts["undefinedFields"]
    >[number]["target"],
    true
  >;
  const restorations = envelopes()
    .map(({ op }) => op)
    .filter((op) => op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS);
  const exercised = new Set(
    restorations.flatMap((op) =>
      [
        ...(op.expected.undefinedFields ?? []),
        ...(op.parts.undefinedFields ?? []),
      ].map(({ target }) => target),
    ),
  );
  expect([...exercised].toSorted()).toEqual(Object.keys(targets).toSorted());
});

test("clipboard package resources remain pinned in the schema-10 wire contract", () => {
  const resources = envelopes().filter(
    ({ op }) => op.type === DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES,
  );
  expect(resources.length).toBeGreaterThan(0);
  const clipboard = resources.find(
    ({ op }) =>
      op.type === DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES &&
      op.media.entries.length > 0,
  );
  expect(clipboard).toBeDefined();
  if (clipboard?.op.type !== DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES) return;
  expect(clipboard.op.media).toMatchObject({
    expected: "omitted",
    next: "present",
    entries: [
      {
        key: "word/media/clipboard.png",
        expected: { type: "absent" },
        next: {
          type: "present",
          value: expect.objectContaining({
            mimeType: "image/png",
            data: [137, 80, 78, 71],
          }),
        },
      },
    ],
  });
  expect(() => envelopes()).not.toThrow();
});

test("schema15 table boundaries, schema17 list starts and empty replacements roundtrip while schema14 envelopes remain refused", async () => {
  const cases = tableBoundaryWireCases();
  expect(cases).toHaveLength(6);
  const first =
    cases.at(0) ?? panic("The boundary fixture needs its first real producer.");
  const seen = new Set<string>();
  for (const { before, op, edit } of cases) {
    seen.add(
      `${op.at.type}:${op.revision === undefined ? "direct" : "tracked"}`,
    );
    let document = edit.document;
    const redo: DocumentOp[] = [];
    for (const envelope of JSON.parse(
      JSON.stringify(edit.inverse.map(toOpEnvelope)),
    )) {
      const restored = applyDocumentOpEnvelope(document, envelope).unwrap();
      document = restored.document;
      redo.unshift(...restored.inverse);
    }
    expect(document).toStrictEqual(before);
    for (const envelope of JSON.parse(JSON.stringify(redo.map(toOpEnvelope))))
      document = applyDocumentOpEnvelope(document, envelope).unwrap().document;
    expect(document).toStrictEqual(edit.document);
    for (const schema of [14, DOCUMENT_OP_SCHEMA_VERSION + 1]) {
      const refused = applyDocumentOpEnvelope(before, { schema, op });
      expect(refused.isErr()).toBe(true);
      if (refused.isErr())
        expect(refused.error.reason).toBe("unsupportedSchema");
    }
  }
  expect([...seen].toSorted()).toEqual([
    "afterTable:direct",
    "afterTable:tracked",
    "beforeTable:direct",
    "beforeTable:tracked",
    "containerStart:direct",
    "containerStart:tracked",
  ]);
  const historical: unknown = await Bun.file(
    new URL("./__fixtures__/ops-v14.json", import.meta.url),
  ).json();
  if (!Array.isArray(historical))
    panic("The historical schema14 fixture must be an array.");
  for (const envelope of historical) {
    const refused = applyDocumentOpEnvelope(first.before, envelope);
    expect(refused.isErr()).toBe(true);
    if (refused.isErr()) expect(refused.error.reason).toBe("unsupportedSchema");
  }
});
