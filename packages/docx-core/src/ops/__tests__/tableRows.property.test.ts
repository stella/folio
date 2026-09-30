/** Property laws for the row-level document operations. */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { Document, Paragraph, Table, TableRow } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import { paragraphIdsIn } from "../ids";
import { DOCUMENT_OP_TYPES, OP_STORIES, type DocumentOp, type RevisionStamp } from "../types";

setDefaultTimeout(propertyTestTimeout(240_000));

// CI scales runs (5x on PRs, 10x nightly), so the laws still see 10^4+ cases per run there.
const NUM_RUNS = 2_000;
const DATE = "2026-05-06T07:08:09Z";
const stamp = (id: number): RevisionStamp => ({ id, author: "Reviewer", date: DATE });
const newIds = (start: number) => ({
  revision: Array.from({ length: 128 }, (_, index) => start + index),
});

type Shape = {
  rows: number;
  cells: number;
  paragraphs: number;
  text: number;
  formatting: number;
};

const shapeArbitrary: fc.Arbitrary<Shape> = fc.record({
  rows: fc.integer({ min: 2, max: 4 }),
  cells: fc.integer({ min: 1, max: 3 }),
  paragraphs: fc.integer({ min: 1, max: 3 }),
  text: fc.integer({ min: 0, max: 0xffff_ffff }),
  formatting: fc.integer({ min: 0, max: 0xffff_ffff }),
});

const formatParagraphId = (value: number): string =>
  value.toString(16).toUpperCase().padStart(8, "0");

const runFormatting = (seed: number) => {
  switch (seed % 3) {
    case 0:
      return { formatting: { bold: true } };
    case 1:
      return { formatting: { italic: true } };
    default:
      return {};
  }
};

const makeParagraph = (paragraphId: string, text: string, formatting: number): Paragraph => ({
  type: "paragraph",
  paraId: paragraphId,
  content: [
    {
      type: "run",
      ...runFormatting(formatting),
      content: [{ type: "text", text }],
    },
  ],
});

const rowFor = (shape: Shape, rowIndex: number, firstId: number): TableRow => ({
  type: "tableRow",
  formatting: rowIndex % 2 === 0 ? { cantSplit: true } : { header: true },
  preservedAttributes: [{ name: "rsidTr", value: "00ABCDEF" }],
  cells: Array.from({ length: 1 + ((shape.cells + rowIndex) % 3) }, (_, cellIndex) => ({
    type: "tableCell" as const,
    content: Array.from(
      { length: 1 + ((shape.paragraphs + rowIndex + cellIndex) % 3) },
      (_paragraphValue, paragraphIndex) => {
        const paragraphId = firstId + cellIndex * 9 + paragraphIndex;
        const text = `cell-${rowIndex}-${cellIndex}-${paragraphIndex}-${shape.text % 997}`;
        return makeParagraph(formatParagraphId(paragraphId), text, shape.formatting + paragraphId);
      },
    ),
  })),
});

type Fixture = { document: Document; rows: TableRow[]; nextId: number };

const fixtureFor = (shape: Shape): Fixture => {
  const rows = Array.from({ length: shape.rows }, (_, rowIndex) =>
    rowFor(shape, rowIndex, 0x100 + rowIndex * 27),
  );
  const table: Table = { type: "table", formatting: { justification: "center" }, rows };
  const before = makeParagraph("00000001", "outside-before", shape.formatting);
  const after = makeParagraph("00000002", "outside-after", shape.formatting + 1);
  return {
    document: { package: { document: { content: [before, table, after] } } },
    rows,
    nextId: 0x100 + shape.rows * 27,
  };
};

const operationTarget = (row: TableRow): string => {
  const paragraphId = paragraphIdsIn(row).at(0);
  if (paragraphId === undefined) throw new Error("Generated rows have a paragraph target.");
  return paragraphId;
};

const tableOf = (document: Document): Table => {
  const table = document.package.document.content.find((block) => block.type === "table");
  if (table?.type !== "table") throw new Error("Generated document has a table.");
  return table;
};

const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  return result.value;
};

const resolveOperation = (
  revisionIds: readonly number[],
  decision: "accept" | "reject",
): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
  story: OP_STORIES.MAIN,
  revisionIds,
  decision,
});

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: "accept" | "reject",
) => {
  if (revisionIds.length === 0) return document;
  return applied(document, resolveOperation(revisionIds, decision)).document;
};

const directOf = (op: DocumentOp): DocumentOp => {
  const direct = structuredClone(op);
  Reflect.deleteProperty(direct, "revision");
  Reflect.deleteProperty(direct, "newIds");
  return direct;
};

const operationFor = (
  shape: Shape,
  seed: number,
  kind: "insertRow" | "deleteRow" | "setTableRows",
) => {
  const fixture = fixtureFor(shape);
  const revision = stamp(100_000);
  const ids = newIds(100_001);
  if (kind === "insertRow") {
    const at = seed % (fixture.rows.length + 1);
    const row = rowFor(shape, fixture.rows.length + 1, fixture.nextId);
    const anchor = fixture.rows[Math.min(at, fixture.rows.length - 1)];
    if (anchor === undefined) throw new Error("Generated table has an insertion anchor.");
    return {
      ...fixture,
      op: {
        type: DOCUMENT_OP_TYPES.INSERT_ROW,
        story: OP_STORIES.MAIN,
        blockId: operationTarget(anchor),
        at,
        row,
        revision,
        newIds: ids,
      } satisfies DocumentOp,
    };
  }
  const targetIndex = seed % fixture.rows.length;
  const target = fixture.rows[targetIndex];
  if (target === undefined) throw new Error("Generated table has a target row.");
  if (kind === "deleteRow") {
    return {
      ...fixture,
      op: {
        type: DOCUMENT_OP_TYPES.DELETE_ROW,
        story: OP_STORIES.MAIN,
        blockId: operationTarget(target),
        expected: target,
        revision,
        newIds: ids,
      } satisfies DocumentOp,
    };
  }
  const rows = fixture.rows.map((row, index) =>
    index === targetIndex ? rowFor(shape, index + 10, fixture.nextId) : row,
  );
  return {
    ...fixture,
    op: {
      type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
      story: OP_STORIES.MAIN,
      blockId: operationTarget(target),
      expected: fixture.rows,
      rows,
    } satisfies DocumentOp,
  };
};

const rowHasRevision = (
  row: TableRow,
  type: "tableRowInsertion" | "tableRowDeletion",
  id: number,
): boolean => row.structuralChange?.type === type && row.structuralChange.info.id === id;

const rowParagraphs = (row: TableRow): Paragraph[] =>
  row.cells.flatMap((cell) =>
    cell.content.flatMap((block) => (block.type === "paragraph" ? [block] : [])),
  );

const assertTrackedRow = (
  row: TableRow,
  type: "tableRowInsertion" | "tableRowDeletion",
  wrapper: "insertion" | "deletion",
  revisionIds: ReadonlySet<number>,
) => {
  expect(rowHasRevision(row, type, 100_000)).toBe(true);
  const wrapperIds = new Set<number>();
  for (const paragraph of rowParagraphs(row)) {
    if (paragraph.content.length === 0) continue;
    expect(paragraph.pPrMark).toBeUndefined();
    const wrappers = paragraph.content.filter((content) => content.type === wrapper);
    expect(wrappers.length).toBeGreaterThan(0);
    for (const content of wrappers) {
      expect(content.info.author).toBe("Reviewer");
      expect(content.info.date).toBe(DATE);
      expect(revisionIds.has(content.info.id)).toBe(true);
      expect(content.info.id).not.toBe(100_000);
      expect(wrapperIds.has(content.info.id)).toBe(false);
      wrapperIds.add(content.info.id);
    }
  }
  expect(wrapperIds.size).toBeGreaterThan(0);
};

const operationFamilies = ["insertRow", "deleteRow", "setTableRows"] as const;

describe("table row operation properties", () => {
  test("removing a table's final row removes the table with an exact inverse", () => {
    const fixture = fixtureFor({ rows: 1, cells: 1, paragraphs: 1, text: 0, formatting: 0 });
    const onlyRow = fixture.rows.at(0);
    if (onlyRow === undefined) throw new Error("Generated table has its only row.");
    const result = applied(fixture.document, {
      type: DOCUMENT_OP_TYPES.DELETE_ROW,
      story: OP_STORIES.MAIN,
      blockId: operationTarget(onlyRow),
      expected: onlyRow,
    });
    expect(result.document.package.document.content.map((block) => block.type)).toEqual([
      "paragraph",
      "paragraph",
    ]);
    const undone = applyDocumentOps(result.document, result.inverse);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toStrictEqual(fixture.document);
  });

  for (const kind of operationFamilies) {
    test(`${kind}: L1 acceptance, L2 rejection, inverse, determinism and locality`, () => {
      assertProperty(
        fc.property(shapeArbitrary, fc.integer({ min: 0, max: 0xffff_ffff }), (shape, seed) => {
          const { document, op } = operationFor(shape, seed, kind);
          const direct = applied(document, directOf(op));
          const tracked = applied(document, op);
          if (kind === "setTableRows") {
            expect(tracked.document).toStrictEqual(direct.document);
            expect(tracked.revisions).toEqual([]);
          } else {
            expect(tracked.revisions.length).toBeGreaterThan(0);
            expect(resolve(tracked.document, tracked.revisions, "accept")).toStrictEqual(
              direct.document,
            );
            expect(resolve(tracked.document, tracked.revisions, "reject")).toStrictEqual(document);

            const revisionIds = new Set(tracked.revisions);
            const expectedStructuralChange =
              kind === "insertRow" ? "tableRowInsertion" : "tableRowDeletion";
            const expectedWrapper = kind === "insertRow" ? "insertion" : "deletion";
            const changedRows = tableOf(tracked.document).rows.filter((row) =>
              rowHasRevision(row, expectedStructuralChange, 100_000),
            );
            expect(changedRows.length).toBeGreaterThan(0);
            for (const row of changedRows) {
              assertTrackedRow(row, expectedStructuralChange, expectedWrapper, revisionIds);
            }
          }

          const undone = applyDocumentOps(tracked.document, tracked.inverse);
          if (undone.isErr()) throw undone.error;
          expect(undone.value.document).toStrictEqual(document);

          if (kind !== "setTableRows") {
            for (const decision of ["accept", "reject"] as const) {
              const resolved = applied(
                tracked.document,
                resolveOperation(tracked.revisions, decision),
              );
              const restored = applyDocumentOps(resolved.document, resolved.inverse);
              if (restored.isErr()) throw restored.error;
              expect(restored.value.document).toStrictEqual(tracked.document);
            }
          }

          const repeated = applied(structuredClone(document), JSON.parse(JSON.stringify(op)));
          expect(repeated.document).toStrictEqual(tracked.document);
          expect(repeated.inverse).toStrictEqual(tracked.inverse);

          const touched = new Set([
            ...tracked.touched.modified,
            ...tracked.touched.inserted,
            ...tracked.touched.removed,
          ]);
          const beforeParagraphs = storyParagraphs(document.package.document);
          const afterParagraphs = new Map(
            storyParagraphs(tracked.document.package.document).map(({ paragraph }) => [
              paragraph.paraId ?? "",
              paragraph,
            ]),
          );
          for (const { paragraph } of beforeParagraphs) {
            if (touched.has(paragraph.paraId ?? "")) continue;
            expect(afterParagraphs.get(paragraph.paraId ?? "")).toBe(paragraph);
          }
          for (const { paragraph } of storyParagraphs(tracked.document.package.document)) {
            expect(paragraph.pPrMark).toBeUndefined();
          }

          if (kind === "insertRow") {
            const row = tableOf(tracked.document).rows.find((candidate) =>
              rowHasRevision(candidate, "tableRowInsertion", 100_000),
            );
            expect(row).toBeDefined();
            if (row === undefined) throw new Error("Tracked insertion retains its row.");
            expect(rowHasRevision(row, "tableRowInsertion", 100_000)).toBe(true);
          } else if (kind === "deleteRow") {
            const row = tableOf(tracked.document).rows.find((candidate) =>
              rowHasRevision(candidate, "tableRowDeletion", 100_000),
            );
            expect(row).toBeDefined();
            if (row === undefined) throw new Error("Tracked deletion retains its row.");
            expect(rowHasRevision(row, "tableRowDeletion", 100_000)).toBe(true);
          }

          if (kind !== "setTableRows") {
            const rejectedOnce = resolve(tracked.document, tracked.revisions, "reject");
            expect(resolve(rejectedOnce, tracked.revisions, "reject")).toStrictEqual(document);
          }
        }),
        { numRuns: NUM_RUNS },
      );
    });
  }

  test("tracked insertion then deletion of another original row obeys the batch law", () => {
    assertProperty(
      fc.property(
        shapeArbitrary,
        fc.integer({ min: 0, max: 0xffff_ffff }),
        fc.integer({ min: 0, max: 0xffff_ffff }),
        (shape, insertionSeed, deletionSeed) => {
          const fixture = fixtureFor(shape);
          const insertionIndex = insertionSeed % (fixture.rows.length + 1);
          const insertedRow = rowFor(shape, fixture.rows.length + 1, fixture.nextId);
          const originalIndex = deletionSeed % fixture.rows.length;
          const originalRow = fixture.rows[originalIndex];
          if (originalRow === undefined) throw new Error("Generated table has an original row.");
          const shiftedIndex = originalIndex >= insertionIndex ? originalIndex + 1 : originalIndex;
          const insert: DocumentOp = {
            type: DOCUMENT_OP_TYPES.INSERT_ROW,
            story: OP_STORIES.MAIN,
            blockId: operationTarget(originalRow),
            at: insertionIndex,
            row: insertedRow,
            revision: stamp(100_000),
            newIds: newIds(100_001),
          };
          const afterInsertRows = [...fixture.rows];
          afterInsertRows.splice(insertionIndex, 0, insertedRow);
          const deletionTarget = afterInsertRows[shiftedIndex];
          if (deletionTarget === undefined)
            throw new Error("Generated batch has a deletion target.");
          const remove: DocumentOp = {
            type: DOCUMENT_OP_TYPES.DELETE_ROW,
            story: OP_STORIES.MAIN,
            blockId: operationTarget(deletionTarget),
            expected: deletionTarget,
            revision: stamp(200_000),
            newIds: newIds(200_001),
          };

          const first = applied(fixture.document, insert);
          const trackedBatch = applied(first.document, remove);
          const directFirst = applied(fixture.document, directOf(insert));
          const directBatch = applied(directFirst.document, directOf(remove));
          const revisionIds = [...new Set([...first.revisions, ...trackedBatch.revisions])];
          const undoneBatch = applyDocumentOps(trackedBatch.document, [
            ...trackedBatch.inverse,
            ...first.inverse,
          ]);
          if (undoneBatch.isErr()) throw undoneBatch.error;
          expect(undoneBatch.value.document).toStrictEqual(fixture.document);
          expect(resolve(trackedBatch.document, revisionIds, "accept")).toStrictEqual(
            directBatch.document,
          );
          expect(resolve(trackedBatch.document, revisionIds, "reject")).toStrictEqual(
            fixture.document,
          );

          const acceptedOnce = resolve(trackedBatch.document, revisionIds, "accept");
          expect(resolve(acceptedOnce, revisionIds, "accept")).toStrictEqual(acceptedOnce);
        },
      ),
      { numRuns: NUM_RUNS },
    );
  });
});
