/** Forward effects are checked against logical slots and content, independently of inverse snapshots. */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES, toOpEnvelope, type DocumentOp } from "../types";

const NUM_RUNS = 120;
const id = (value: number) => value.toString(16).toUpperCase().padStart(8, "0");
const paragraph = (value: number): Paragraph => ({
  type: "paragraph",
  paraId: id(value),
  content: [{ type: "run", content: [{ type: "text", text: `cell-${value}` }] }],
});
const cell = (value: number, span = 1): TableCell => ({
  type: "tableCell",
  content: [paragraph(value)],
  ...(span === 1 ? {} : { formatting: { gridSpan: span } }),
});
const documentOf = (table: Table): Document => ({
  package: {
    document: {
      content: [paragraph(1), table, paragraph(2)],
    },
  },
});
const address = (blockId: string) => ({ story: OP_STORIES.MAIN, blockId });
const tableOf = (document: Document): Table => {
  const table = document.package.document.content.at(1);
  if (table?.type !== "table") throw new Error("Fixture table missing.");
  return table;
};
const targetOf = (table: Table) => {
  const first = table.rows.at(0)?.cells.at(0)?.content.at(0);
  if (first?.type !== "paragraph" || !first.paraId) throw new Error("Fixture target missing.");
  return first.paraId;
};

// Expanded physical content slots, including omitted slots; no production grid helper is imported.
const independentSlots = (table: Table) => {
  let width = table.columnWidths?.length;
  let previous = new Map<number, { end: number; cell: TableCell }>();
  return table.rows.map((row) => {
    const current = new Map<number, { end: number; cell: TableCell }>();
    const result: (TableCell | null)[] = Array.from(
      { length: row.formatting?.gridBefore ?? 0 },
      () => null,
    );
    for (const value of row.cells) {
      if (value.structuralChange?.type === "tableCellDeletion") continue;
      const span = value.formatting?.gridSpan ?? 1;
      expect(Number.isSafeInteger(span) && span > 0).toBe(true);
      const start = result.length;
      if (value.formatting?.vMerge === "continue") {
        const above = previous.get(start);
        expect(above?.end).toBe(start + span);
        expect(
          above?.cell.formatting?.vMerge === "restart" ||
            above?.cell.formatting?.vMerge === "continue",
        ).toBe(true);
      }
      current.set(start, { end: start + span, cell: value });
      for (let column = 0; column < span; column++) result.push(value);
    }
    for (let column = 0; column < (row.formatting?.gridAfter ?? 0); column++) result.push(null);
    width ??= result.length;
    expect(result.length).toBe(width);
    previous = current;
    return result;
  });
};
const payloads = (slots: (TableCell | null)[][]) =>
  slots.map((row) => row.map((value) => value?.content ?? null));
const check = (document: Document) => {
  expect(contractViolation(document)).toBeUndefined();
  const ids: string[] = [];
  const walk = (blocks: readonly BlockContent[]) => {
    for (const block of blocks) {
      if (block.type === "paragraph" && block.paraId) ids.push(block.paraId);
      else if (block.type === "table") {
        independentSlots(block);
        for (const row of block.rows) for (const value of row.cells) walk(value.content);
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml") walk(block.content);
    }
  };
  walk(document.package.document.content);
  expect(new Set(ids).size).toBe(ids.length);
  // Model and wire serialization are independent: either can lose data while the snapshot inverse still passes.
  expect(JSON.parse(JSON.stringify(document))).toStrictEqual(document);
};
const applied = (document: Document, op: DocumentOp) => {
  const snapshot = structuredClone(document);
  const wire = JSON.parse(JSON.stringify(toOpEnvelope(op)));
  expect(wire.op).toStrictEqual(op);
  const result = applyDocumentOp(document, wire.op);
  if (result.isErr()) throw result.error;
  check(result.value.document);
  const inverse = applyDocumentOps(result.value.document, result.value.inverse);
  if (inverse.isErr()) throw inverse.error;
  expect(inverse.value.document).toStrictEqual(snapshot);
  expect(document).toStrictEqual(snapshot);
  return result.value;
};
const refused = (document: Document, op: DocumentOp, reason: string) => {
  const snapshot = structuredClone(document);
  const result = applyDocumentOp(document, op);
  if (result.isOk()) throw new Error(`${op.type} must refuse ${reason}.`);
  expect(result.error.reason).toBe(reason);
  expect(document).toStrictEqual(snapshot);
};
const geometry = fc.record({
  rows: fc.integer({ min: 2, max: 4 }),
  before: fc.boolean(),
  after: fc.boolean(),
  span: fc.boolean(),
  nested: fc.boolean(),
  vertical: fc.boolean(),
});
type Geometry = ReturnType<typeof geometry.generate>["value"];
const fixture = ({ rows, before, after, span, nested, vertical }: Geometry) => {
  const table: Table = {
    type: "table",
    columnWidths: Array.from({ length: 3 + Number(before) + Number(after) }, () => 900),
    rows: Array.from({ length: rows }, (_, rowIndex): TableRow => {
      const cells = span
        ? [cell(0x100 + rowIndex * 16, 2), cell(0x102 + rowIndex * 16)]
        : [cell(0x100 + rowIndex * 16), cell(0x101 + rowIndex * 16), cell(0x102 + rowIndex * 16)];
      const first = cells.at(0);
      if (!first) throw new Error("Fixture first cell missing.");
      if (vertical)
        first.formatting = { ...first.formatting, vMerge: rowIndex === 0 ? "restart" : "continue" };
      if (nested)
        first.content.push(
          {
            type: "table",
            columnWidths: [300],
            rows: [{ type: "tableRow", cells: [cell(0x1000 + rowIndex)] }],
          },
          paragraph(0x2000 + rowIndex),
        );
      return {
        type: "tableRow",
        formatting: { gridBefore: Number(before), gridAfter: Number(after) },
        cells,
      };
    }),
  };
  return documentOf(table);
};

const insertionOracle = (table: Table, column: number, firstId: number) => {
  const expected = independentSlots(table);
  const newBlockIds: string[] = [];
  for (const [rowIndex, row] of expected.entries()) {
    const before = table.rows.at(rowIndex)?.formatting?.gridBefore ?? 0;
    const end = row.length - (table.rows.at(rowIndex)?.formatting?.gridAfter ?? 0);
    const left = row.at(column - 1);
    const right = row.at(column);
    let added: TableCell | null;
    if (column < before || column > end) added = null;
    else if (
      column > before &&
      column < end &&
      left === right &&
      left !== null &&
      left !== undefined
    )
      added = left;
    else {
      const paraId = id(firstId + newBlockIds.length);
      newBlockIds.push(paraId);
      added = { type: "tableCell", content: [{ type: "paragraph", paraId, content: [] }] };
    }
    row.splice(column, 0, added);
  }
  return { expected: payloads(expected), newBlockIds };
};

describe("independent table forward oracles", () => {
  test(
    "column edits preserve every surviving slot, including vertical owners, spans, omissions and nested contents",
    () => {
      // Width-only assertions admit dropping whole cells. The oracle checks every surviving physical payload.
      assertProperty(
        fc.property(geometry, fc.nat(), (shape, seed) => {
          const document = fixture(shape);
          const table = tableOf(document);
          const width = table.columnWidths?.length ?? 0;
          const column = seed % width;
          const expected = independentSlots(table);
          for (const row of expected) row.splice(column, 1);
          const deletion = applied(document, {
            ...address(targetOf(table)),
            type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
            column,
          });
          expect(payloads(independentSlots(tableOf(deletion.document)))).toStrictEqual(
            payloads(expected),
          );
          expect(tableOf(deletion.document).columnWidths).toEqual(
            Array.from({ length: width - 1 }, () => 900),
          );
          const insertion = insertionOracle(table, column, 0x3000);
          const inserted = applied(document, {
            ...address(targetOf(table)),
            type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
            column,
            width: 1200,
            newBlockIds: insertion.newBlockIds,
          });
          expect(payloads(independentSlots(tableOf(inserted.document)))).toStrictEqual(
            insertion.expected,
          );
          expect(tableOf(inserted.document).columnWidths).toEqual([
            ...Array.from({ length: column }, () => 900),
            1200,
            ...Array.from({ length: width - column }, () => 900),
          ]);
        }),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(),
  );

  test(
    "multi-step changes match an independent slot oracle after every forward step",
    () => {
      // Final rollback permits paired errors. Intermediate payload checks prevent an insertion/deletion pair masking lost data.
      assertProperty(
        fc.property(
          geometry,
          fc.array(fc.integer({ min: 0, max: 5 }), { minLength: 3, maxLength: 10 }),
          (shape, steps) => {
            let document = fixture(shape);
            const original = structuredClone(document);
            const inverses: (readonly DocumentOp[])[] = [];
            for (const [step, kind] of steps.entries()) {
              const table = tableOf(document);
              const target = address(targetOf(table));
              const width = independentSlots(table).at(0)?.length ?? 0;
              let expected = payloads(independentSlots(table));
              let op: DocumentOp;
              switch (kind) {
                case 0: {
                  const column = step % (width + 1);
                  const insertion = insertionOracle(table, column, 0x4000 + step * 16);
                  expected = insertion.expected;
                  op = {
                    ...target,
                    type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
                    column,
                    width: 700,
                    newBlockIds: insertion.newBlockIds,
                  };
                  break;
                }
                case 1:
                  op = {
                    ...target,
                    type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
                    columnWidths: Array.from({ length: width }, () => 800 + step),
                  };
                  break;
                case 2:
                  op = {
                    ...target,
                    type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
                    patch: { verticalAlign: "bottom", fitText: true },
                  };
                  break;
                case 3:
                  op = {
                    ...target,
                    type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
                    patch: { hidden: true, cantSplit: true },
                  };
                  break;
                case 4:
                  op = {
                    ...target,
                    type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
                    patch: { styleId: `forward-${step}`, layout: "autofit" },
                  };
                  break;
                default: {
                  const column = width - 1;
                  const rowCanDelete = table.rows.every(
                    (row) =>
                      (row.formatting?.gridAfter ?? 0) > 0 ||
                      row.cells.length > 1 ||
                      (row.cells.at(-1)?.formatting?.gridSpan ?? 1) > 1,
                  );
                  if (rowCanDelete && width > 2) {
                    for (const row of expected) row.splice(column, 1);
                    op = { ...target, type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column };
                  } else
                    op = {
                      ...target,
                      type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
                      columnWidths: Array.from({ length: width }, () => 1300 + step),
                    };
                }
              }
              const result = applied(document, op);
              expect(payloads(independentSlots(tableOf(result.document)))).toStrictEqual(expected);
              const after = tableOf(result.document);
              switch (op.type) {
                case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
                  expect(after.columnWidths).toStrictEqual(op.columnWidths);
                  break;
                case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
                  expect(after.rows.at(0)?.cells.at(0)?.formatting?.fitText).toBe(true);
                  break;
                case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
                  expect(after.rows.at(0)?.formatting?.hidden).toBe(true);
                  break;
                case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
                  expect(after.formatting?.styleId).toBe(`forward-${step}`);
                  break;
              }
              inverses.push(result.inverse);
              document = result.document;
            }
            const undone = applyDocumentOps(document, inverses.reverse().flat());
            if (undone.isErr()) throw undone.error;
            expect(undone.value.document).toStrictEqual(original);
          },
        ),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(),
  );

  test(
    "mixed-chain merges preserve nested content; horizontal splits preserve owners and refuse vertical groups",
    () => {
      // A merge that drops continuation/nested content passes snapshot restoration. Pin every moved block and new empty slot.
      assertProperty(
        fc.property(geometry, (shape) => {
          const document = fixture({ ...shape, vertical: true });
          const table = tableOf(document);
          const left = Number(shape.before);
          const content = table.rows.flatMap((row) => row.cells.flatMap((value) => value.content));
          const merged = applied(document, {
            ...address(targetOf(table)),
            type: DOCUMENT_OP_TYPES.MERGE_CELLS,
            top: 0,
            bottom: shape.rows,
            left,
            right: left + 3,
            newBlockIds: Array.from({ length: shape.rows - 1 }, (_, index) => id(0x5000 + index)),
          });
          const mergedTable = tableOf(merged.document);
          const mergeExpected = table.rows.map((_, rowIndex) => {
            const blocks =
              rowIndex === 0
                ? content
                : [{ type: "paragraph", paraId: id(0x5000 + rowIndex - 1), content: [] }];
            return [
              ...Array.from({ length: left }, () => null),
              blocks,
              blocks,
              blocks,
              ...Array.from({ length: Number(shape.after) }, () => null),
            ];
          });
          expect(payloads(independentSlots(mergedTable))).toStrictEqual(mergeExpected);
          for (const [rowIndex, row] of mergedTable.rows.entries()) {
            expect(row.cells.length).toBe(1);
            expect(row.cells.at(0)?.formatting?.gridSpan).toBe(3);
            expect(row.cells.at(0)?.formatting?.vMerge).toBe(
              rowIndex === 0 ? "restart" : "continue",
            );
          }
          for (const row of mergedTable.rows) {
            const first = row.cells.at(0)?.content.at(0);
            if (first?.type !== "paragraph" || !first.paraId)
              throw new Error("Merged owner paragraph missing.");
            refused(
              merged.document,
              { ...address(first.paraId), type: DOCUMENT_OP_TYPES.SPLIT_CELL, newBlockIds: [] },
              DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
            );
          }
          const horizontalDocument = fixture({ ...shape, vertical: false });
          const horizontalTable = tableOf(horizontalDocument);
          const horizontalMerge = applied(horizontalDocument, {
            ...address(targetOf(horizontalTable)),
            type: DOCUMENT_OP_TYPES.MERGE_CELLS,
            top: 0,
            bottom: 1,
            left,
            right: left + 3,
            newBlockIds: [],
          });
          const horizontalMerged = tableOf(horizontalMerge.document);
          const split = applied(horizontalMerge.document, {
            ...address(targetOf(horizontalMerged)),
            type: DOCUMENT_OP_TYPES.SPLIT_CELL,
            newBlockIds: [id(0x6000), id(0x6001)],
          });
          const splitExpected = payloads(independentSlots(horizontalMerged));
          splitExpected.splice(0, 1, [
            ...Array.from({ length: left }, () => null),
            horizontalMerged.rows.at(0)?.cells.at(0)?.content ?? [],
            [{ type: "paragraph", paraId: id(0x6000), content: [] }],
            [{ type: "paragraph", paraId: id(0x6001), content: [] }],
            ...Array.from({ length: Number(shape.after) }, () => null),
          ]);
          const after = tableOf(split.document);
          expect(payloads(independentSlots(after))).toStrictEqual(splitExpected);
          for (const value of after.rows.at(0)?.cells ?? []) {
            expect(value.formatting?.gridSpan ?? 1).toBe(1);
            expect(value.formatting?.vMerge).toBeUndefined();
          }
          const restored = applied(split.document, {
            ...address(targetOf(after)),
            type: DOCUMENT_OP_TYPES.SET_TABLE,
            expected: after,
            table,
          });
          expect(restored.document).toStrictEqual(document);
        }),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(),
  );

  test(
    "nonrectangular merge selections and row-emptying deletion refuse their exact reasons",
    () => {
      // A fixture with no omissions cannot detect clipped-grid merges; asymmetric rows and active omissions close that gap.
      assertProperty(
        fc.property(geometry, (shape) => {
          const document = fixture({
            ...shape,
            before: true,
            after: true,
            span: true,
            vertical: true,
          });
          const table = tableOf(document);
          const ops = [
            {
              ...address(targetOf(table)),
              type: DOCUMENT_OP_TYPES.MERGE_CELLS,
              top: 0,
              bottom: shape.rows,
              left: 0,
              right: 4,
              newBlockIds: [],
            },
            {
              ...address(targetOf(table)),
              type: DOCUMENT_OP_TYPES.MERGE_CELLS,
              top: 0,
              bottom: shape.rows,
              left: 2,
              right: 4,
              newBlockIds: [],
            },
            {
              ...address(targetOf(table)),
              type: DOCUMENT_OP_TYPES.MERGE_CELLS,
              top: 0,
              bottom: 1,
              left: 1,
              right: 4,
              newBlockIds: [],
            },
          ] as const satisfies readonly DocumentOp[];
          for (const op of ops)
            refused(document, op, DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
          const sparse: Table = {
            type: "table",
            columnWidths: [900, 900, 900],
            rows: [
              {
                type: "tableRow",
                formatting: { gridBefore: 1, gridAfter: 1 },
                cells: [cell(0x100)],
              },
            ],
          };
          refused(
            documentOf(sparse),
            { ...address(id(0x100)), type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column: 1 },
            DOCUMENT_OP_REFUSAL_REASONS.TABLE_ROW_EMPTY,
          );
        }),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(),
  );

  test("indexed markup and existing reviews have consistent refusals across topology edits", () => {
    // Different op-specific guards can assign different reasons to the same cause; exercise every affected family.
    const counts = new Map<string, number>();
    for (const cause of ["markup", "review"] as const) {
      const document = fixture({
        rows: 2,
        before: false,
        after: false,
        span: true,
        nested: true,
        vertical: false,
      });
      const table = tableOf(document);
      const target = address(targetOf(table));
      const reason =
        cause === "markup"
          ? DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE
          : DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT;
      if (cause === "markup")
        table.bookmarks = [
          { index: 0, marker: { type: "bookmarkStart", id: 42, name: "_RefWholeTable" } },
          { index: table.rows.length, marker: { type: "bookmarkEnd", id: 42 } },
        ];
      else
        table.propertyChanges = [
          {
            type: "tablePropertyChange",
            info: { id: 42, author: "Reviewer", date: "2026-10-02T10:00:00Z" },
            previousFormatting: { layout: "fixed" },
          },
        ];
      const operations = [
        {
          ...target,
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          column: 0,
          width: 900,
          newBlockIds: [id(0xa000), id(0xa001)],
        },
        { ...target, type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column: 2 },
        {
          ...target,
          type: DOCUMENT_OP_TYPES.MERGE_CELLS,
          top: 0,
          bottom: 1,
          left: 0,
          right: 3,
          newBlockIds: [],
        },
        { ...target, type: DOCUMENT_OP_TYPES.SPLIT_CELL, newBlockIds: [id(0xa000)] },
      ] as const satisfies readonly DocumentOp[];
      for (const op of operations) {
        refused(document, op, reason);
        counts.set(reason, (counts.get(reason) ?? 0) + 1);
      }
    }
    expect(counts.get(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE)).toBe(4);
    expect(counts.get(DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT)).toBe(4);
    expect(counts.size).toBe(2);
  });

  test("sparse raw identity, width and revision arrays refuse their shape before use", () => {
    // Array.every skips holes; sparse production-shaped payloads close the gap left by ordinary missing-field cases.
    const document = fixture({
      rows: 2,
      before: false,
      after: false,
      span: true,
      nested: false,
      vertical: false,
    });
    const target = address(targetOf(tableOf(document)));
    const sparseIds = [id(0xa000), id(0xa001)];
    Reflect.deleteProperty(sparseIds, "0");
    const sparseWidths = [900, 900, 900];
    Reflect.deleteProperty(sparseWidths, "1");
    const sparseRevisions = [0xb001, 0xb002, 0xb003];
    Reflect.deleteProperty(sparseRevisions, "0");
    const operations = [
      {
        ...target,
        type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
        column: 0,
        width: 900,
        newBlockIds: sparseIds,
      },
      { ...target, type: DOCUMENT_OP_TYPES.SET_TABLE_GRID, columnWidths: sparseWidths },
      {
        ...target,
        type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
        columnWidths: [900, 950, 1000],
        revision: { id: 0xb000, author: "Reviewer", date: "2026-10-02T10:00:00Z" },
        newIds: { revision: sparseRevisions },
      },
    ] as const satisfies readonly DocumentOp[];
    for (const op of operations)
      refused(document, op, DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION);
  });

  test("required raw payloads refuse missing and null fields without mutation or throws", () => {
    // Typed callers hide malformed JSON. Delete required keys from otherwise valid operation-shaped values at the boundary.
    const document = fixture({
      rows: 2,
      before: false,
      after: false,
      span: true,
      nested: false,
      vertical: true,
    });
    const target = address(targetOf(tableOf(document)));
    const cases = [
      {
        op: {
          ...target,
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          column: 0,
          width: 900,
          newBlockIds: [id(0x7000), id(0x7001)],
        },
        fields: ["newBlockIds", "column", "width"],
      },
      { op: { ...target, type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column: 0 }, fields: ["column"] },
      {
        op: {
          ...target,
          type: DOCUMENT_OP_TYPES.MERGE_CELLS,
          top: 0,
          bottom: 2,
          left: 0,
          right: 3,
          newBlockIds: [id(0x7000)],
        },
        fields: ["top", "bottom", "left", "right", "newBlockIds"],
      },
      {
        op: {
          ...target,
          type: DOCUMENT_OP_TYPES.SPLIT_CELL,
          newBlockIds: [id(0x7000), id(0x7001)],
        },
        fields: ["newBlockIds"],
      },
      {
        op: { ...target, type: DOCUMENT_OP_TYPES.SET_TABLE_GRID, columnWidths: [900, 900, 900] },
        fields: ["columnWidths"],
      },
      {
        op: { ...target, type: DOCUMENT_OP_TYPES.SET_CELL_PROPS, patch: { fitText: true } },
        fields: ["patch"],
      },
      {
        op: { ...target, type: DOCUMENT_OP_TYPES.SET_ROW_PROPS, patch: { hidden: true } },
        fields: ["patch"],
      },
      {
        op: { ...target, type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS, patch: { bidi: true } },
        fields: ["patch"],
      },
    ] as const satisfies readonly { op: DocumentOp; fields: readonly string[] }[];
    let refusals = 0;
    for (const { op, fields } of cases)
      for (const field of ["story", "blockId", ...fields]) {
        const missing = structuredClone(op);
        Reflect.deleteProperty(missing, field);
        refused(document, missing, DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION);
        const nullField = structuredClone(op);
        Reflect.set(nullField, field, null);
        refused(document, nullField, DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION);
        refusals += 2;
      }
    expect(refusals).toBe(cases.reduce((count, entry) => count + 2 * (entry.fields.length + 2), 0));
  });

  test("revision id collisions refuse locally and identical tracked patches create no records", () => {
    // Package-uniqueness checks cannot detect duplicate unused ids before creation. Exercise duplicate pools explicitly.
    const document = fixture({
      rows: 2,
      before: false,
      after: false,
      span: false,
      nested: false,
      vertical: false,
    });
    const target = address(targetOf(tableOf(document)));
    const revision = { id: 0x8000, author: "Reviewer", date: "2026-10-02T10:00:00Z" };
    for (const ids of [
      [0x8001, 0x8001, 0x8002],
      [0x8000, 0x8001, 0x8002],
    ]) {
      refused(
        document,
        {
          ...target,
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          column: 0,
          width: 900,
          newBlockIds: [id(0x9000), id(0x9001)],
          revision,
          newIds: { revision: ids },
        },
        DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      );
    }
    const table = tableOf(document);
    const row = table.rows.at(0);
    const value = row?.cells.at(0);
    if (!row || !value) throw new Error("Fixture cell missing.");
    value.formatting = { fitText: true };
    row.formatting = { cantSplit: true };
    table.formatting = { bidi: true };
    const ops = [
      { ...target, type: DOCUMENT_OP_TYPES.SET_CELL_PROPS, patch: { fitText: true }, revision },
      { ...target, type: DOCUMENT_OP_TYPES.SET_ROW_PROPS, patch: { cantSplit: true }, revision },
      { ...target, type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS, patch: { bidi: true }, revision },
    ] as const satisfies readonly DocumentOp[];
    for (const op of ops) {
      const result = applied(document, op);
      expect(result.document).toStrictEqual(document);
      expect(result.inverse).toEqual([]);
      expect(result.revisions ?? []).toEqual([]);
      expect(result.touched).toEqual({ modified: [], inserted: [], removed: [] });
    }
  });

  test("splitting a unit cell is an explicit no-change refusal", () => {
    // A splitter can silently strip provenance on a 1x1 cell; no-op geometry must stop before mutation.
    const document = fixture({
      rows: 2,
      before: false,
      after: false,
      span: false,
      nested: true,
      vertical: false,
    });
    const table = tableOf(document);
    const first = table.rows.at(0)?.cells.at(0);
    if (!first) throw new Error("Fixture cell missing.");
    first.formatting = { sourceXml: '<w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr>' };
    refused(
      document,
      { ...address(targetOf(table)), type: DOCUMENT_OP_TYPES.SPLIT_CELL, newBlockIds: [] },
      DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
    );
  });

  test("deletion refuses bookmark, comment and review dependencies but permits span narrowing", () => {
    // Crossing anchors must use producer-shaped range ids; dropping one endpoint would otherwise pass grid/inverse checks.
    for (const kind of ["bookmark", "comment", "revision"] as const) {
      const removed = cell(0x100);
      const survivor = cell(0x101);
      const first = removed.content.at(0);
      const second = survivor.content.at(0);
      if (first?.type !== "paragraph" || second?.type !== "paragraph")
        throw new Error("Fixture paragraph missing.");
      switch (kind) {
        case "bookmark":
          first.content.unshift({ type: "bookmarkStart", id: 42, name: "_RefCrossCell" });
          second.content.push({ type: "bookmarkEnd", id: 42 });
          break;
        case "comment":
          first.content.unshift({ type: "commentRangeStart", id: 42 });
          second.content.push({ type: "commentRangeEnd", id: 42 });
          break;
        case "revision":
          first.content = [
            {
              type: "insertion",
              info: { id: 42, author: "Reviewer", date: "2026-10-02T10:00:00Z" },
              content: first.content,
            },
          ];
          break;
      }
      const table: Table = {
        type: "table",
        columnWidths: [900, 900],
        rows: [{ type: "tableRow", cells: [removed, survivor] }],
      };
      const document = documentOf(table);
      refused(
        document,
        { ...address(id(0x100)), type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column: 0 },
        DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
      );
      const sole: Table = {
        type: "table",
        columnWidths: [900],
        rows: [{ type: "tableRow", cells: [removed] }],
      };
      const soleDocument = documentOf(sole);
      soleDocument.package.document.content.push(second);
      refused(
        soleDocument,
        { ...address(id(0x100)), type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column: 0 },
        DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
      );
      removed.formatting = { gridSpan: 2 };
      table.columnWidths = [900, 900, 900];
      const narrowed = applied(document, {
        ...address(id(0x100)),
        type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
        column: 1,
      });
      expect(tableOf(narrowed.document).rows.at(0)?.cells.at(0)?.content).toStrictEqual(
        removed.content,
      );
      expect(tableOf(narrowed.document).rows.at(0)?.cells.at(0)?.formatting?.gridSpan ?? 1).toBe(1);
    }
  });
});
