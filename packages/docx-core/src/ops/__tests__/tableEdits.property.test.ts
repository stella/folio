/** Table edit laws use an independent slot oracle, never the production grid helper. */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  Run,
  Table,
  TableCell,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  type TableIntentOperation,
} from "../editorIntent";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES, toOpEnvelope, type DocumentOp } from "../types";

setDefaultTimeout(propertyTestTimeout(120_000));
const NUM_RUNS = 160;
const id = (value: number) => value.toString(16).toUpperCase().padStart(8, "0");
const paragraph = (value: number, empty: boolean): Paragraph => ({
  type: "paragraph",
  paraId: id(value),
  content: empty ? [] : [{ type: "run", content: [{ type: "text", text: `payload-${value}` }] }],
});

// toEqual and JSON equality can erase an own undefined key. Own-key recursion closes that gap.
const exact = (actual: unknown, expected: unknown): void => {
  expect(actual === null).toBe(expected === null);
  expect(typeof actual).toBe(typeof expected);
  if (
    actual === null ||
    expected === null ||
    typeof actual !== "object" ||
    typeof expected !== "object"
  ) {
    expect(actual).toBe(expected);
    return;
  }
  expect(Array.isArray(actual)).toBe(Array.isArray(expected));
  const keys = Object.keys(expected).sort();
  expect(Object.keys(actual).sort()).toEqual(keys);
  for (const key of keys) exact(Reflect.get(actual, key), Reflect.get(expected, key));
};
const shape = fc.record({
  width: fc.integer({ min: 3, max: 6 }),
  height: fc.integer({ min: 2, max: 4 }),
  variant: fc.integer({ min: 0, max: 3 }),
  nested: fc.boolean(),
  empty: fc.boolean(),
  explicitGrid: fc.boolean(),
});
type Shape = ReturnType<typeof shape.generate>["value"];
const fixture = ({ width, height, variant, nested, empty, explicitGrid }: Shape) => {
  let next = 0x100;
  const cell = (span: number, row: number): TableCell => ({
    type: "tableCell",
    content: [paragraph(next++, empty)],
    ...(span === 1 && variant !== 2
      ? {}
      : {
          formatting: {
            ...(span === 1 ? {} : { gridSpan: span }),
            ...(variant === 2 ? { vMerge: row === 0 ? "restart" : "continue" } : {}),
          },
        }),
  });
  const table: Table = {
    type: "table",
    formatting: { justification: "center", layout: "fixed" },
    ...(explicitGrid ? { columnWidths: Array.from({ length: width }, () => 700) } : {}),
    rows: Array.from({ length: height }, (_, row) => {
      const omitted = variant === 3;
      const spans =
        variant === 1 || variant === 2
          ? [2, ...Array.from({ length: width - 2 }, () => 1)]
          : Array.from({ length: width - (omitted ? 2 : 0) }, () => 1);
      return {
        type: "tableRow",
        formatting: { header: row === 0, ...(omitted ? { gridBefore: 1, gridAfter: 1 } : {}) },
        cells: spans.map((span) => cell(span, row)),
      };
    }),
  };
  const wrapped: Table = {
    type: "table",
    columnWidths: [4000],
    rows: [
      {
        type: "tableRow",
        cells: [{ type: "tableCell", content: [table, paragraph(0x20, false)] }],
      },
    ],
  };
  const document: Document = {
    package: {
      document: {
        content: [paragraph(1, false), nested ? wrapped : table, paragraph(2, false)],
      },
    },
  };
  return { document, nested, target: id(0x100), width, height };
};
const tableOf = (document: Document, nested: boolean): Table => {
  const outer = document.package.document.content.at(1);
  if (outer?.type !== "table") throw new Error("Fixture table missing.");
  if (!nested) return outer;
  const inner = outer.rows.at(0)?.cells.at(0)?.content.at(0);
  if (inner?.type !== "table") throw new Error("Nested fixture table missing.");
  return inner;
};
const targetOf = (table: Table): string => {
  const first = table.rows.at(0)?.cells.at(0)?.content.at(0);
  if (first?.type !== "paragraph" || !first.paraId) throw new Error("Target paragraph missing.");
  return first.paraId;
};

// An inverse alone accepts a no-op, or two mutually wrong algorithms. Slots and effects are separate oracles.
const slots = (table: Table) => {
  let width: number | undefined = table.columnWidths?.length;
  const rows = table.rows.map((row) => {
    const expanded: (TableCell | undefined)[] = Array.from({
      length: row.formatting?.gridBefore ?? 0,
    });
    for (const cell of row.cells) {
      if (cell.structuralChange?.type === "tableCellDeletion") continue;
      const span = cell.formatting?.gridSpan ?? 1;
      expect(Number.isSafeInteger(span) && span > 0).toBe(true);
      for (let offset = 0; offset < span; offset++) expanded.push(cell);
    }
    for (let offset = 0; offset < (row.formatting?.gridAfter ?? 0); offset++)
      expanded.push(undefined);
    width ??= expanded.length;
    expect(expanded.length).toBe(width);
    return expanded;
  });
  for (const [rowIndex, row] of rows.entries()) {
    for (const [column, cell] of row.entries()) {
      if (cell?.formatting?.vMerge !== "continue") continue;
      const above = rows.at(rowIndex - 1)?.at(column);
      expect(rowIndex).toBeGreaterThan(0);
      expect(
        above?.formatting?.vMerge === "restart" || above?.formatting?.vMerge === "continue",
      ).toBe(true);
      expect(above?.formatting?.gridSpan ?? 1).toBe(cell.formatting.gridSpan ?? 1);
      expect(row.indexOf(cell)).toBe(rows.at(rowIndex - 1)?.indexOf(above));
    }
  }
  return rows;
};
const invariants = (document: Document) => {
  expect(contractViolation(document)).toBeUndefined();
  const ids: string[] = [];
  const walk = (blocks: readonly BlockContent[]) => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        if (block.paraId) ids.push(block.paraId);
      } else if (block.type === "table") {
        slots(block);
        for (const row of block.rows) for (const cell of row.cells) walk(cell.content);
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml") walk(block.content);
    }
  };
  walk(document.package.document.content);
  expect(new Set(ids).size).toBe(ids.length);
};
const applied = (document: Document, op: DocumentOp) => {
  const original = structuredClone(document);
  // All generated operations are JSON-safe. A wire round trip must preserve both effect and inverse.
  const envelope = JSON.parse(JSON.stringify(toOpEnvelope(op)));
  exact(envelope.op, op);
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  const fromWire = applyDocumentOp(document, envelope.op);
  if (fromWire.isErr()) throw fromWire.error;
  exact(fromWire.value, result.value);
  exact(document, original);
  invariants(result.value.document);
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  exact(undone.value.document, original);
  exact(
    result.value.document.package.document.content.at(0),
    original.package.document.content.at(0),
  );
  exact(
    result.value.document.package.document.content.at(-1),
    original.package.document.content.at(-1),
  );
  return result.value;
};
const idsForInsert = (table: Table, column: number, first: number) => {
  let count = 0;
  for (const row of table.rows) {
    let start = row.formatting?.gridBefore ?? 0;
    const end = start + row.cells.reduce((sum, cell) => sum + (cell.formatting?.gridSpan ?? 1), 0);
    if (column < start || column > end) continue;
    if (column === start || column === end) {
      count++;
      continue;
    }
    for (const cell of row.cells) {
      start += cell.formatting?.gridSpan ?? 1;
      if (column === start) {
        count++;
        break;
      }
    }
  }
  return Array.from({ length: count }, (_, offset) => id(first + offset));
};
const target = (blockId: string) => ({ story: OP_STORIES.MAIN, blockId });

describe("semantic table edit properties", () => {
  test("final-column intent allocation covers every tracked paragraph and inline span", () => {
    // The general edit sequence keeps at least three columns, so it never exercises
    // the whole-table tracking path reached by deleting the final column.
    const cellContent = fc.array(
      fc.array(
        fc.record({
          text: fc.constantFrom("x", "é😀", "multiple words"),
          bold: fc.boolean(),
          hyperlink: fc.boolean(),
        }),
        { minLength: 0, maxLength: 5 },
      ),
      { minLength: 1, maxLength: 4 },
    );
    assertProperty(
      fc.property(
        fc.array(cellContent, { minLength: 2, maxLength: 4 }),
        fc.boolean(),
        fc.boolean(),
        (rows, nested, explicitGrid) => {
          const f = fixture({
            width: 1,
            height: rows.length,
            variant: 0,
            nested,
            empty: true,
            explicitGrid,
          });
          const table = tableOf(f.document, nested);
          let next = 0x100;
          const paragraphs = rows.map((content) =>
            content.map(
              (spans): Paragraph => ({
                type: "paragraph",
                paraId: id(next++),
                content: spans.map(({ text, bold, hyperlink }) => {
                  const run = {
                    type: "run",
                    formatting: { bold },
                    content: [{ type: "text", text }],
                  } satisfies Run;
                  return hyperlink
                    ? { type: "hyperlink", href: "https://example.org", children: [run] }
                    : run;
                }),
              }),
            ),
          );
          for (const [rowIndex, row] of table.rows.entries()) {
            const cell = row.cells.at(0);
            const content = paragraphs.at(rowIndex);
            if (!cell || !content) throw new Error("Generated cell missing.");
            cell.content = content;
          }
          const intent = {
            type: "table",
            operation: {
              type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
              ...target(targetOf(table)),
              column: 0,
            },
          } as const;
          const allocation = allocateEditorIntentIds(f.document, intent);
          const direct = compileEditorIntent(f.document, {
            intent,
            mode: { type: "editing", newIds: allocation.newIds },
          });
          const tracked = compileEditorIntent(f.document, {
            intent,
            mode: {
              type: "suggesting",
              revision: {
                id: allocation.revisionId,
                author: "Reviewer",
                date: "2026-10-02T10:00:00Z",
              },
              newIds: allocation.newIds,
            },
          });
          if (direct.isErr()) throw direct.error;
          if (tracked.isErr()) throw tracked.error;
          const directOp = direct.value.ops.at(0);
          const trackedOp = tracked.value.ops.at(0);
          if (!directOp || !trackedOp) throw new Error("Compiled deletion missing.");
          const deleted = applied(f.document, directOp);
          const suggested = applied(f.document, trackedOp);
          expect(suggested.revisions.length).toBeGreaterThanOrEqual(rows.length);
          for (const decision of ["accept", "reject"] as const) {
            const resolved = applyDocumentOp(suggested.document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: suggested.revisions,
              decision,
            });
            if (resolved.isErr()) throw resolved.error;
            exact(resolved.value.document, decision === "accept" ? deleted.document : f.document);
          }
        },
      ),
      { numRuns: NUM_RUNS },
    );
  });
  test("generated column edits change logical width and preserve surviving slot content", () => {
    // Wrong implementations that merely change tblGrid pass width checks; surviving-slot identity closes that gap.
    assertProperty(
      fc.property(shape, fc.nat(), (value, seed) => {
        const f = fixture(value);
        const before = tableOf(f.document, f.nested);
        const column = seed % (f.width + 1);
        const inserted = applied(f.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          column,
          width: 900,
          newBlockIds: idsForInsert(before, column, 0x1000),
        });
        const next = tableOf(inserted.document, f.nested);
        const oldSlots = slots(before);
        const newSlots = slots(next);
        for (const [rowIndex, row] of oldSlots.entries()) {
          expect(newSlots.at(rowIndex)?.length).toBe(f.width + 1);
          for (const [oldColumn, cell] of row.entries()) {
            exact(
              newSlots.at(rowIndex)?.at(oldColumn < column ? oldColumn : oldColumn + 1)?.content,
              cell?.content,
            );
          }
        }
        if (before.columnWidths)
          expect(next.columnWidths).toEqual([
            ...before.columnWidths.slice(0, column),
            900,
            ...before.columnWidths.slice(column),
          ]);
        const deleted = applied(inserted.document, {
          ...target(targetOf(next)),
          type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
          column,
        });
        const deletedSlots = slots(tableOf(deleted.document, f.nested));
        expect(deletedSlots.at(0)?.length).toBe(f.width);
        for (const [rowIndex, row] of deletedSlots.entries()) {
          for (const [remainingColumn, cell] of row.entries()) {
            exact(
              cell?.content,
              newSlots
                .at(rowIndex)
                ?.at(remainingColumn < column ? remainingColumn : remainingColumn + 1)?.content,
            );
          }
        }
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("merge and split have independent rectangle, content and topology effects", () => {
    // A merge that drops lower-row content still round trips via snapshots; pin the exact moved block order.
    assertProperty(
      fc.property(shape, (value) => {
        const f = fixture({ ...value, variant: 0 });
        const before = tableOf(f.document, f.nested);
        const merged = applied(f.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.MERGE_CELLS,
          top: 0,
          bottom: f.height,
          left: 0,
          right: 2,
          newBlockIds: Array.from({ length: f.height - 1 }, (_, offset) => id(0x2000 + offset)),
        });
        const table = tableOf(merged.document, f.nested);
        const first = table.rows.at(0)?.cells.at(0);
        exact(
          first?.content,
          before.rows.flatMap((row) => row.cells.slice(0, 2).flatMap((cell) => cell.content)),
        );
        expect(first?.formatting?.gridSpan).toBe(2);
        expect(first?.formatting?.vMerge).toBe("restart");
        for (const row of table.rows.slice(1)) {
          expect(row.cells.at(0)?.formatting?.vMerge).toBe("continue");
          expect(row.cells.length).toBe(f.width - 1);
        }
        // The generator previously coupled every split to a vertical merge, despite
        // that geometry being refused. Cover every group target plus a horizontal success path.
        const mergedSnapshot = structuredClone(merged.document);
        for (const row of table.rows) {
          const mergedParagraph = row.cells.at(0)?.content.at(0);
          if (mergedParagraph?.type !== "paragraph" || !mergedParagraph.paraId)
            throw new Error("Merged target paragraph missing.");
          const refusal = applyDocumentOp(merged.document, {
            ...target(mergedParagraph.paraId),
            type: DOCUMENT_OP_TYPES.SPLIT_CELL,
            newBlockIds: [],
          });
          if (refusal.isOk()) throw new Error("Vertical-group split must refuse.");
          expect(refusal.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
          exact(merged.document, mergedSnapshot);
        }
        const horizontal = fixture({ ...value, variant: 0 });
        const horizontalBefore = tableOf(horizontal.document, horizontal.nested);
        const horizontalMerged = applied(horizontal.document, {
          ...target(horizontal.target),
          type: DOCUMENT_OP_TYPES.MERGE_CELLS,
          top: 0,
          bottom: 1,
          left: 0,
          right: 2,
          newBlockIds: [],
        });
        const horizontalTable = tableOf(horizontalMerged.document, horizontal.nested);
        const topContent = horizontalBefore.rows
          .at(0)
          ?.cells.slice(0, 2)
          .flatMap((cell) => cell.content);
        exact(horizontalTable.rows.at(0)?.cells.at(0)?.content, topContent);
        expect(horizontalTable.rows.at(0)?.cells.at(0)?.formatting?.vMerge).toBeUndefined();
        const split = applied(horizontalMerged.document, {
          ...target(horizontal.target),
          type: DOCUMENT_OP_TYPES.SPLIT_CELL,
          newBlockIds: [id(0x3000)],
        });
        const splitTable = tableOf(split.document, horizontal.nested);
        const topRow = splitTable.rows.at(0);
        expect(topRow?.cells.length).toBe(horizontal.width);
        exact(topRow?.cells.at(0)?.content, topContent);
        exact(topRow?.cells.at(1)?.content, [
          { type: "paragraph", paraId: id(0x3000), content: [] },
        ]);
        for (const cell of topRow?.cells ?? []) {
          expect(cell.formatting?.gridSpan ?? 1).toBe(1);
          expect(cell.formatting?.vMerge).toBeUndefined();
        }
        for (let rowIndex = 1; rowIndex < horizontal.height; rowIndex++)
          exact(splitTable.rows.at(rowIndex), horizontalBefore.rows.at(rowIndex));
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("split generated horizontal and vertical spans preserves each row's original content", () => {
    // A splitter that handles only the merge operation's output misses existing OOXML vertical groups.
    assertProperty(
      fc.property(shape, fc.boolean(), (value, vertical) => {
        const f = fixture({ ...value, variant: vertical ? 2 : 1 });
        const before = tableOf(f.document, f.nested);
        const splitHeight = vertical ? f.height : 1;
        const op = {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SPLIT_CELL,
          newBlockIds: Array.from({ length: splitHeight }, (_, offset) => id(0x7000 + offset)),
        } satisfies DocumentOp;
        if (vertical) {
          const refused = applyDocumentOp(f.document, op);
          expect(refused.isErr() && refused.error.reason).toBe(
            DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          );
          exact(tableOf(f.document, f.nested), before);
          return;
        }
        const result = applied(f.document, op);
        const after = tableOf(result.document, f.nested);
        for (let rowIndex = 0; rowIndex < splitHeight; rowIndex++) {
          const row = after.rows.at(rowIndex);
          expect(row?.cells.length).toBe(f.width);
          exact(row?.cells.at(0)?.content, before.rows.at(rowIndex)?.cells.at(0)?.content);
          expect(row?.cells.at(0)?.formatting?.gridSpan ?? 1).toBe(1);
          expect(row?.cells.at(0)?.formatting?.vMerge).toBeUndefined();
          expect(row?.cells.at(1)?.formatting?.vMerge).toBeUndefined();
        }
        for (let rowIndex = splitHeight; rowIndex < f.height; rowIndex++)
          exact(after.rows.at(rowIndex), before.rows.at(rowIndex));
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("generated mixed sequences reverse exactly and every property patch takes effect", () => {
    // Checking only final identity allows skipped patches. Inspect each edited field before accumulating inverses.
    assertProperty(
      fc.property(
        shape,
        fc.array(fc.integer({ min: 0, max: 4 }), { minLength: 3, maxLength: 12 }),
        (value, sequence) => {
          const f = fixture(value);
          let document = f.document;
          const inverseBatches: (readonly DocumentOp[])[] = [];
          for (const [step, kind] of sequence.entries()) {
            const before = tableOf(document, f.nested);
            const address = target(targetOf(before));
            let op: DocumentOp;
            switch (kind) {
              case 0:
                op = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
                  patch: { verticalAlign: step % 2 ? "bottom" : "center", noWrap: true },
                };
                break;
              case 1:
                op = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
                  patch: { cantSplit: true, heightRule: "exact" },
                };
                break;
              case 2:
                op = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
                  patch: { justification: "right", styleId: `style-${step}` },
                };
                break;
              case 3:
                op = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
                  columnWidths: Array.from(
                    { length: slots(before).at(0)?.length ?? 0 },
                    () => 1000 + step,
                  ),
                };
                break;
              default:
                op = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
                  column: 0,
                  width: 600,
                  newBlockIds: idsForInsert(before, 0, 0x4000 + step * 16),
                };
            }
            const result = applied(document, op);
            const after = tableOf(result.document, f.nested);
            switch (op.type) {
              case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
                expect(after.rows.at(0)?.cells.at(0)?.formatting?.verticalAlign).toBe(
                  op.patch.verticalAlign,
                );
                expect(after.rows.at(0)?.cells.at(0)?.formatting?.noWrap).toBe(true);
                break;
              case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
                expect(after.rows.at(0)?.formatting?.cantSplit).toBe(true);
                expect(after.rows.at(0)?.formatting?.heightRule).toBe("exact");
                break;
              case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
                expect(after.formatting?.justification).toBe("right");
                expect(after.formatting?.styleId).toBe(`style-${step}`);
                break;
              case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
                exact(after.columnWidths, op.columnWidths);
                break;
              case DOCUMENT_OP_TYPES.INSERT_COLUMN:
                expect(slots(after).at(0)?.length).toBe((slots(before).at(0)?.length ?? 0) + 1);
                break;
            }
            inverseBatches.push(result.inverse);
            document = result.document;
          }
          const undone = applyDocumentOps(document, inverseBatches.reverse().flat());
          if (undone.isErr()) throw undone.error;
          exact(undone.value.document, f.document);
        },
      ),
      { numRuns: NUM_RUNS },
    );
  });

  test("null patches restore missing, empty and own undefined formatting exactly", () => {
    // A JSON snapshot inverse loses undefined; these deliberately non-wire documents exercise own-key restoration.
    assertProperty(
      fc.property(shape, fc.integer({ min: 0, max: 2 }), (value, spelling) => {
        const f = fixture(value);
        const table = tableOf(f.document, f.nested);
        const row = table.rows.at(0);
        const cell = row?.cells.at(0);
        if (!row || !cell) throw new Error("Fixture cell missing.");
        if (spelling === 0) {
          Reflect.deleteProperty(table, "formatting");
        } else if (spelling === 1) {
          table.formatting = {};
        } else {
          Reflect.set(table, "formatting", undefined);
        }
        const set = applied(f.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
          patch: { styleId: "temporary", layout: "autofit", justification: "left" },
        });
        const cleared = applied(set.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
          patch: { styleId: null, layout: null, justification: null },
        });
        expect(tableOf(cleared.document, f.nested).formatting?.styleId).toBeUndefined();
        const cellSet = applied(f.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
          patch: { fitText: true, hideMark: true, verticalAlign: "bottom" },
        });
        const cellCleared = applied(cellSet.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
          patch: { fitText: null, hideMark: null, verticalAlign: null },
        });
        expect(
          tableOf(cellCleared.document, f.nested).rows.at(0)?.cells.at(0)?.formatting?.fitText,
        ).toBeUndefined();
        const rowSet = applied(f.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
          patch: { cantSplit: true, hidden: true },
        });
        const rowCleared = applied(rowSet.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
          patch: { cantSplit: null, hidden: null },
        });
        expect(
          tableOf(rowCleared.document, f.nested).rows.at(0)?.formatting?.hidden,
        ).toBeUndefined();
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("generated tracked compiler sequences accept to direct edits and reject to the exact baseline", () => {
    // An inverse-only oracle misses broken review semantics and wrong compiler selections. Resolve each step,
    // then continue on its accepted document: generated sequences exercise changed grids and fresh identities.
    assertProperty(
      fc.property(
        shape,
        fc.array(fc.integer({ min: 0, max: 7 }), { minLength: 2, maxLength: 8 }),
        (value, sequence) => {
          const f = fixture({ ...value, explicitGrid: true });
          let document = f.document;
          for (const [step, kind] of sequence.entries()) {
            const before = tableOf(document, f.nested);
            const address = target(targetOf(before));
            const width = slots(before).at(0)?.length ?? 0;
            const column = step % width;
            const firstSpan = before.rows.at(0)?.cells.at(0)?.formatting?.gridSpan ?? 1;
            const groupHeight =
              before.rows.at(0)?.cells.at(0)?.formatting?.vMerge === "restart"
                ? before.rows.filter((row) => row.cells.at(0)?.formatting?.vMerge !== undefined)
                    .length
                : 1;
            const trailingInsert = () =>
              ({
                ...address,
                type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
                column: width,
                width: 950,
              }) as const satisfies TableIntentOperation;
            let operation: TableIntentOperation;
            switch (kind) {
              case 0:
                operation = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
                  patch: { noWrap: step % 2 === 0, fitText: true },
                };
                break;
              case 1:
                operation = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
                  patch: { cantSplit: true, hidden: step % 2 === 0 },
                };
                break;
              case 2:
                operation = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
                  patch: { styleId: `sequence-${step}` },
                };
                break;
              case 3:
                operation = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
                  columnWidths: Array.from({ length: width }, () => 1100 + step),
                };
                break;
              case 4:
                operation = {
                  ...address,
                  type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
                  column,
                  width: 850,
                };
                break;
              case 5:
                operation =
                  width > 3 &&
                  before.rows.every(
                    (row) =>
                      (row.formatting?.gridBefore ?? 0) > 0 ||
                      row.cells.length > 1 ||
                      (row.cells.at(0)?.formatting?.gridSpan ?? 1) > 1,
                  )
                    ? { ...address, type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column: 0 }
                    : trailingInsert();
                break;
              case 6:
                operation =
                  groupHeight > 1 || (firstSpan === 1 && groupHeight === 1)
                    ? trailingInsert()
                    : { ...address, type: DOCUMENT_OP_TYPES.SPLIT_CELL };
                break;
              default:
                operation = trailingInsert();
            }
            const intent = { type: "table", operation } as const;
            const allocation = allocateEditorIntentIds(document, intent);
            const revision = {
              id: allocation.revisionId,
              author: "Reviewer",
              date: "2026-10-02T10:00:00Z",
            };
            const editing = compileEditorIntent(document, {
              intent,
              mode: { type: "editing", newIds: allocation.newIds },
            });
            const suggesting = compileEditorIntent(document, {
              intent,
              mode: {
                type: "suggesting",
                revision,
                newIds: allocation.newIds,
              },
            });
            if (editing.isErr()) throw editing.error;
            if (suggesting.isErr()) throw suggesting.error;
            expect(editing.value.ops.length).toBe(1);
            expect(suggesting.value.ops.length).toBe(1);
            const directOp = editing.value.ops.at(0);
            const trackedOp = suggesting.value.ops.at(0);
            if (!directOp || !trackedOp) throw new Error("Compiler must emit one table operation.");
            const direct = applied(document, directOp);
            const tracked = applied(document, trackedOp);
            if ("newIds" in directOp) {
              expect(directOp.newIds?.revision ?? []).toEqual([]);
              expect(directOp.newIds?.control ?? []).toEqual([]);
            }
            if ("newIds" in trackedOp) {
              expect((trackedOp.newIds?.revision ?? []).toSorted((a, b) => a - b)).toEqual(
                tracked.revisions
                  .filter((revisionId) => revisionId !== revision.id)
                  .toSorted((a, b) => a - b),
              );
              expect(trackedOp.newIds?.control ?? []).toEqual([]);
            }
            // Physical model serialization is an additional oracle, separate from operation-envelope serialization.
            exact(JSON.parse(JSON.stringify(document)), document);
            exact(JSON.parse(JSON.stringify(tracked.document)), tracked.document);
            invariants(JSON.parse(JSON.stringify(tracked.document)));
            const liveIds: string[] = [];
            const collect = (blocks: readonly BlockContent[]) => {
              for (const block of blocks) {
                if (block.type === "paragraph" && block.paraId) liveIds.push(block.paraId);
                else if (block.type === "table")
                  for (const row of block.rows) for (const cell of row.cells) collect(cell.content);
              }
            };
            collect(direct.document.package.document.content);
            expect(liveIds.includes(editing.value.selection.blockId)).toBe(true);
            expect(editing.value.selection.offset).toBe(0);
            for (const decision of ["accept", "reject"] as const) {
              if ((tracked.revisions?.length ?? 0) === 0) {
                exact(tracked.document, direct.document);
                exact(tracked.document, document);
              } else {
                const resolved = applyDocumentOp(tracked.document, {
                  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                  story: OP_STORIES.MAIN,
                  revisionIds: tracked.revisions ?? [],
                  decision,
                });
                if (resolved.isErr()) throw resolved.error;
                exact(resolved.value.document, decision === "accept" ? direct.document : document);
                invariants(resolved.value.document);
              }
            }
            document = direct.document;
          }
        },
      ),
      { numRuns: NUM_RUNS },
    );
  });

  test("deleting the only column removes its table, restores exactly and selects a surviving sibling", () => {
    // A width check misses an invalid empty table; assert exact removal from the owning block list.
    assertProperty(
      fc.property(shape, (value) => {
        const f = fixture({ ...value, variant: 0, explicitGrid: true });
        const before = tableOf(f.document, f.nested);
        before.columnWidths = [700];
        for (const row of before.rows) row.cells = row.cells.slice(0, 1);
        const expected = structuredClone(f.document);
        if (f.nested) {
          const outer = expected.package.document.content.at(1);
          if (outer?.type !== "table") throw new Error("Fixture outer table missing.");
          const carrier = outer.rows.at(0)?.cells.at(0);
          if (!carrier) throw new Error("Fixture carrier missing.");
          carrier.content = carrier.content.slice(1);
        } else {
          expected.package.document.content.splice(1, 1);
        }
        const operation = {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
          column: 0,
        } as const;
        const intent = { type: "table", operation } as const;
        const editing = compileEditorIntent(f.document, { intent, mode: { type: "editing" } });
        const suggesting = compileEditorIntent(f.document, {
          intent,
          mode: {
            type: "suggesting",
            revision: { id: 0x70000, author: "Reviewer", date: "2026-10-02T10:00:00Z" },
            newIds: { revision: Array.from({ length: 64 }, (_, offset) => 0x70001 + offset) },
          },
        });
        if (editing.isErr()) throw editing.error;
        if (suggesting.isErr()) throw suggesting.error;
        const directOp = editing.value.ops.at(0);
        const trackedOp = suggesting.value.ops.at(0);
        if (!directOp || !trackedOp) throw new Error("Compiler must emit column deletion.");
        const direct = applied(f.document, directOp);
        const tracked = applied(f.document, trackedOp);
        exact(direct.document, expected);
        const survivingId = f.nested ? id(0x20) : id(1);
        expect(editing.value.selection.blockId).toBe(survivingId);
        expect(suggesting.value.selection.blockId).toBe(survivingId);
        expect(editing.value.selection.offset).toBe(0);
        expect(suggesting.value.selection.offset).toBe(0);
        for (const decision of ["accept", "reject"] as const) {
          const result = applyDocumentOp(tracked.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: tracked.revisions ?? [],
            decision,
          });
          if (result.isErr()) throw result.error;
          exact(result.value.document, decision === "accept" ? expected : f.document);
          invariants(result.value.document);
        }
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("empty vertical merge tracking keeps cell identities and compiles in both modes", () => {
    // A tracked merge that replaces continuation paragraphs can still look correct; pin every physical id.
    assertProperty(
      fc.property(shape, (value) => {
        const f = fixture({ ...value, variant: 0, empty: true, explicitGrid: true });
        const before = tableOf(f.document, f.nested);
        const operation = {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.MERGE_CELLS,
          top: 0,
          bottom: f.height,
          left: 0,
          right: 1,
          newBlockIds: [],
        } as const;
        const editing = compileEditorIntent(f.document, {
          intent: { type: "table", operation },
          mode: { type: "editing" },
        });
        const suggesting = compileEditorIntent(f.document, {
          intent: { type: "table", operation },
          mode: {
            type: "suggesting",
            revision: { id: 0x60000, author: "Reviewer", date: "2026-10-02T10:00:00Z" },
            newIds: { revision: Array.from({ length: 24 }, (_, offset) => 0x60001 + offset) },
          },
        });
        if (editing.isErr()) throw editing.error;
        if (suggesting.isErr()) throw suggesting.error;
        const directOp = editing.value.ops.at(0);
        const trackedOp = suggesting.value.ops.at(0);
        if (!directOp || !trackedOp) throw new Error("Compiler must emit table merge.");
        const direct = applied(f.document, directOp);
        const tracked = applied(f.document, trackedOp);
        const table = tableOf(tracked.document, f.nested);
        for (const [rowIndex, row] of table.rows.entries()) {
          exact(row.cells.at(0)?.content, before.rows.at(rowIndex)?.cells.at(0)?.content);
        }
        for (const decision of ["accept", "reject"] as const) {
          const result = applyDocumentOp(tracked.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: tracked.revisions ?? [],
            decision,
          });
          if (result.isErr()) throw result.error;
          exact(result.value.document, decision === "accept" ? direct.document : f.document);
        }
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("fresh-id counts and rectangles cutting spans are refused before mutation", () => {
    // Geometry-only fixtures miss id allocation failures; exact and surplus counts exercise the allocation boundary.
    assertProperty(
      fc.property(shape, (value) => {
        const f = fixture({ ...value, variant: 2 });
        const address = target(f.target);
        const cases = [
          {
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.SPLIT_CELL,
              newBlockIds: [],
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          },
          {
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.MERGE_CELLS,
              top: 0,
              bottom: f.height,
              left: 1,
              right: 2,
              newBlockIds: [],
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          },
          {
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.MERGE_CELLS,
              top: 0,
              bottom: 1,
              left: 0,
              right: 2,
              newBlockIds: [],
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          },
          {
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
              patch: { gridSpan: 3 },
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          },
          {
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
              patch: { gridBefore: 1 },
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          },
        ];
        const snapshot = structuredClone(f.document);
        for (const { op, reason } of cases) {
          const result = applyDocumentOp(f.document, op);
          if (result.isOk()) throw new Error(`Expected ${reason} for ${op.type}.`);
          expect(result.error.reason).toBe(reason);
          exact(f.document, snapshot);
        }
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("tracked property edits retain previous formatting and exact inverses", () => {
    // Merely accepting a revision field is insufficient: require the typed OOXML property-change record.
    assertProperty(
      fc.property(shape, (value) => {
        const f = fixture(value);
        const before = tableOf(f.document, f.nested);
        const revision = { id: 0x12000, author: "Reviewer", date: "2026-10-02T10:00:00Z" };
        const cell = applied(f.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
          patch: { noWrap: true },
          revision,
        });
        const cellChange = tableOf(cell.document, f.nested)
          .rows.at(0)
          ?.cells.at(0)
          ?.propertyChanges?.at(0);
        expect(cellChange?.type).toBe("tableCellPropertyChange");
        expect(cellChange?.info.id).toBe(revision.id);
        exact(cellChange?.previousFormatting, before.rows.at(0)?.cells.at(0)?.formatting);
        const row = applied(f.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_ROW_PROPS,
          patch: { cantSplit: true },
          revision,
        });
        const rowChange = tableOf(row.document, f.nested).rows.at(0)?.propertyChanges?.at(0);
        expect(rowChange?.type).toBe("tableRowPropertyChange");
        exact(rowChange?.previousFormatting, before.rows.at(0)?.formatting);
        const table = applied(f.document, {
          ...target(f.target),
          type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
          patch: { styleId: "tracked" },
          revision,
        });
        const tableChange = tableOf(table.document, f.nested).propertyChanges?.at(0);
        expect(tableChange?.type).toBe("tablePropertyChange");
        exact(tableChange?.previousFormatting, before.formatting);
      }),
      { numRuns: NUM_RUNS },
    );
  });

  test("invalid geometry, colliding identities, tracked edits and stale inverses refuse atomically", () => {
    // Silent continue would hide a universally refusing implementation. Positive laws above require success;
    // here the complete declared refusal set is counted and each generated case must refuse exactly once.
    const tallies = new Map<string, number>();
    let attempts = 0;
    assertProperty(
      fc.property(shape, (value) => {
        const f = fixture({ ...value, variant: 0 });
        const address = target(f.target);
        const inserted = applied(f.document, {
          ...address,
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          column: 0,
          width: 800,
          newBlockIds: Array.from({ length: f.height }, (_, offset) => id(0x6000 + offset)),
        });
        const changed = applied(inserted.document, {
          ...target(targetOf(tableOf(inserted.document, f.nested))),
          type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
          patch: { styleId: "concurrent" },
        });
        const inverse = inserted.inverse.at(0);
        if (!inverse) throw new Error("Edit must have an inverse.");
        const cases = [
          {
            document: f.document,
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
              column: f.width,
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          },
          {
            document: f.document,
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.SET_TABLE_GRID,
              columnWidths: [1],
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          },
          {
            document: f.document,
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.MERGE_CELLS,
              top: 1,
              bottom: 0,
              left: 0,
              right: 2,
              newBlockIds: [],
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
          },
          {
            document: f.document,
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
              column: 0,
              width: 800,
              newBlockIds: Array.from({ length: f.height }, () => f.target),
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
          },
          {
            document: f.document,
            op: {
              ...address,
              type: DOCUMENT_OP_TYPES.MERGE_CELLS,
              top: 0,
              bottom: 1,
              left: 0,
              right: 2,
              newBlockIds: [],
              revision: { id: 8000, author: "Reviewer", date: "2026-10-02T10:00:00Z" },
            } satisfies DocumentOp,
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
          },
          { document: changed.document, op: inverse, reason: DOCUMENT_OP_REFUSAL_REASONS.STALE },
        ];
        for (const { document, op, reason } of cases) {
          const snapshot = structuredClone(document);
          const result = applyDocumentOp(document, op);
          attempts++;
          if (result.isOk()) throw new Error(`Expected ${reason} for ${op.type}.`);
          expect(result.error.reason).toBe(reason);
          tallies.set(reason, (tallies.get(reason) ?? 0) + 1);
          exact(document, snapshot);
        }
      }),
      { numRuns: NUM_RUNS },
    );
    expect([...tallies.keys()].sort()).toEqual(
      [
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        DOCUMENT_OP_REFUSAL_REASONS.STALE,
      ].sort(),
    );
    expect([...tallies.values()].reduce((sum, count) => sum + count, 0)).toBe(attempts);
    // Each reason has a fixed maximum share, making reason drift visible rather than skipped.
    expect(tallies.get(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH)).toBe(attempts / 2);
    for (const reason of [
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
    ]) {
      expect(tallies.get(reason)).toBe(attempts / 6);
    }
  });
});
