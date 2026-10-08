import assert from "node:assert/strict";
import { describe, test } from "node:test";

import {
  assertRequestedOutcome,
  capture,
  compareComments,
  compareWithModel,
  expectOperation,
  modelOf,
  type Row,
} from "../support/oracle.ts";
import { sequentialGroups } from "../support/metamorphic.ts";
import { FIXTURES, openReviewer } from "../support/documents.ts";
import { coreBatch, GENERATORS, MODES, supports } from "../support/operations.ts";
import { createRandom } from "../support/random.ts";

type CellLocation = { rowIndex: number; column: number; columnSpan?: number; tableIndex?: number };

const row = (
  id: string,
  text: string,
  { rowIndex, column, columnSpan = 1, tableIndex = 0 }: CellLocation,
): Row => ({
  id,
  text,
  kind: "paragraph",
  table: {
    outerTableIndex: tableIndex,
    tableIndex,
    rowIndex,
    cellIndex: column,
    gridColumnIndex: column,
    columnSpan,
    rowSpan: 1,
  },
});

describe("table operation source coordinates", () => {
  test("seed 201 removes table-owned text and comments without removing an outside join", async () => {
    const reviewer = await openReviewer(await FIXTURES.tables());
    const random = createRandom(201);
    for (const type of Object.keys(GENERATORS)) {
      if (!supports(type, "tracked-changes")) continue;
      const operation = GENERATORS[type]?.(reviewer.getContent(), random);
      if (!operation) continue;
      const pre = await capture(reviewer, "tracked-changes");
      const result = reviewer.applyDocumentOperations(
        coreBatch([{ ...operation, id: `op-${type}` }], "tracked-changes") as never,
      );
      if (type !== "deleteTable") continue;
      assert.equal(result.applied.length, 1);
      const acceptedIds = new Set(pre.rows.map(({ id }) => id));
      const outside = pre.liveRows.filter(
        ({ id, text, table }) => !acceptedIds.has(id) && text.length > 0 && table === undefined,
      );
      assert.equal(outside.length, 1, "the pinned generator must reach the outside join");
      assert.ok(pre.comments.length >= 2, "the pinned generator must reach table comments");
      await assertRequestedOutcome(
        reviewer,
        pre,
        { applied: [operation] },
        "table ownership seed 201",
      );
      return;
    }
    assert.fail("seed 201 did not reach deleteTable");
  });

  test("table removals retain outside join sources and remove only owned cell comments", () => {
    for (let width = 2; width <= 4; width++) {
      for (let sourceCount = 0; sourceCount <= 3; sourceCount++) {
        for (const type of ["deleteTable", "deleteTableRow", "deleteTableColumn"] as const) {
          const outside = Array.from(
            { length: sourceCount },
            (_, index): Row => ({
              id: `outside-${index}`,
              text: "Shared ",
              kind: "paragraph",
            }),
          );
          const cells = Array.from({ length: width * 2 }, (_, index) =>
            row(
              `r${Math.floor(index / width)}c${index % width}`,
              index === 0 ? "Shared" : `Cell ${index}`,
              { rowIndex: Math.floor(index / width), column: index % width },
            ),
          );
          const first = cells.at(0);
          assert.ok(first);
          const tail: Row = { id: "tail", text: "Following paragraph.", kind: "paragraph" };
          const prefix = outside.map(({ text }) => text).join("");
          const accepted = [{ ...first, text: prefix + first.text }, ...cells.slice(1), tail];
          const model = modelOf(accepted, [...outside, ...cells, tail]);
          model.mode = "tracked-changes";
          // Either the joined destination or a sibling cell may anchor the removal.
          const anchor = type === "deleteTable" ? cells.at(-1) : first;
          assert.ok(anchor);
          expectOperation(model, { type, blockId: anchor.id });
          const keptCells = cells.filter((cell) => {
            assert.ok(cell.table);
            switch (type) {
              case "deleteTable":
                return false;
              case "deleteTableRow":
                return cell.table.rowIndex !== 0;
              case "deleteTableColumn":
                return cell.table.gridColumnIndex !== 0;
              default: {
                const exhaustive: never = type;
                return exhaustive;
              }
            }
          });
          const following = [...keptCells, tail];
          const next = following.at(0);
          assert.ok(next);
          const expected = [{ ...next, text: prefix + next.text }, ...following.slice(1)];
          assert.deepEqual(
            compareWithModel(model, expected),
            [],
            `${type}, width ${width}, sources ${sourceCount}`,
          );
          if (sourceCount > 0) {
            assert.ok(
              compareWithModel(model, following).some((problem) =>
                problem.startsWith("block texts differ"),
              ),
            );
          }
          const cellComment = { id: 1, text: "Cell comment", anchor: "Shared", blockId: first.id };
          const sourceComment = {
            id: 2,
            text: "Outside comment",
            anchor: "Shared",
            blockId: first.id,
          };
          const outsideFirst = outside.at(0);
          const before = outsideFirst ? [cellComment, sourceComment] : [cellComment];
          const liveBefore = outsideFirst
            ? [cellComment, { ...sourceComment, blockId: outsideFirst.id }]
            : before;
          const after = outsideFirst ? [{ ...sourceComment, blockId: next.id }] : [];
          assert.deepEqual(
            compareComments(model, before, after, liveBefore, "tracked-changes"),
            [],
          );
          assert.ok(
            compareComments(
              model,
              before,
              [...after, cellComment],
              liveBefore,
              "tracked-changes",
            ).some((problem) => problem.includes("outlived the block")),
          );
        }
      }
    }
  });

  test("successive column removals carry an outside join across each removed destination", () => {
    const outside: Row = { id: "outside", text: "Prefix ", kind: "paragraph" };
    const cells = Array.from({ length: 3 }, (_, column) =>
      row(`c${column}`, `Cell ${column}`, { rowIndex: 0, column }),
    );
    const first = cells.at(0);
    const last = cells.at(-1);
    assert.ok(first && last);
    const model = modelOf(
      [{ ...first, text: outside.text + first.text }, ...cells.slice(1)],
      [outside, ...cells],
    );
    for (const blockId of ["c0", "c1"])
      expectOperation(model, { type: "deleteTableColumn", blockId });
    assert.deepEqual(compareWithModel(model, [{ ...last, text: outside.text + last.text }]), []);
  });

  test("column removal then table removal carries the outside prefix exactly once", () => {
    const outside: Row = { id: "12345678", text: "Prefix ", kind: "paragraph" };
    const first = row("23456789", "Cell one", { rowIndex: 0, column: 0 });
    const second = row("34567890", "Cell two", { rowIndex: 0, column: 1 });
    const tail: Row = { id: "45678901", text: "Tail.", kind: "paragraph" };
    const model = modelOf(
      [{ ...first, text: outside.text + first.text }, second, tail],
      [outside, first, second, tail],
    );
    const firstModelRow = model.rows.at(0);
    assert.ok(firstModelRow);
    assert.equal(model.pendingJoins.get(firstModelRow)?.sources.at(0)?.id, outside.id);
    expectOperation(model, { type: "deleteTableColumn", blockId: first.id });
    assert.deepEqual(compareWithModel(model, [{ ...second, text: "Prefix Cell two" }, tail]), []);
    expectOperation(model, { type: "deleteTable", blockId: second.id });
    assert.deepEqual(compareWithModel(model, [{ ...tail, text: "Prefix Tail." }]), []);
    assert.ok(compareWithModel(model, [{ ...tail, text: "Prefix Prefix Tail." }]).length > 0);
  });

  test("multi-operation table removals consume each outside join contribution once", () => {
    const sequences = [
      ["deleteTableColumn"],
      ["deleteTableRow"],
      ["deleteTableColumn", "deleteTableRow"],
      ["deleteTableRow", "deleteTableColumn"],
      ["deleteTableColumn", "deleteTableColumn"],
      ["deleteTableRow", "deleteTableRow"],
    ] as const;
    for (let width = 3; width <= 5; width++) {
      for (let height = 3; height <= 5; height++) {
        for (let sourceCount = 1; sourceCount <= 3; sourceCount++) {
          for (const sequence of sequences) {
            const sources = Array.from(
              { length: sourceCount },
              (_, index): Row => ({
                id: `outside-${index}`,
                text: `Source ${index}. `,
                kind: "paragraph",
              }),
            );
            const outsideSource = sources.at(0);
            assert.ok(outsideSource);
            const prefix = sources.map(({ text }) => text).join("");
            const cells = Array.from({ length: width * height }, (_, index) =>
              row(`r${Math.floor(index / width)}c${index % width}`, `Cell ${index}`, {
                rowIndex: Math.floor(index / width),
                column: index % width,
              }),
            );
            const first = cells.at(0);
            assert.ok(first);
            const tail: Row = { id: "tail", text: "Tail.", kind: "paragraph" };
            const model = modelOf(
              [{ ...first, text: prefix + first.text }, ...cells.slice(1), tail],
              [...sources, ...cells, tail],
            );
            let kept = cells;
            for (const type of sequence) {
              const anchor = kept.at(0);
              assert.ok(anchor?.table);
              expectOperation(model, { type, blockId: anchor.id });
              const location = anchor.table;
              kept = kept.filter(({ table }) =>
                type === "deleteTableColumn"
                  ? table?.gridColumnIndex !== location.gridColumnIndex
                  : table?.rowIndex !== location.rowIndex,
              );
              const next = kept.at(0);
              assert.ok(next);
              assert.deepEqual(
                compareWithModel(model, [
                  { ...next, text: prefix + next.text },
                  ...kept.slice(1),
                  tail,
                ]),
                [],
              );
            }
            const anchor = kept.at(0);
            assert.ok(anchor);
            expectOperation(model, { type: "deleteTable", blockId: anchor.id });
            assert.deepEqual(compareWithModel(model, [{ ...tail, text: prefix + tail.text }]), []);
            assert.equal(model.pendingJoins.size, 0);
            assert.deepEqual(
              [...model.preservedJoinPrefixes.keys()],
              [model.rows.find(({ pre }) => pre?.id === tail.id)],
            );
            const cellComments = cells.map(({ id, text }, index) => ({
              id: index + 1,
              text: "Cell comment",
              anchor: text,
              blockId: id,
            }));
            const outsideComment = {
              id: cells.length + 1,
              text: "Outside comment",
              anchor: prefix,
              blockId: first.id,
            };
            assert.deepEqual(
              compareComments(
                model,
                [...cellComments, outsideComment],
                [{ ...outsideComment, blockId: tail.id }],
                [...cellComments, { ...outsideComment, blockId: outsideSource.id }],
                "tracked-changes",
              ),
              [],
            );
          }
        }
      }
    }
  });

  test("removing a final table retains the outside paragraph when no later destination exists", () => {
    for (let sourceCount = 1; sourceCount <= 3; sourceCount++) {
      const sources = Array.from(
        { length: sourceCount },
        (_, index): Row => ({
          id: `outside-${index}`,
          text: `Source ${index}. `,
          kind: "paragraph",
        }),
      );
      const cell = row("cell", "Owned cell", { rowIndex: 0, column: 0 });
      const prefix = sources.map(({ text }) => text).join("");
      const model = modelOf([{ ...cell, text: prefix + cell.text }], [...sources, cell]);
      expectOperation(model, { type: "deleteTable", blockId: cell.id });
      assert.deepEqual(
        compareWithModel(model, [{ id: "retained", text: prefix, kind: "paragraph" }]),
        [],
      );
    }
  });

  test("ownership reconstruction uses pre-state facts after earlier modeled edits", () => {
    const source: Row = { id: "outside", text: "Prefix ", kind: "paragraph" };
    const cell = row("cell", "Owned cell", { rowIndex: 0, column: 0 });
    const tail: Row = { id: "tail", text: "Tail.", kind: "paragraph" };
    const model = modelOf([{ ...cell, text: source.text + cell.text }, tail], [source, cell, tail]);
    expectOperation(model, {
      type: "replaceInBlock",
      blockId: tail.id,
      find: "Tail",
      replace: "Updated tail",
    });
    expectOperation(model, { type: "deleteTable", blockId: cell.id });
    assert.deepEqual(compareWithModel(model, [{ ...tail, text: "Prefix Updated tail." }]), []);
    const malformed = modelOf([{ ...cell, text: "Unowned text" }, tail], [source, cell, tail]);
    assert.throws(
      () => expectOperation(malformed, { type: "deleteTable", blockId: cell.id }),
      /does not reconstruct/u,
    );
  });

  test("replays row payloads before columns in source-grid execution order", () => {
    for (let width = 3; width <= 5; width += 1) {
      const rows = Array.from({ length: 2 }, (_, rowIndex) => {
        const cells: Row[] = [];
        for (let column = 0; column < width;) {
          const columnSpan = rowIndex === 0 && column === 0 ? 2 : 1;
          cells.push(row(`t0r${rowIndex}c${column}`, "cell", { rowIndex, column, columnSpan }));
          column += columnSpan;
        }
        return cells;
      }).flat();
      rows.push(
        ...Array.from({ length: width }, (_, column) =>
          row(`t1c${column}`, "cell", { rowIndex: 0, column, tableIndex: 1 }),
        ),
      );
      const operations = [
        { id: "delete-col-0", type: "deleteTableColumn", blockId: "t0r0c0" },
        { id: "insert-col-last", type: "insertTableColumn", blockId: `t0r0c${width - 1}` },
        {
          id: "row",
          type: "insertTableRow",
          blockId: "t0r1c0",
          cellTexts: Array.from({ length: width }, (_, index) => `cell ${index}`),
        },
        { id: "delete-col-2", type: "deleteTableColumn", blockId: "t0r0c2" },
        { id: "insert-col-tie", type: "insertTableColumn", blockId: "t0r0c0" },
        { id: "insert-other-table", type: "insertTableColumn", blockId: "t1c0", position: "after" },
      ];
      for (const mode of MODES) {
        const ordered = sequentialGroups({ applied: operations, preRows: rows, mode }).flatMap(
          (group) => group,
        );
        assert.deepEqual(
          ordered.map(({ id }) => id),
          [
            "row",
            "insert-other-table",
            "insert-col-last",
            "insert-col-tie",
            "delete-col-2",
            "delete-col-0",
          ],
          `${mode} / table width ${width}`,
        );
      }
    }
  });

  test("places a row beside a pending deleted row in live row order", () => {
    for (let rowCount = 2; rowCount <= 5; rowCount += 1) {
      for (let deletedRow = 0; deletedRow < rowCount; deletedRow += 1) {
        const live = Array.from({ length: rowCount }, (_, rowIndex) => [
          row(`r${rowIndex}c0`, `Row ${rowIndex} cell 1`, { rowIndex, column: 0 }),
          row(`r${rowIndex}c1`, `Row ${rowIndex} cell 2`, { rowIndex, column: 1 }),
        ]).flat();
        const accepted = live
          .filter((cell) => cell.table?.rowIndex !== deletedRow)
          .map((cell) => {
            assert.ok(cell.table);
            return row(cell.id, cell.text, {
              rowIndex: cell.table.rowIndex - Number(cell.table.rowIndex > deletedRow),
              column: cell.table.gridColumnIndex,
            });
          });
        const pending = live.map((cell) => {
          assert.ok(cell.table);
          return cell.table.rowIndex === deletedRow ? Object.assign({}, cell, { text: "" }) : cell;
        });
        const model = modelOf(accepted, pending);
        model.mode = "tracked-changes";
        expectOperation(model, {
          type: "insertTableRow",
          blockId: `r${deletedRow}c0`,
          position: "after",
          cellTexts: ["New cell 1", "New cell 2"],
        });

        const expected: Row[] = [];
        for (let sourceRow = 0; sourceRow < rowCount; sourceRow += 1) {
          if (sourceRow === deletedRow) {
            expected.push(
              row("newc0", "New cell 1", { rowIndex: 0, column: 0 }),
              row("newc1", "New cell 2", { rowIndex: 0, column: 1 }),
            );
            continue;
          }
          expected.push(
            row(`r${sourceRow}c0`, `Row ${sourceRow} cell 1`, {
              rowIndex: sourceRow - Number(sourceRow > deletedRow),
              column: 0,
            }),
            row(`r${sourceRow}c1`, `Row ${sourceRow} cell 2`, {
              rowIndex: sourceRow - Number(sourceRow > deletedRow),
              column: 1,
            }),
          );
        }
        assert.deepEqual(
          compareWithModel(model, expected),
          [],
          `row count ${rowCount}, deleted row ${deletedRow}`,
        );
      }
    }
  });

  test("keeps row payloads on source columns across insertions and deletions", () => {
    for (const edit of ["insert", "delete"] as const) {
      for (let width = 2; width <= 4; width += 1) {
        for (let columnMask = 1; columnMask < 2 ** width; columnMask += 1) {
          if (edit === "delete" && columnMask === 2 ** width - 1) continue;
          const cells = Array.from({ length: width * 2 }, (_, index) => {
            const rowIndex = Math.floor(index / width);
            const column = index % width;
            return row(`r${rowIndex}c${column}`, `${rowIndex}:${column}`, {
              rowIndex,
              column,
            });
          });
          const model = modelOf(cells);
          for (let sourceColumn = 0; sourceColumn < width; sourceColumn += 1) {
            if ((columnMask & (1 << sourceColumn)) === 0) continue;
            expectOperation(model, {
              type: edit === "insert" ? "insertTableColumn" : "deleteTableColumn",
              blockId: `r0c${sourceColumn}`,
              position: "after",
            });
          }
          const rowAnchorColumn =
            edit === "delete"
              ? Array.from({ length: width }, (_, column) => column).find(
                  (column) => (columnMask & (1 << column)) === 0,
                )
              : 0;
          assert.ok(rowAnchorColumn !== undefined);
          expectOperation(model, {
            type: "insertTableRow",
            blockId: `r1c${rowAnchorColumn}`,
            position: "after",
            cellTexts: Array.from({ length: width }, (_, column) => `New cell ${column + 1}`),
          });

          const table = model.tables.tables[0];
          assert.ok(table);
          const actualRow = table.cells
            .filter(({ row: rowIndex }) => rowIndex === 2)
            .sort((left, right) => left.column - right.column)
            .map(({ column, paragraphs }) => [column, paragraphs]);
          const expectedRow: [number, string[]][] = [];
          let finalColumn = 0;
          for (let sourceColumn = 0; sourceColumn < width; sourceColumn += 1) {
            if (edit === "delete" && (columnMask & (1 << sourceColumn)) !== 0) continue;
            expectedRow.push([finalColumn, [`New cell ${sourceColumn + 1}`]]);
            finalColumn += 1;
            if (edit === "insert" && (columnMask & (1 << sourceColumn)) !== 0) {
              expectedRow.push([finalColumn, [""]]);
              finalColumn += 1;
            }
          }
          assert.deepEqual(actualRow, expectedRow, `${edit}, width ${width}, mask ${columnMask}`);
        }
      }
    }
  });
});
