import assert from "node:assert/strict";
import { test } from "node:test";

import type { ReviewerEngine } from "../support/documents.ts";
import {
  compareWithModel,
  expectOperation,
  modelOf,
  type Row,
} from "../support/oracle.ts";

const ENGINES = ["canonical"] as const satisfies readonly ReviewerEngine[];

const row = (index: number): Row => ({
  id: `block-${index}`,
  text: `Block ${index}`,
  kind: "paragraph",
});

type MergeSkipsDeletedBlocksOptions = {
  count: number;
  deleted: number;
  followingTable?: boolean;
  engine?: ReviewerEngine;
};

/**
 * A merge across the paragraphs its batch deletes joins the next surviving
 * paragraph. Before a table, the legacy engine stops the join at the table;
 * the canonical engine runs the words into its first cell (D1).
 */
const assertMergeSkipsDeletedBlocks = ({
  count,
  deleted,
  followingTable = false,
  engine,
}: MergeSkipsDeletedBlocksOptions) => {
  const rows: Row[] = Array.from({ length: count }, (_, index) => row(index));
  if (followingTable) {
    rows.push({
      ...row(count),
      table: {
        outerTableIndex: 0,
        tableIndex: 0,
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        columnSpan: 1,
        rowSpan: 1,
      },
    });
  }
  const model = modelOf(rows);
  if (engine !== undefined) model.engine = engine;
  const first = rows.at(0);
  assert.ok(first);
  expectOperation(model, {
    id: "merge",
    type: "mergeBlockWithNext",
    blockId: first.id,
    separator: " ",
  });
  for (const removed of rows.slice(1, deleted + 1)) {
    expectOperation(model, {
      id: `delete-${removed.id}`,
      type: "deleteBlock",
      blockId: removed.id,
    });
  }
  const survivor = rows.at(deleted + 1);
  if (survivor?.table !== undefined && model.engine === "canonical") {
    const runIn = [
      { ...survivor, text: `Block 0 ${survivor.text}` },
      ...rows.slice(deleted + 2),
    ];
    assert.deepEqual(compareWithModel(model, runIn), []);
    return;
  }
  const joinsSurvivor = survivor !== undefined && survivor.table === undefined;
  const actual = [
    {
      ...first,
      text: joinsSurvivor
        ? `Block 0 ${survivor.text}`
        : `Block 0${survivor ? " " : ""}`,
    },
    ...rows.slice(deleted + (joinsSurvivor ? 2 : 1)),
  ];
  assert.deepEqual(compareWithModel(model, actual), []);
};

for (const seed of [
  2146507783, 1809931145, 1809931025, 1809931181, 1101252716, 1101252706,
]) {
  test(`seed ${seed}: a merge crosses blocks deleted by its batch`, () => {
    assertMergeSkipsDeletedBlocks({ count: 4, deleted: 1 + (seed % 2) });
  });
}

for (const engine of ENGINES) {
  test(`merge/delete oracle property covers every following deletion span (${engine})`, () => {
    for (let count = 2; count <= 12; count++) {
      for (let deleted = 1; deleted < count; deleted++) {
        assertMergeSkipsDeletedBlocks({ count, deleted, engine });
        assertMergeSkipsDeletedBlocks({
          count,
          deleted,
          followingTable: true,
          engine,
        });
      }
    }
  });

  test(`merge whose next surviving block is a table cell keeps the separator (${engine})`, () => {
    assertMergeSkipsDeletedBlocks({
      count: 2,
      deleted: 1,
      followingTable: true,
      engine,
    });
  });
}

const textsOf = (texts: string[], prefix = "block"): Row[] =>
  texts.map((text, index) => ({
    id: `${prefix}-${index}`,
    text,
    kind: "paragraph",
  }));

test("a merge across a deleted paragraph stays apart from a table the batch inserts there", () => {
  const rows = textsOf(["Block 0", "Block 1"]);
  const model = modelOf(rows);
  model.mode = "tracked-changes";
  const [first, second] = rows;
  assert.ok(first && second);
  expectOperation(model, {
    id: "table",
    type: "insertTable",
    blockId: second.id,
    rows: [["Term", "Value"]],
  });
  expectOperation(model, {
    id: "merge",
    type: "mergeBlockWithNext",
    blockId: first.id,
    separator: " ",
  });
  expectOperation(model, {
    id: "delete",
    type: "deleteBlock",
    blockId: second.id,
  });
  const cells = textsOf(["Term", "Value"], "cell");
  assert.deepEqual(
    compareWithModel(model, [{ ...first, text: "Block 0 " }, ...cells]),
    [],
  );
  assert.notDeepEqual(
    compareWithModel(model, [
      { ...first, text: "Block 0 Term" },
      ...cells.slice(1),
    ]),
    [],
  );
});

test("a merge into a last paragraph whose words are pending deletions adds no separator", () => {
  const rows = textsOf(["Block 0", ""]);
  const [first, last] = rows;
  assert.ok(first && last);
  const merged = (deletedContentIds: ReadonlySet<string>) => {
    const model = modelOf(rows, rows, {
      type: "marked",
      ids: new Set(),
      emptyIds: new Set(),
      deletedContentIds,
    });
    model.mode = "tracked-changes";
    expectOperation(model, {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: first.id,
      separator: " ",
    });
    return model;
  };
  const actual = (text: string) => [{ ...first, text }];
  assert.deepEqual(
    compareWithModel(merged(new Set([last.id])), actual("Block 0")),
    [],
  );
  // A blank paragraph of its own still takes the separator.
  assert.deepEqual(compareWithModel(merged(new Set()), actual("Block 0 ")), []);
});

for (const engine of ENGINES)
  test(`a merge across a deleted story paragraph before a table keeps the separator (${engine})`, () => {
    const rows: Row[] = [
      ...textsOf(["Block 0", "Block 1"]),
      {
        id: "cell-0",
        text: "Term",
        kind: "paragraph",
        table: {
          outerTableIndex: 0,
          tableIndex: 0,
          rowIndex: 0,
          cellIndex: 0,
          gridColumnIndex: 0,
          columnSpan: 1,
          rowSpan: 1,
        },
      },
    ];
    const [first, second, cell] = rows;
    assert.ok(first && second && cell);
    const model = modelOf(rows);
    model.engine = engine;
    model.mode = "tracked-changes";
    // The table follows the paragraph in the story, so the paragraph does not
    // end it: deleting it removes it. The legacy engine stops the join before
    // the table; the canonical one runs it into the first cell (D1).
    expectOperation(model, {
      id: "delete",
      type: "deleteBlock",
      blockId: second.id,
    });
    expectOperation(model, {
      id: "merge",
      type: "mergeBlockWithNext",
      blockId: first.id,
      separator: " ",
    });
    if (engine === "canonical") {
      assert.deepEqual(
        compareWithModel(model, [{ ...cell, text: "Block 0 Term" }]),
        [],
      );
      return;
    }
    assert.deepEqual(
      compareWithModel(model, [{ ...first, text: "Block 0 " }, cell]),
      [],
    );
    assert.notDeepEqual(
      compareWithModel(model, [{ ...first, text: "Block 0" }, cell]),
      [],
    );
  });

test("a pending join's source does not follow the last paragraph of its story", () => {
  // Accepted, the source's words run into `joined`; the model lists the
  // live-only source after every row.
  const rows = textsOf(["Head Block 0", "Block 1", ""]);
  const [joined, middle, last] = rows;
  assert.ok(joined && middle && last);
  const live: Row[] = [
    { id: "source", text: "Head ", kind: "paragraph" },
    { ...joined, text: "Block 0" },
    middle,
    last,
  ];
  const model = modelOf(rows, live, {
    type: "marked",
    ids: new Set(["source"]),
    emptyIds: new Set(),
    deletedContentIds: new Set([last.id]),
  });
  model.mode = "tracked-changes";
  expectOperation(model, {
    id: "merge",
    type: "mergeBlockWithNext",
    blockId: middle.id,
    separator: " ",
  });
  assert.deepEqual(compareWithModel(model, [joined, middle]), []);
  assert.notDeepEqual(
    compareWithModel(model, [joined, { ...middle, text: "Block 1 " }]),
    [],
  );
});
