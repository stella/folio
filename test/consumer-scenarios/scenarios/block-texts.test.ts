import assert from "node:assert/strict";
import { test } from "node:test";

import { compareWithModel, expectOperation, modelOf, type Row } from "../support/oracle.ts";

const row = (index: number): Row => ({
  id: `block-${index}`,
  text: `Block ${index}`,
  kind: "paragraph",
});

const assertMergeSkipsDeletedBlocks = (count: number, deleted: number, followingTable = false) => {
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
  const joinsSurvivor = survivor !== undefined && survivor.table === undefined;
  const actual = [
    {
      ...first,
      text: joinsSurvivor ? `Block 0 ${survivor.text}` : `Block 0${survivor ? " " : ""}`,
    },
    ...rows.slice(deleted + (joinsSurvivor ? 2 : 1)),
  ];
  assert.deepEqual(compareWithModel(model, actual), []);
};

for (const seed of [2146507783, 1809931145, 1809931025, 1809931181, 1101252716, 1101252706]) {
  test(`seed ${seed}: a merge crosses blocks deleted by its batch`, () => {
    assertMergeSkipsDeletedBlocks(4, 1 + (seed % 2));
  });
}

test("merge/delete oracle property covers every following deletion span", () => {
  for (let count = 2; count <= 12; count++) {
    for (let deleted = 1; deleted < count; deleted++) {
      assertMergeSkipsDeletedBlocks(count, deleted);
      assertMergeSkipsDeletedBlocks(count, deleted, true);
    }
  }
});

test("merge whose next surviving block is a table cell keeps the separator", () => {
  assertMergeSkipsDeletedBlocks(2, 1, true);
});
