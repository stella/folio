import assert from "node:assert/strict";
import { test } from "node:test";

import { compareWithModel, expectOperation, modelOf, type Row } from "../support/oracle.ts";

const row = (index: number): Row => ({
  id: `block-${index}`,
  text: `Block ${index}`,
  kind: "paragraph",
});

const assertMergeSkipsDeletedBlocks = (count: number, deleted: number) => {
  const rows = Array.from({ length: count }, (_, index) => row(index));
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
  const actual = [
    { ...first, text: survivor ? `Block 0 ${survivor.text}` : "Block 0" },
    ...rows.slice(deleted + 2),
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
    }
  }
});
