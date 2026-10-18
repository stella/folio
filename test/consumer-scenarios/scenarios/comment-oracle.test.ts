import assert from "node:assert/strict";
import { test } from "node:test";

import { compareComments, expectOperation, modelOf, type Row } from "../support/oracle.ts";

const row = (id: string, text: string): Row => ({ id, text, kind: "paragraph" });

test("a comment range moved onto a surviving paragraph remains anchored", () => {
  const old = row("old", "Anchor");
  const next = row("next", "survives");
  const model = modelOf([old, next]);
  expectOperation(model, { type: "deleteBlock", blockId: old.id });
  const before = { id: 1, text: "review", anchor: "Anchor", blockId: old.id };

  assert.deepEqual(
    compareComments(model, [before], [{ ...before, blockId: next.id }], [next], [before], "direct"),
    [],
  );
  assert.match(
    compareComments(model, [before], [before], [next], [before], "direct").join("\n"),
    /outlived/u,
  );
});

test("a comment moved onto another deleted paragraph is still an orphan", () => {
  const first = row("first", "Anchor");
  const second = row("second", "gone");
  const third = row("third", "survives");
  const model = modelOf([first, second, third]);
  expectOperation(model, { type: "deleteBlock", blockId: first.id });
  expectOperation(model, { type: "deleteBlock", blockId: second.id });
  const before = { id: 1, text: "review", anchor: "Anchor", blockId: first.id };

  assert.match(
    compareComments(
      model,
      [before],
      [{ ...before, blockId: second.id }],
      [third],
      [before],
      "direct",
    ).join("\n"),
    /outlived/u,
  );
  assert.deepEqual(
    compareComments(
      model,
      [before],
      [{ ...before, blockId: third.id }],
      [third],
      [before],
      "direct",
    ),
    [],
  );
});

test("migrated comment ranges require a surviving destination over generated ids", () => {
  for (let deleted = 0; deleted < 8; deleted += 1) {
    for (let destination = 0; destination < 8; destination += 1) {
      const deletedRow = row(`deleted-${String(deleted)}`, "Anchor");
      const destinationRow = row(`destination-${String(destination)}`, "survives");
      const model = modelOf([deletedRow, destinationRow]);
      expectOperation(model, { type: "deleteBlock", blockId: deletedRow.id });
      const before = { id: 1, text: "review", anchor: "Anchor", blockId: deletedRow.id };
      const survivor = { ...before, blockId: destinationRow.id };

      assert.deepEqual(
        compareComments(model, [before], [survivor], [destinationRow], [before], "direct"),
        [],
      );
      assert.match(
        compareComments(model, [before], [survivor], [], [before], "direct").join("\n"),
        /outlived/u,
      );
    }
  }
});
