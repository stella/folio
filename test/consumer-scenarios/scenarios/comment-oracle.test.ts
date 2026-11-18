import assert from "node:assert/strict";
import { test } from "node:test";

import {
  compareComments,
  expectOperation,
  modelOf,
  type Row,
} from "../support/oracle.ts";

const row = (id: string, text: string): Row => ({
  id,
  text,
  kind: "paragraph",
});

test("a comment range moved onto a surviving paragraph remains anchored", () => {
  const old = row("old", "Anchor");
  const next = row("next", "survives");
  const model = modelOf([old, next]);
  expectOperation(model, { type: "deleteBlock", blockId: old.id });
  const before = { id: 1, text: "review", anchor: "Anchor", blockId: old.id };

  assert.deepEqual(
    compareComments(
      model,
      [before],
      [{ ...before, blockId: next.id }],
      [next],
      [before],
      "direct",
    ),
    [],
  );
  assert.match(
    compareComments(model, [before], [before], [next], [before], "direct").join(
      "\n",
    ),
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
      const destinationRow = row(
        `destination-${String(destination)}`,
        "survives",
      );
      const model = modelOf([deletedRow, destinationRow]);
      expectOperation(model, { type: "deleteBlock", blockId: deletedRow.id });
      const before = {
        id: 1,
        text: "review",
        anchor: "Anchor",
        blockId: deletedRow.id,
      };
      const survivor = { ...before, blockId: destinationRow.id };

      assert.deepEqual(
        compareComments(
          model,
          [before],
          [survivor],
          [destinationRow],
          [before],
          "direct",
        ),
        [],
      );
      assert.match(
        compareComments(
          model,
          [before],
          [survivor],
          [],
          [before],
          "direct",
        ).join("\n"),
        /outlived/u,
      );
    }
  }
});

test("a joined comment reads under the last paragraph once the batch deletes it", () => {
  // Live, "Kept. " runs on into "Joined." through a pending paragraph-mark
  // deletion; accepted, they are one paragraph, "joined".
  const joined = row("joined", "Kept. Joined.");
  const closing = row("closing", "Closing.");
  const source = row("source", "Kept. ");
  const live = [source, { ...joined, text: "Joined." }, closing];
  const before = { id: 1, text: "review", anchor: "Kept.", blockId: joined.id };
  const liveBefore = { ...before, blockId: source.id };
  const compare = (deletesClosing: boolean, blockId: string) => {
    const model = modelOf([joined, closing], live);
    if (deletesClosing)
      expectOperation(model, { type: "deleteBlock", blockId: closing.id });
    return compareComments(
      model,
      [before],
      [{ ...before, blockId }],
      deletesClosing ? [{ ...joined, id: closing.id }] : [joined, closing],
      [liveBefore],
      "tracked-changes",
    );
  };

  // Deleting a story's last paragraph keeps its mark: the joined words run
  // into it and read under its id.
  assert.deepEqual(compare(true, closing.id), []);
  assert.match(compare(true, joined.id).join("\n"), /no comment/u);
  // Otherwise they keep the id of the paragraph the join runs into.
  assert.deepEqual(compare(false, joined.id), []);
  assert.match(compare(false, closing.id).join("\n"), /no comment/u);
});
