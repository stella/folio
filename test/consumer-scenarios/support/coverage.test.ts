import assert from "node:assert/strict";
import { test } from "node:test";

import { hitUnreachableCells, type Expectations, type Ledger } from "./coverage.ts";

const expectations: Expectations = {
  required: [],
  unreachable: [
    {
      cell: { story: "header", feature: "pendingRevision", mode: "suggested" },
      reason: "no public reader exposes a pending header suggestion",
    },
    {
      cell: { op: "mergeTableCells" },
      reason: "no generator yet",
    },
  ],
};

test("unreachable patterns reject both applied and refused hits across wildcard dimensions", () => {
  const merged: Ledger = {
    version: 1,
    cells: {
      "insertText | header | suggested | pendingRevision | fresh": { applied: 2, refused: 0 },
      "deleteText | header | suggested | pendingRevision | reopened": {
        applied: 0,
        refused: 3,
      },
      "mergeTableCells | main | direct | table | fresh": { applied: 0, refused: 1 },
      "insertText | footer | direct | pendingRevision | fresh": { applied: 4, refused: 2 },
    },
  };

  assert.deepEqual(hitUnreachableCells(merged, expectations), [
    {
      cell: expectations.unreachable[0]?.cell,
      reason: expectations.unreachable[0]?.reason,
      applied: 2,
      refused: 3,
    },
    {
      cell: expectations.unreachable[1]?.cell,
      reason: expectations.unreachable[1]?.reason,
      applied: 0,
      refused: 1,
    },
  ]);
});

test("unreachable patterns with zero hits pass", () => {
  const merged: Ledger = {
    version: 1,
    cells: {
      "insertText | footer | direct | pendingRevision | fresh": { applied: 1, refused: 0 },
    },
  };

  assert.deepEqual(hitUnreachableCells(merged, expectations), []);
});
