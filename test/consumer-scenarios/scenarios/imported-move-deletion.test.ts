import assert from "node:assert/strict";
import { test } from "node:test";

import { IMPORTED_MOVE_FIXTURE } from "../support/documents.ts";
import { runFlowFile, vacuousSteps } from "../support/fuzz.ts";

// Ordinary tracked-operation fixtures did not generate imported move pairs.
// Exercise both endpoints on every run, with the existing accept/reject oracle.
for (const [endpoint, blockId] of [
  ["source", "@1"],
  ["destination", "@3"],
] as const) {
  test(`deleteBlock over an imported move ${endpoint}`, async () => {
    const result = await runFlowFile({
      version: 1,
      kind: "random",
      generation: "targeted",
      fixture: IMPORTED_MOVE_FIXTURE,
      mode: "direct",
      seed: 10357178,
      steps: [
        {
          action: "core batch",
          seed: 2849829543,
          operations: [{ type: "deleteBlock", blockId }],
          atomic: true,
        },
      ],
    });
    assert.deepEqual(vacuousSteps(result), []);
  });
}
