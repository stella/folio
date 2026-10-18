import assert from "node:assert/strict";
import { test } from "node:test";

import { runFlow } from "../support/fuzz.ts";
import { compareComments, expectOperation, modelOf, type Row } from "../support/oracle.ts";

const row = (id: string, text: string): Row => ({ id, text, kind: "paragraph" });

test("a comment range moved off a deleted paragraph remains anchored", () => {
  const model = modelOf([row("old", "Anchor"), row("next", "survives")]);
  expectOperation(model, { type: "deleteBlock", blockId: "old" });
  const before = { id: 1, text: "review", anchor: "Anchor", blockId: "old" };

  assert.deepEqual(compareComments(model, [before], [{ ...before, blockId: "next" }]), []);
  assert.match(compareComments(model, [before], [before]).join("\n"), /outlived/u);
});

test("a comment moved onto another deleted paragraph is still an orphan", () => {
  const model = modelOf([row("first", "Anchor"), row("second", "gone"), row("third", "survives")]);
  expectOperation(model, { type: "deleteBlock", blockId: "first" });
  expectOperation(model, { type: "deleteBlock", blockId: "second" });
  const before = { id: 1, text: "review", anchor: "Anchor", blockId: "first" };

  assert.match(
    compareComments(model, [before], [{ ...before, blockId: "second" }]).join("\n"),
    /outlived/u,
  );
  assert.deepEqual(compareComments(model, [before], [{ ...before, blockId: "third" }]), []);
});

test("migrated comment ranges are distinguished from orphans over generated ids", () => {
  for (let deleted = 0; deleted < 8; deleted += 1) {
    for (let destination = 0; destination < 8; destination += 1) {
      const deletedId = `block-${String(deleted)}`;
      const destinationId = `block-${String(destination)}`;
      const model = modelOf([row(deletedId, "Anchor")]);
      expectOperation(model, { type: "deleteBlock", blockId: deletedId });
      const before = { id: 1, text: "review", anchor: "Anchor", blockId: deletedId };
      const problems = compareComments(model, [before], [{ ...before, blockId: destinationId }]);
      if (destinationId === deletedId) assert.match(problems.join("\n"), /outlived/u);
      else assert.deepEqual(problems, []);
    }
  }
});

const PINNED_COMMENT_SEEDS = [
  { issue: 1752, seed: 1_809_930_833, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1763, seed: 746_947_958, steps: 16, kind: "random", swarm: "disabled" },
  { issue: 1812, seed: 1_339_716_521, steps: 20, kind: "random", swarm: "enabled" },
  { issue: 1813, seed: 1_339_716_112, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1824, seed: 1_101_252_668, steps: 16, kind: "random", swarm: "disabled" },
  { issue: 1862, seed: 740_097_170, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1866, seed: 740_097_276, steps: 20, kind: "collisions", swarm: "enabled" },
  { issue: 1867, seed: 740_097_637, steps: 20, kind: "random", swarm: "enabled" },
  { issue: 1878, seed: 740_097_416, steps: 20, kind: "collisions", swarm: "enabled" },
] as const;

for (const { issue, seed, steps, kind, swarm } of PINNED_COMMENT_SEEDS) {
  test(
    `issue #${String(issue)} comment locator regression (seed ${String(seed)})`,
    { timeout: 30_000 },
    async () => {
      await runFlow(seed, steps, kind, { swarm });
    },
  );
}
