import assert from "node:assert/strict";
import { test } from "node:test";

import registry from "../test/consumer-scenarios/scenario-seeds.json" with { type: "json" };
import { parseScenarioSeedRegistry } from "../test/consumer-scenarios/support/scenario-seeds.ts";

test("the pinned scenario registry validates and keeps its declared replay order", () => {
  const flows = parseScenarioSeedRegistry(registry);

  assert.equal(flows.at(0)?.seed, 20_260_933);
  // Every pinned flow replays in file order, whatever a fix appends.
  assert.deepEqual(
    flows.map(({ seed, kind }) => `${String(seed)}:${kind}`),
    registry.flows.map(({ seed, kind }) => `${String(seed)}:${kind}`),
  );
  assert.deepEqual(
    flows.filter(({ seed }) => seed === 1_250_352_731).map(({ kind }) => kind),
    ["random", "collisions"],
  );
});

test("the registry reader rejects malformed and duplicate flows", () => {
  assert.throws(() => parseScenarioSeedRegistry({ flows: [{ seed: "bad" }] }), TypeError);
  assert.throws(
    () =>
      parseScenarioSeedRegistry({
        flows: [
          {
            seed: 1,
            steps: 1,
            kind: "random",
            generation: "targeted",
            title: "first",
          },
          {
            seed: 1,
            steps: 1,
            kind: "random",
            generation: "targeted",
            title: "duplicate",
          },
        ],
      }),
    /duplicate flow/u,
  );
});
