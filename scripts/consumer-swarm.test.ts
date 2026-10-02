import { expect, test } from "bun:test";

import { type FlowFile, flowId, parseFlowFile } from "../test/consumer-scenarios/support/flow-file";
import { mutateFlow } from "../test/consumer-scenarios/support/flow-mutate";
import { createRandom } from "../test/consumer-scenarios/support/random";
import { shrinkFlow } from "../test/consumer-scenarios/support/shrink";
import { drawSwarm, swarmIncludesBatch } from "../test/consumer-scenarios/support/swarm";

const kinds = ["replaceInBlock", "insertAfterBlock", "deleteBlock", "formatRange"];
const parent = {
  version: 1,
  kind: "collisions",
  generation: "targeted",
  fixture: "plain",
  mode: "direct",
  seed: 7,
  steps: [
    { action: "core batch", seed: 1 },
    { action: "save and reopen", seed: 2 },
  ],
} satisfies FlowFile;

test("seeded swarms disable kinds, remain nonempty, and reproduce independently", () => {
  const seen = new Set<string>();
  const enabledKinds = new Set<string>();
  for (let seed = 0; seed < 256; seed += 1) {
    const selected = drawSwarm(seed, kinds);
    expect(selected).toEqual(drawSwarm(seed, kinds));
    expect(selected.length).toBeGreaterThan(0);
    expect(selected.length).toBeLessThan(kinds.length);
    expect(new Set(selected).size).toBe(selected.length);
    for (const kind of selected) {
      expect(kinds).toContain(kind);
      enabledKinds.add(kind);
    }
    seen.add(JSON.stringify(selected));
  }
  expect([...enabledKinds].sort()).toEqual([...kinds].sort());
  expect(seen.size).toBeGreaterThan(kinds.length);
  expect(drawSwarm(7, ["replaceInBlock"])).toEqual(["replaceInBlock"]);
  expect(() => drawSwarm(7, [])).toThrow(TypeError);
});

test("swarm data roundtrips, changes identity, and rejects empty or duplicate kinds", () => {
  const configured = { ...parent, swarm: drawSwarm(parent.seed, kinds) };
  expect(parseFlowFile(JSON.parse(JSON.stringify(configured)))).toEqual(configured);
  expect(flowId(configured)).not.toBe(flowId(parent));
  expect(parseFlowFile(parent)).toEqual(parent);
  for (const swarm of [[], [""], ["replaceInBlock", "replaceInBlock"], [7]]) {
    expect(() => parseFlowFile({ ...parent, swarm })).toThrow(/swarm/u);
  }
});

test("collision batches and pinned operations cannot bypass any disabled kind", () => {
  for (let seed = 0; seed < 256; seed += 1) {
    const swarm = drawSwarm(seed, kinds);
    expect(
      swarmIncludesBatch(
        swarm,
        swarm.map((type) => ({ type })),
      ),
    ).toBe(true);
    for (const type of kinds) {
      expect(swarmIncludesBatch(swarm, [{ type }])).toBe(swarm.includes(type));
    }
    expect(
      swarmIncludesBatch(
        swarm,
        kinds.map((type) => ({ type })),
      ),
    ).toBe(false);
  }
  expect(
    swarmIncludesBatch(
      undefined,
      kinds.map((type) => ({ type })),
    ),
  ).toBe(true);
});

test("mutation keeps the swarm and redraws donors from another swarm", () => {
  const configured = { ...parent, swarm: ["replaceInBlock"] };
  const donor = {
    ...parent,
    swarm: ["deleteBlock"],
    steps: [{ action: "core batch", seed: 999, operations: [{ type: "deleteBlock" }] }],
  } satisfies FlowFile;
  let spliced = false;
  for (let seed = 0; seed < 256; seed += 1) {
    const mutated = mutateFlow(configured, createRandom(seed), { donors: [donor] });
    expect(mutated.flow.swarm).toEqual(configured.swarm);
    expect(parseFlowFile(mutated.flow)).toEqual(mutated.flow);
    for (const step of mutated.flow.steps) {
      expect(step.operations).toBeUndefined();
      if (step.seed === 999) spliced = true;
    }
  }
  expect(spliced).toBe(true);
});

test("shrinking keeps the recorded swarm while deleting steps", async () => {
  const configured = { ...parent, swarm: ["replaceInBlock"] };
  const shrunk = await shrinkFlow(configured, {
    holds: (candidate) => Promise.resolve(candidate.steps.some(({ seed }) => seed === 1)),
    materialize: (candidate) => Promise.resolve(candidate),
    budget: { maxAttempts: 20, deadline: Date.now() + 10_000 },
  });
  expect(shrunk.flow.steps).toHaveLength(1);
  expect(shrunk.flow.swarm).toEqual(configured.swarm);
});
