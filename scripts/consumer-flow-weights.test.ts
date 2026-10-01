import { createHash } from "node:crypto";
import { expect, test } from "bun:test";

import {
  featureOperationHits,
  gapWeight,
  parseFeatureCoverage,
  weightedChoice,
} from "../test/consumer-scenarios/support/feature-coverage";
import { type FlowFile, flowId, parseFlowFile } from "../test/consumer-scenarios/support/flow-file";
import { mutateFlow } from "../test/consumer-scenarios/support/flow-mutate";
import { createRandom } from "../test/consumer-scenarios/support/random";
import { shrinkFlow } from "../test/consumer-scenarios/support/shrink";

const parent = {
  version: 1,
  kind: "random",
  generation: "targeted",
  fixture: "plain",
  mode: "direct",
  seed: 7,
  steps: [
    { action: "core batch", seed: 1 },
    { action: "save and reopen", seed: 2 },
  ],
} satisfies FlowFile;
const input = {
  version: 1,
  cells: { "replaceInBlock | heading | none": 200, "deleteBlock | none | none": 0 },
  operations: ["replaceInBlock", "deleteBlock"],
  emptyCells: ["a derived field that must stay out of the flow"],
};

test("flow weights snapshot validated inputs and discard derived report data", () => {
  const recorded = parseFlowFile({ ...parent, weights: input });
  expect(recorded.weights).toEqual({
    version: 1,
    cells: input.cells,
    operations: input.operations,
  });
  expect(recorded.weights?.cells).not.toBe(input.cells);
  expect(recorded.weights?.operations).not.toBe(input.operations);
  expect(flowId(recorded)).not.toBe(flowId(parent));
  const oldId = createHash("sha256")
    .update(
      JSON.stringify([
        parent.kind,
        parent.generation,
        parent.fixture,
        parent.mode,
        parent.seed,
        parent.steps,
      ]),
    )
    .digest("hex")
    .slice(0, 16);
  expect(flowId(parent)).toBe(oldId);
  for (const weights of [
    null,
    { ...input, version: 2 },
    { ...input, operations: [] },
    { ...input, cells: { "replaceInBlock | heading | none": -1 } },
    { ...input, cells: { "replaceInBlock | missing | none": 1 } },
    { ...input, cells: { "replaceInBlock | heading | none": Infinity } },
  ]) {
    expect(() => parseFlowFile({ ...parent, weights })).toThrow(TypeError);
  }
});

test("recorded seed generation remains identical after the external feedback changes", () => {
  const file = parseFlowFile({ ...parent, weights: input });
  const replay = parseFlowFile(JSON.parse(JSON.stringify(file)));
  const changed = parseFeatureCoverage({
    ...input,
    cells: { "replaceInBlock | heading | none": 0, "deleteBlock | none | none": 200 },
  });
  const draw = (weights: NonNullable<FlowFile["weights"]>, seed: number) => {
    const random = createRandom(seed);
    const hits = featureOperationHits(weights);
    return Array.from({ length: 30 }, () => [
      weightedChoice(weights.operations, random, (type) => gapWeight(hits[type] ?? 0)),
      random.chance(0.5),
    ]);
  };
  for (let seed = 0; seed < 100; seed += 1) {
    expect(file.weights && draw(file.weights, seed)).toEqual(
      replay.weights && draw(replay.weights, seed),
    );
  }
  expect(file.weights && draw(file.weights, file.seed)).not.toEqual(draw(changed, file.seed));
});

test("operation hit counts combine every feature and selection in the recorded input", () => {
  const weights = parseFeatureCoverage({
    ...input,
    cells: { ...input.cells, "replaceInBlock | footnote | paragraph-range": 7 },
  });
  expect(featureOperationHits(weights)).toEqual({ replaceInBlock: 207, deleteBlock: 0 });
});

test("mutation and shrinking retain feedback and swarm together", async () => {
  const file = parseFlowFile({ ...parent, weights: input, swarm: ["replaceInBlock"] });
  for (let seed = 0; seed < 100; seed += 1) {
    const mutated = mutateFlow(file, createRandom(seed)).flow;
    expect(mutated.weights).toEqual(file.weights);
    expect(mutated.swarm).toEqual(file.swarm);
  }
  const shrunk = await shrinkFlow(file, {
    holds: (candidate) => Promise.resolve(candidate.steps.some(({ seed }) => seed === 1)),
    materialize: (candidate) => Promise.resolve(candidate),
    budget: { maxAttempts: 20, deadline: Date.now() + 10_000 },
  });
  expect(shrunk.flow.steps).toHaveLength(1);
  expect(shrunk.flow.weights).toEqual(file.weights);
  expect(shrunk.flow.swarm).toEqual(file.swarm);
});
