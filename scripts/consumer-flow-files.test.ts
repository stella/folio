import { describe, expect, test } from "bun:test";
import { mkdtempSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { admit, loadCorpus, newElements, prune } from "../test/consumer-scenarios/support/corpus";
import {
  type FlowFile,
  flowId,
  flowShape,
  parseFlowFile,
  TARGETED_ACTIONS,
} from "../test/consumer-scenarios/support/flow-file";
import { mutateFlow } from "../test/consumer-scenarios/support/flow-mutate";
import { createRandom } from "../test/consumer-scenarios/support/random";
import { shrinkFlow, shrinkSequence } from "../test/consumer-scenarios/support/shrink";

const flow = (steps: FlowFile["steps"], fixture = "plain"): FlowFile => ({
  version: 1,
  kind: "random",
  generation: "targeted",
  fixture,
  mode: "direct",
  seed: 7,
  steps,
});

const budget = () => ({ maxAttempts: 1_000, deadline: Date.now() + 60_000 });

describe("flow files", () => {
  test("a random generator resumes from its recorded state", () => {
    const random = createRandom(12_345);
    random.next();
    const resumed = createRandom(random.state());
    expect([random.next(), random.int(100), random.next()]).toEqual([
      resumed.next(),
      resumed.int(100),
      resumed.next(),
    ]);
  });

  test("parsing checks every step and keeps pinned operations only on batch steps", () => {
    const valid = flow([
      { action: "core batch", seed: 1, operations: [{ type: "insertText" }] },
      { action: "accept all", seed: 4_294_967_295 },
    ]);
    expect(parseFlowFile(JSON.parse(JSON.stringify(valid)))).toEqual(valid);
    expect(() => parseFlowFile({ ...valid, steps: [] })).toThrow(TypeError);
    expect(() => parseFlowFile({ ...valid, steps: [{ action: "explode", seed: 1 }] })).toThrow(
      /unknown action/u,
    );
    expect(() => parseFlowFile({ ...valid, steps: [{ action: "accept all", seed: -1 }] })).toThrow(
      /seed/u,
    );
    expect(() =>
      parseFlowFile({
        ...valid,
        steps: [{ action: "accept all", seed: 1, operations: [{ type: "insertText" }] }],
      }),
    ).toThrow(/batch step/u);
    expect(flowId(valid)).toBe(flowId({ ...valid, origin: "elsewhere" }));
  });

  test("a flow's shape names its operations and actions, not its ids or text", () => {
    const shaped = flow([
      {
        action: "story batch",
        seed: 1,
        operations: [
          { type: "replaceRange", blockId: "@1" },
          { type: "insertAfterBlock", blockId: "0F88C890", text: "A heading." },
          { type: "replaceRange", blockId: "1508FAF4" },
        ],
      },
      { action: "save and reopen", seed: 2 },
      { action: "core batch", seed: 3 },
    ]);
    expect(flowShape(shaped)).toBe("insertAfterBlock+replaceRange > save and reopen > core batch");
  });
});

describe("shrinking", () => {
  test("delta debugging finds the steps a failure needs", async () => {
    const items = Array.from({ length: 20 }, (_, index) => index);
    const tracker = { attempts: 0, budget: budget() };
    const kept = await shrinkSequence(
      items,
      (candidate) => Promise.resolve(candidate.includes(3) && candidate.includes(17)),
      tracker,
    );
    expect(kept).toEqual([3, 17]);
    expect(tracker.attempts).toBeLessThan(60);
  });

  test("a spent budget stops the shrink with what it has", async () => {
    const tracker = { attempts: 0, budget: { maxAttempts: 3, deadline: Date.now() + 60_000 } };
    const kept = await shrinkSequence(
      [1, 2, 3, 4, 5, 6, 7, 8],
      () => Promise.resolve(true),
      tracker,
    );
    expect(tracker.attempts).toBe(3);
    expect(kept.length).toBeGreaterThan(0);
  });

  test("a flow shrinks over its steps, then over a pinned batch's operations", async () => {
    const operations = [{ type: "a" }, { type: "b" }, { type: "c" }, { type: "d" }];
    const original = flow(
      Array.from({ length: 8 }, (_, index) => ({
        action: index === 5 ? ("core batch" as const) : ("accept all" as const),
        seed: index,
      })),
    );
    // Fails while step seed 5 applies operation "c" (drawn: all four).
    const fails = (candidate: FlowFile): Promise<boolean> =>
      Promise.resolve(
        candidate.steps.some(
          (step) =>
            step.seed === 5 &&
            (step.operations ?? operations).some((operation) => operation["type"] === "c"),
        ),
      );
    const materialize = (candidate: FlowFile): Promise<FlowFile> =>
      Promise.resolve({
        ...candidate,
        steps: candidate.steps.map((step) =>
          step.action === "core batch" ? { ...step, operations } : step,
        ),
      });
    const shrunk = await shrinkFlow(original, { holds: fails, materialize, budget: budget() });
    expect(shrunk.from).toBe(8);
    expect(shrunk.flow.steps).toEqual([
      { action: "core batch", seed: 5, operations: [{ type: "c" }] },
    ]);
    expect(await fails(shrunk.flow)).toBe(true);
  });
});

describe("mutation", () => {
  const parent = flow(
    Array.from({ length: 6 }, (_, index) => ({
      action: TARGETED_ACTIONS[
        index % TARGETED_ACTIONS.length
      ] as FlowFile["steps"][number]["action"],
      seed: index,
      ...(index === 0 ? { operations: [{ type: "a" }, { type: "b" }] } : {}),
    })),
  );

  test("mutants are valid flow files within the step limit, and replay from their seed", () => {
    for (let seed = 0; seed < 200; seed += 1) {
      const { flow: mutant, mutations } = mutateFlow(parent, createRandom(seed), {
        donors: [
          parent,
          flow([{ action: "core batch", seed: 9, operations: [{ type: "x" }] }], "table"),
        ],
        maxSteps: 8,
      });
      expect(mutations.length).toBeGreaterThan(0);
      expect(parseFlowFile(JSON.parse(JSON.stringify(mutant)))).toMatchObject({
        fixture: parent.fixture,
        mode: parent.mode,
      });
      expect(mutant.steps.length).toBeGreaterThan(0);
      expect(mutant.steps.length).toBeLessThanOrEqual(8);
      expect(
        mutateFlow(parent, createRandom(seed), { donors: [parent], maxSteps: 8 }).flow,
      ).toEqual(mutateFlow(parent, createRandom(seed), { donors: [parent], maxSteps: 8 }).flow);
    }
  });

  test("a splice from another fixture keeps the donor's steps but not its pinned ids", () => {
    const donor = flow([{ action: "core batch", seed: 9, operations: [{ type: "x" }] }], "table");
    for (let seed = 0; seed < 200; seed += 1) {
      const { flow: mutant, mutations } = mutateFlow(parent, createRandom(seed), {
        donors: [donor],
      });
      if (mutations.length !== 1 || mutations[0] !== "splice") continue;
      const spliced = mutant.steps.find((step) => step.seed === 9);
      expect(spliced).toEqual({ action: "core batch", seed: 9 });
    }
  });
});

describe("corpus", () => {
  test("a flow joins only with a new signature element, and the cache persists it", () => {
    const dir = mkdtempSync(path.join(tmpdir(), "folio-corpus-"));
    try {
      const corpus = loadCorpus(dir);
      const first = flow([{ action: "accept all", seed: 1 }]);
      const second = flow([{ action: "reject all", seed: 2 }]);
      expect(admit(corpus, first, ["a", "b"], "2026-09-29T00:00:00Z")).toEqual(["a", "b"]);
      expect(admit(corpus, second, ["a"], "2026-09-29T00:01:00Z")).toEqual([]);
      expect(newElements(corpus, ["a", "c"])).toEqual(["c"]);
      expect(admit(corpus, second, ["a", "c"], "2026-09-29T00:02:00Z")).toEqual(["c"]);
      expect(readdirSync(dir)).toHaveLength(2);
      const reloaded = loadCorpus(dir, [flow([{ action: "save and reopen", seed: 3 }])]);
      expect(reloaded.entries.map((entry) => entry.id).sort()).toEqual(
        [
          flowId(first),
          flowId(second),
          flowId(flow([{ action: "save and reopen", seed: 3 }])),
        ].sort(),
      );
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test("pruning drops redundant entries first, then the oldest, never a checked-in seed", () => {
    const seedFlow = flow([{ action: "save and reopen", seed: 3 }]);
    const corpus = loadCorpus(null, [seedFlow]);
    admit(corpus, flow([{ action: "accept all", seed: 1 }]), ["a"], "2026-09-29T00:00:00Z");
    admit(corpus, flow([{ action: "accept all", seed: 2 }]), ["a", "b"], "2026-09-29T00:01:00Z");
    admit(corpus, flow([{ action: "accept all", seed: 3 }]), ["c"], "2026-09-29T00:02:00Z");
    expect(prune(corpus, 3)).toBe(1);
    expect(corpus.entries.map((entry) => entry.signature)).toEqual([[], ["a", "b"], ["c"]]);
    expect(prune(corpus, 1)).toBe(2);
    expect(corpus.entries.map((entry) => entry.id)).toEqual([flowId(seedFlow)]);
  });
});
