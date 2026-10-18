import { expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty } from "../test/property-testing";
import { consumerScenarioInputs, parseConsumerArgs } from "./consumer-scenario-args";

test("CLI replay flags resolve to the same inputs as environment variables", () => {
  assertProperty(
    fc.property(
      fc.record({
        seed: fc.string({ minLength: 1, maxLength: 20 }),
        steps: fc.nat().map(String),
        runs: fc.nat().map(String),
        collisions: fc.nat().map(String),
        swarm: fc.constantFrom("0", "1"),
        flow: fc.string({ minLength: 1, maxLength: 50 }),
        corpus: fc.stringMatching(/^[a-zA-Z0-9_-]{1,20}$/u),
      }),
      ({ seed, steps, runs, collisions, swarm, flow, corpus }) => {
        const cli = parseConsumerArgs(
          [
            "--seed",
            seed,
            "--steps",
            steps,
            "--runs",
            runs,
            "--collision-runs",
            collisions,
            "--swarm",
            swarm,
            "--flow",
            flow,
            "--public-corpus-dir",
            corpus,
          ],
          "/repo",
        );
        const empty = parseConsumerArgs([], "/repo");
        expect(consumerScenarioInputs(cli, {})).toEqual(
          consumerScenarioInputs(empty, {
            FOLIO_SCENARIO_SEED: seed,
            FOLIO_SCENARIO_FUZZ_STEPS: steps,
            FOLIO_SCENARIO_FUZZ_RUNS: runs,
            FOLIO_SCENARIO_COLLISION_RUNS: collisions,
            FOLIO_SCENARIO_SWARM: swarm,
            FOLIO_SCENARIO_FLOW: flow,
            FOLIO_SCENARIO_PUBLIC_CORPUS_DIR: cli.publicCorpusDir ?? corpus,
          }),
        );
      },
    ),
    { numRuns: 100, id: "CLI and environment replay inputs agree" },
  );
});
