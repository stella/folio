import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import path from "node:path";

import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import {
  CONSUMER_SCENARIO_REPLAY_INPUTS,
  consumerScenarioInputs,
  consumerScenarioReplay,
  consumerScenarioReplayInputs,
  parseConsumerArgs,
} from "./consumer-scenario-args";

setDefaultTimeout(propertyTestTimeout(30_000));

test("consumer scenario relative imports stay inside the copied tree", async () => {
  const scenarioRoot = path.join(import.meta.dir, "../test/consumer-scenarios");
  const transpiler = new Bun.Transpiler({ loader: "ts" });
  const escapingImports: string[] = [];
  const glob = new Bun.Glob("**/*.ts");

  for await (const file of glob.scan({ cwd: scenarioRoot, onlyFiles: true })) {
    const source = await Bun.file(path.join(scenarioRoot, file)).text();
    for (const imported of transpiler.scanImports(source)) {
      if (!imported.path.startsWith(".")) continue;
      const resolved = path.resolve(scenarioRoot, path.dirname(file), imported.path);
      const relative = path.relative(scenarioRoot, resolved);
      if (relative === ".." || relative.startsWith(`..${path.sep}`)) {
        escapingImports.push(`${file}: ${imported.path}`);
      }
    }
  }

  expect(escapingImports.sort()).toEqual([]);
});

const argvOf = (command: string): string[] =>
  [...command.matchAll(/'(?:'\\''|[^'])*'|\S+/gu)]
    .map(([word]) => (word.startsWith("'") ? word.slice(1, -1).replaceAll("'\\''", "'") : word))
    .slice(2);

test("flow replays put relation flags after the executable and parse them back", () => {
  const previous = {
    relations: process.env["FOLIO_SCENARIO_RELATIONS"],
    depth: process.env["FOLIO_SCENARIO_RELATIONS_DEPTH"],
  };
  process.env["FOLIO_SCENARIO_RELATIONS"] = "preserve-comments,round-trip";
  process.env["FOLIO_SCENARIO_RELATIONS_DEPTH"] = "3";
  try {
    const command = consumerScenarioReplay({
      inputs: { ...consumerScenarioReplayInputs(process.env), flow: "flow.json" },
      files: ["flow-corpus.test.ts"],
    });
    expect(command.split(" ").at(0)).toBe("bun");
    const parsed = parseConsumerArgs(argvOf(command), "/repo");
    expect(parsed.relations).toBe("preserve-comments,round-trip");
    expect(parsed.relationsDepth).toBe("3");
  } finally {
    if (previous.relations === undefined) delete process.env["FOLIO_SCENARIO_RELATIONS"];
    else process.env["FOLIO_SCENARIO_RELATIONS"] = previous.relations;
    if (previous.depth === undefined) delete process.env["FOLIO_SCENARIO_RELATIONS_DEPTH"];
    else process.env["FOLIO_SCENARIO_RELATIONS_DEPTH"] = previous.depth;
  }
});

test("the replay input table drives parsing, environment mapping and rendering", () => {
  for (const [index, input] of CONSUMER_SCENARIO_REPLAY_INPUTS.entries()) {
    const value = `/replay-input-${index}`;
    const environment = { [input.env]: value };
    const parsedFlag = parseConsumerArgs([input.flag, value], "/repo");
    const option = input.flag
      .slice(2)
      .replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase());
    const expected = value;
    expect(parsedFlag[option as keyof typeof parsedFlag]).toBe(expected);
    expect(consumerScenarioInputs(parseConsumerArgs([], "/repo"), environment)[option]).toBe(value);

    const replay = consumerScenarioReplay({
      inputs: consumerScenarioReplayInputs(environment),
    });
    expect(parseConsumerArgs(argvOf(replay), "/repo")[option as keyof typeof parsedFlag]).toBe(
      expected,
    );
  }
});

test("every environment name embedded in a scenario replay is in the input table", async () => {
  const replayInputs = new Set(CONSUMER_SCENARIO_REPLAY_INPUTS.map(({ env }) => env));
  const usedInReplays = new Set<string>();
  for (const root of ["scripts", "test/consumer-scenarios"]) {
    const glob = new Bun.Glob(`${root}/**/*`);
    for await (const file of glob.scan({ cwd: import.meta.dir + "/..", onlyFiles: true })) {
      const source = await Bun.file(import.meta.dir + "/../" + file).text();
      for (const [replay] of source.matchAll(/[^\n]*bun scripts\/consumer-scenarios\.ts[^\n]*/gu)) {
        for (const [name] of replay.matchAll(/FOLIO_SCENARIO_[A-Z_]+/gu)) usedInReplays.add(name);
      }
    }
  }
  expect([...usedInReplays].sort()).toEqual(
    [...usedInReplays].filter((name) => replayInputs.has(name)).sort(),
  );
});

test("generated replay settings round-trip entirely through CLI flags", () => {
  assertProperty(
    fc.property(
      fc.record({
        seed: fc.integer().map(String),
        steps: fc.nat().map(String),
        runs: fc.nat().map(String),
        collisionRuns: fc.nat().map(String),
        swarm: fc.constantFrom("0", "1"),
        relations: fc.stringMatching(/^[a-zA-Z0-9_-]{1,30}$/u),
        relationsDepth: fc.nat().map(String),
        flow: fc.oneof(
          fc.string({ minLength: 1, maxLength: 80 }),
          fc.stringMatching(/^[a-zA-Z0-9_-]{1,20}\.json$/u),
        ),
      }),
      (settings) => {
        const command = consumerScenarioReplay({
          inputs: settings,
          files: ["fuzz.test.ts"],
        });
        expect(command).toStartWith("bun ");
        expect(command).not.toContain("FOLIO_SCENARIO_");
        const parsed = parseConsumerArgs(argvOf(command), "/repo");
        expect(parsed).toMatchObject({ ...settings, files: ["fuzz.test.ts"] });
      },
    ),
    { numRuns: 100, id: "generated scenario replay settings round-trip" },
  );
});

test("repeated replay flags are rejected", () => {
  expect(() => parseConsumerArgs(["--seed", "1", "--seed", "2"], "/repo")).toThrow(
    "consumer-scenarios: repeated replay flag --seed",
  );
});

test("flow replay inputs override overlapping run inputs", () => {
  const runInputs = { seed: "792792449", runs: "12", collisionRuns: "8" };
  const flowInputs = { seed: "20260933", runs: "0", collisionRuns: "0" };
  const command = consumerScenarioReplay({ inputs: { ...runInputs, ...flowInputs } });

  expect(command.match(/--seed/gu)).toHaveLength(1);
  expect(command.match(/--runs/gu)).toHaveLength(1);
  expect(command.match(/--collision-runs/gu)).toHaveLength(1);
  expect(parseConsumerArgs(argvOf(command), "/repo")).toMatchObject(flowInputs);
});

test("merged run and flow replay inputs round-trip without repeated flags", () => {
  const inputNames = CONSUMER_SCENARIO_REPLAY_INPUTS.map(({ flag }) =>
    flag.slice(2).replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase()),
  );
  const inputs = fc.dictionary(
    fc.constantFrom(...inputNames),
    fc.stringMatching(/^[a-zA-Z0-9_-]{1,30}$/u),
  );
  assertProperty(
    fc.property(inputs, inputs, (runInputs, flowInputs) => {
      const merged = { ...runInputs, ...flowInputs };
      const command = consumerScenarioReplay({ inputs: merged });
      const argv = argvOf(command);
      const flags = argv.filter((argument) => argument.startsWith("--"));
      expect(new Set(flags).size).toBe(flags.length);
      const expected = Object.fromEntries(
        Object.entries(merged).map(([option, value]) => {
          const input = CONSUMER_SCENARIO_REPLAY_INPUTS.find(
            ({ flag }) =>
              flag
                .slice(2)
                .replaceAll(/-([a-z])/gu, (_, letter: string) => letter.toUpperCase()) === option,
          );
          return [option, input?.kind === "path" ? path.resolve(value) : value];
        }),
      );
      expect(consumerScenarioInputs(parseConsumerArgs(argv, "/repo"), {})).toMatchObject(expected);
    }),
    { numRuns: 100, id: "merged replay inputs round-trip once" },
  );
});

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
