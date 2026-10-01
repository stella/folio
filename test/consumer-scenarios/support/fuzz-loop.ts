/**
 * Continuous fuzzing: for a time budget, run fresh seeded flows and mutants
 * of corpus flows, keep every flow that reaches a state the corpus has not
 * seen, and shrink each new failure to a minimal flow file with a one-line
 * replay. A few flows that reach a new document structure or step outcome
 * are also shrunk to the fewest steps that still reach it, as candidates for
 * the checked-in corpus. Driven by scenarios/continuous-fuzz.test.ts.
 */

import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";

import { admit, loadCorpus, prune } from "./corpus.ts";
import {
  failureMarker,
  failureRecord,
  type FailureRecord,
  logFailureMarker,
  shellQuote,
  writeFailureRecord,
} from "./failure-fingerprints.ts";
import { compactFlow, type FlowFile, flowId, type FlowKind, flowShape } from "./flow-file.ts";
import { mutateFlow } from "./flow-mutate.ts";
import { describeFlow, FlowError, type FlowRun, runFlow, runFlowFile } from "./fuzz.ts";
import { createRandom } from "./random.ts";
import { shrinkFlow } from "./shrink.ts";

/** The relation settings a replay needs to fail the same way, as `NAME=value` words. */
export const relationEnv = (): string =>
  ["FOLIO_SCENARIO_RELATIONS", "FOLIO_SCENARIO_RELATIONS_DEPTH"]
    .flatMap((name) => {
      const value = process.env[name];
      return value === undefined ? [] : [`${name}=${shellQuote(value)}`];
    })
    .join(" ");

const words = (...parts: string[]): string => parts.filter((part) => part !== "").join(" ");

/** The one line that replays a seeded flow through scenarios/fuzz.test.ts. */
export const seedReplay = (seed: number, steps: number, kind: FlowKind): string => {
  const [runs, label] =
    kind === "random"
      ? ["FOLIO_SCENARIO_FUZZ_RUNS=1 FOLIO_SCENARIO_COLLISION_RUNS=0", "fuzz"]
      : ["FOLIO_SCENARIO_FUZZ_RUNS=0 FOLIO_SCENARIO_COLLISION_RUNS=1", "collision"];
  return words(
    `FOLIO_SCENARIO_SEED=${seed} FOLIO_SCENARIO_FUZZ_STEPS=${steps} ${runs}`,
    relationEnv(),
    process.env["FOLIO_SCENARIO_SWARM"] === "1" ? "FOLIO_SCENARIO_SWARM=1" : "",
    `bun scripts/consumer-scenarios.ts --only '^${label} run 0 \\(' -- fuzz.test.ts`,
  );
};

/** The one line that replays a flow file through scenarios/flow-corpus.test.ts. */
export const flowReplay = (flow: FlowFile): string =>
  words(
    `FOLIO_SCENARIO_FLOW=${shellQuote(compactFlow(flow))}`,
    relationEnv(),
    "bun scripts/consumer-scenarios.ts --only '^replay FOLIO_SCENARIO_FLOW' -- flow-corpus.test.ts",
  );

/** The test name a flow's failures are fingerprinted under (as in fuzz.test.ts). */
export const flowTestName = (flow: Pick<FlowFile, "fixture" | "mode">): string =>
  `consumer flow ${flow.fixture} / ${flow.mode}`;

const fingerprintOf = (flow: FlowFile, failure: unknown): string =>
  failureMarker({ test: flowTestName(flow), seed: 0, repro: "", failure }).fingerprint;

/** Whether `flow` fails with `fingerprint`. */
const failsWith = async (flow: FlowFile, fingerprint: string): Promise<boolean> => {
  try {
    await runFlowFile(flow);
    return false;
  } catch (error) {
    return error instanceof FlowError && fingerprintOf(flow, error) === fingerprint;
  }
};

/** `flow` with every batch step pinned to what it applied. */
const materialize = async (flow: FlowFile): Promise<FlowFile> => {
  try {
    return (await runFlowFile(flow)).flow;
  } catch (error) {
    if (error instanceof FlowError) return error.flow;
    throw error;
  }
};

export type ShrinkLimits = { maxAttempts: number; seconds: number };

/**
 * The record for a failed flow: its marker, and, when the failing flow
 * replays from its file, the minimized flow and a replay of it first.
 */
export const recordFailure = async (
  error: FlowError,
  { seed, repro }: { seed: number; repro: string },
  limits: ShrinkLimits | null,
): Promise<FailureRecord> => {
  const test = flowTestName(error.flow);
  const primary = failureMarker({ test, seed, repro, failure: error });
  if (limits === null || !(await failsWith(error.flow, primary.fingerprint))) {
    logFailureMarker(primary);
    return failureRecord(primary, error);
  }
  const shrunk = await shrinkFlow(error.flow, {
    holds: (candidate) => failsWith(candidate, primary.fingerprint),
    materialize,
    budget: { maxAttempts: limits.maxAttempts, deadline: Date.now() + limits.seconds * 1_000 },
  });
  // What the minimized flow does is part of which bug this is.
  const marker = failureMarker({ test, seed, repro, failure: error, flow: flowShape(shrunk.flow) });
  logFailureMarker(marker);
  const replay = flowReplay(shrunk.flow);
  return failureRecord(marker, error, {
    replays: replay === repro ? [replay] : [replay, repro],
    flow: shrunk.flow,
    shrink: { steps: shrunk.flow.steps.length, from: shrunk.from, attempts: shrunk.attempts },
  });
};

/** Signature elements a checked-in flow is worth keeping for: not coverage cells alone. */
const notable = (elements: readonly string[]): string[] =>
  elements.filter((element) => !element.startsWith("cell "));

export type MinimizeLimits = { dir: string; max: number; seconds: number };

/** Entries a corpus needs before a flow's new elements say something. */
const WARM_CORPUS = 20;

/**
 * `flow` shrunk to the fewest steps that still pass and reach every one of
 * `elements`, or null when a replay of it does not.
 */
export const minimizeReaching = async (
  flow: FlowFile,
  elements: readonly string[],
  seconds: number,
): Promise<FlowFile | null> => {
  const reaches = async (candidate: FlowFile): Promise<boolean> => {
    try {
      const { signature } = await runFlowFile(candidate, { signature: true });
      return elements.every((element) => signature.includes(element));
    } catch {
      return false;
    }
  };
  if (!(await reaches(flow))) return null;
  const shrunk = await shrinkFlow(flow, {
    holds: reaches,
    materialize: async (candidate) => {
      try {
        return (await runFlowFile(candidate)).flow;
      } catch {
        return candidate;
      }
    },
    budget: { maxAttempts: 80, deadline: Date.now() + seconds * 1_000 },
  });
  return { ...shrunk.flow, title: `reaches ${elements.join("; ")}` };
};

export type LoopOptions = {
  seed: number;
  seconds: number;
  /** Steps of a fresh flow. */
  steps: number;
  /** Longest mutant. */
  maxSteps: number;
  /** Share of runs that mutate a corpus flow (once the corpus has one). */
  mutateShare: number;
  corpusDir: string | null;
  /** Kept at most this many cached entries. */
  corpusSize: number;
  seeds: readonly FlowFile[];
  failuresDir: string | null;
  shrink: ShrinkLimits | null;
  /** Where to write minimized flows that reach a new structure or outcome, and how many. */
  minimize: MinimizeLimits | null;
  log?: (line: string) => void;
};

export type LoopResult = {
  flows: number;
  mutants: number;
  admitted: number;
  corpus: number;
  minimized: number;
  failures: { record: FailureRecord; count: number }[];
};

/** Fuzz until the budget runs out; see the module comment. */
export const fuzzFor = async (options: LoopOptions): Promise<LoopResult> => {
  const log = options.log ?? ((line: string) => console.log(line));
  const deadline = Date.now() + options.seconds * 1_000;
  // Shrinking may run past the budget by one failure's shrink time, no more.
  const hardStop = deadline + (options.shrink?.seconds ?? 0) * 1_000;
  const shrinkLimits = (): ShrinkLimits | null => {
    const left = Math.floor((hardStop - Date.now()) / 1_000);
    return options.shrink === null || left <= 0
      ? null
      : { ...options.shrink, seconds: Math.min(options.shrink.seconds, left) };
  };
  const corpus = loadCorpus(options.corpusDir, options.seeds);
  const control = createRandom(options.seed ^ 0x6c6f_6f70);
  const failures = new Map<string, { record: FailureRecord; count: number }>();
  const result: LoopResult = {
    flows: 0,
    mutants: 0,
    admitted: 0,
    corpus: 0,
    minimized: 0,
    failures: [],
  };
  log(
    `continuous fuzz: ${options.seconds}s from seed ${options.seed}, corpus of ${corpus.entries.length} flows`,
  );
  for (let run = 0; Date.now() < deadline; run += 1) {
    const seed = options.seed + run;
    const mutate = corpus.entries.length > 0 && control.chance(options.mutateShare);
    let repro = "";
    let label = `run ${run}`;
    try {
      let running: Promise<FlowRun>;
      if (mutate) {
        const parent = control.pick(corpus.entries);
        const { flow, mutations } = mutateFlow(parent.flow, createRandom(seed), {
          donors: corpus.entries.map((entry) => entry.flow),
          maxSteps: options.maxSteps,
        });
        repro = flowReplay(flow);
        label = `mutant ${seed} (${mutations.join("+")} of ${parent.id}): ${flow.fixture} / ${flow.mode}, ${flow.steps.length} steps`;
        result.mutants += 1;
        running = runFlowFile(flow, { signature: true });
      } else {
        const kind: FlowKind = run % 2 === 0 ? "random" : "collisions";
        const { fixture, mode } = describeFlow(seed, kind);
        repro = seedReplay(seed, options.steps, kind);
        label = `${kind} flow ${seed}: ${fixture} / ${mode}, ${options.steps} steps`;
        result.flows += 1;
        running = runFlow(seed, options.steps, kind, { signature: true });
      }
      const ran = await running;
      // Against a cold corpus everything is new; only a warm one says what is.
      const warm = corpus.entries.length >= WARM_CORPUS;
      const fresh = admit(corpus, ran.flow, ran.signature);
      if (fresh.length > 0) result.admitted += 1;
      log(`✓ ${label}${fresh.length > 0 ? `; new: ${fresh.length}` : ""}`);
      const [target] = notable(fresh);
      if (
        options.minimize !== null &&
        warm &&
        result.minimized < options.minimize.max &&
        target !== undefined &&
        Date.now() < deadline
      ) {
        const minimized = await minimizeReaching(ran.flow, [target], options.minimize.seconds);
        if (minimized !== null) {
          result.minimized += 1;
          mkdirSync(options.minimize.dir, { recursive: true });
          const file = join(options.minimize.dir, `${flowId(minimized)}.json`);
          writeFileSync(file, `${JSON.stringify(minimized, null, 2)}\n`);
          log(`  minimized to ${minimized.steps.length} steps: ${file}`);
        }
      }
    } catch (error) {
      if (!(error instanceof FlowError)) throw error;
      const fingerprint = fingerprintOf(error.flow, error);
      const seen = failures.get(fingerprint);
      log(`✗ ${label} (fingerprint ${fingerprint})\n  Replay: ${repro}`);
      if (seen !== undefined) {
        // The same failure again: its line carries the shrunk one's fingerprint.
        seen.count += 1;
        logFailureMarker({ ...seen.record.marker, seed, repro });
        continue;
      }
      const record = await recordFailure(error, { seed, repro }, shrinkLimits());
      failures.set(fingerprint, { record, count: 1 });
      if (options.failuresDir !== null) writeFailureRecord(options.failuresDir, record);
      log(
        record.shrink === undefined
          ? "  not shrunk (no time left, or its flow file does not replay it)"
          : `  shrunk from ${record.shrink.from} to ${record.shrink.steps} steps; replay: ${record.replays[0]}`,
      );
    }
  }
  prune(corpus, options.corpusSize);
  result.corpus = corpus.entries.length;
  result.failures = [...failures.values()];
  return result;
};
