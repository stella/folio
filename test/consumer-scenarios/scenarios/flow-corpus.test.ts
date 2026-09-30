/**
 * Flow files (support/flow-file.ts) replayed step by step: the checked-in
 * regression corpus in test/consumer-scenarios/flows, each of which must
 * pass, or, when FOLIO_SCENARIO_FLOW holds a flow file's JSON (the replay
 * line a shrunk failure prints), that one flow alone.
 */

import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { readFlowDir } from "../support/corpus.ts";
import { reportScenarioFailure } from "../support/failure-fingerprints.ts";
import { type FlowFile, parseFlowFile } from "../support/flow-file.ts";
import { flowReplay, flowTestName, relationEnv } from "../support/fuzz-loop.ts";
import { runFlowFile } from "../support/fuzz.ts";

const FLOWS_DIR = fileURLToPath(new URL("../flows/", import.meta.url));

const replay = (flow: FlowFile, repro: string) => async () => {
  try {
    await runFlowFile(flow);
  } catch (error) {
    reportScenarioFailure({ test: flowTestName(flow), seed: flow.seed, repro, failure: error });
  }
};

const requested = process.env["FOLIO_SCENARIO_FLOW"];

if (requested !== undefined && requested !== "") {
  const flow = parseFlowFile(JSON.parse(requested));
  test(
    `replay FOLIO_SCENARIO_FLOW: ${flow.fixture} / ${flow.mode}, ${flow.steps.length} steps`,
    replay(flow, flowReplay(flow)),
  );
} else {
  for (const { file, flow } of readFlowDir(FLOWS_DIR)) {
    const name = `checked-in flow ${file}`;
    const repro = [
      relationEnv(),
      `bun scripts/consumer-scenarios.ts --only '^${name.replaceAll(".", "\\.")}' -- flow-corpus.test.ts`,
    ]
      .filter((part) => part !== "")
      .join(" ");
    test(`${name}: ${flow.title ?? flow.origin ?? "regression flow"}`, replay(flow, repro));
  }
}
