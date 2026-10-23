/**
 * Flow files (support/flow-file.ts) replayed step by step: the checked-in
 * regression corpus in test/consumer-scenarios/flows, each of which must
 * pass with every step acting (support/fuzz.ts `vacuousSteps`), or, when
 * FOLIO_SCENARIO_FLOW holds a flow file's JSON (the replay line a shrunk
 * failure prints), that one flow alone. support/known-issues.ts lists the
 * checked-in flows that reproduce a finding and those with steps that act
 * on nothing.
 */

import assert from "node:assert/strict";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { readFlowDir } from "../support/corpus.ts";
import { reportScenarioFailure } from "../support/failure-fingerprints.ts";
import { type FlowFile, parseFlowFile } from "../support/flow-file.ts";
import {
  flowReplay,
  flowTestName,
  relationInputs,
} from "../support/fuzz-loop.ts";
import {
  type FlowRun,
  runFlowFile,
  unstableFixtureRefs,
  vacuousSteps,
} from "../support/fuzz.ts";
import {
  expectedFailure,
  FINDING_SYMPTOMS,
  KNOWN_FAILING_CHECKED_IN_FLOWS,
  VACUOUS_CHECKED_IN_FLOWS,
} from "../support/known-issues.ts";
import { consumerScenarioReplay } from "../support/replay-inputs.ts";

const FLOWS_DIR = fileURLToPath(new URL("../flows/", import.meta.url));

type FlowSource = { type: "requested" } | { type: "checked-in"; file: string };
type ReplayOptions = { flow: FlowFile; repro: string; source: FlowSource };

const replay =
  ({ flow, repro, source }: ReplayOptions) =>
  async () => {
    let run: FlowRun;
    try {
      run = await runFlowFile(flow);
    } catch (error) {
      reportScenarioFailure({
        test: flowTestName(flow),
        seed: flow.seed,
        repro,
        failure: error,
      });
      return;
    }
    switch (source.type) {
      case "requested":
        return;
      case "checked-in": {
        // A checked-in flow guards something only while its steps act.
        const known = VACUOUS_CHECKED_IN_FLOWS[source.file] ?? [];
        assert.deepEqual(
          vacuousSteps(run),
          known,
          `${repro}\nthe steps that applied nothing or found nothing to act on are not the ones support/known-issues.ts VACUOUS_CHECKED_IN_FLOWS lists; a flow whose steps act on nothing no longer reaches what it guards`,
        );
        return;
      }
      default: {
        const unhandled: never = source;
        throw new Error(`Unhandled flow source ${JSON.stringify(unhandled)}`);
      }
    }
  };

test("flow files retain explicit tool coverage and reject misplaced options", () => {
  const step = {
    action: "suggest_changes",
    seed: 1,
    operationTypes: ["commentOnBlock", "deleteBlock", "insertTable"],
    expectedApplied: 3,
    operations: [
      {
        type: "commentOnBlock",
        blockId: "12345678",
        comment: { text: "Note" },
      },
      { type: "deleteBlock", blockId: "12345678" },
      { type: "insertTable", blockId: "23456789", rows: [["Term", "Value"]] },
    ],
  };
  const flow = {
    version: 1,
    kind: "random",
    generation: "targeted",
    fixture: "comments",
    mode: "direct",
    seed: 1,
    steps: [step],
  };
  assert.deepEqual(parseFlowFile(flow).steps, [step]);
  for (const invalid of [
    { ...step, action: "core batch" },
    { ...step, operationTypes: ["commentOnBlock", "commentOnBlock"] },
    { ...step, expectedApplied: 0 },
  ])
    assert.throws(
      () => parseFlowFile({ ...flow, steps: [invalid] }),
      TypeError,
    );
});

const requested = process.env["FOLIO_SCENARIO_FLOW"];

if (requested !== undefined && requested !== "") {
  const flow = parseFlowFile(JSON.parse(requested));
  test(
    `replay FOLIO_SCENARIO_FLOW: ${flow.fixture} / ${flow.mode}, ${flow.steps.length} steps`,
    replay({ flow, repro: flowReplay(flow), source: { type: "requested" } }),
  );
} else {
  const corpus = readFlowDir(FLOWS_DIR);
  const files = new Set(corpus.map(({ file }) => file));
  test("support/known-issues.ts lists only checked-in flow files", () => {
    const listed = [
      ...Object.keys(KNOWN_FAILING_CHECKED_IN_FLOWS),
      ...Object.keys(VACUOUS_CHECKED_IN_FLOWS),
    ];
    assert.deepEqual(
      listed.filter((file) => !files.has(file)),
      [],
    );
  });
  test("checked-in flows name fixture blocks only by the paraIds fixture models pin", async () => {
    // A fixture block id that came from serialization or a build-time edit
    // changes without the fixture changing, and the flow then acts on nothing.
    const unstable: Record<string, string[]> = {};
    for (const { file, flow } of corpus) {
      const refs = await unstableFixtureRefs(flow);
      if (refs.length > 0) unstable[file] = refs;
    }
    assert.deepEqual(unstable, {});
  });
  for (const { file, flow } of corpus) {
    const name = `checked-in flow ${file}`;
    const repro = consumerScenarioReplay({
      inputs: relationInputs(),
      flags: [`--only '^${name.replaceAll(".", "\\.")}'`],
      files: ["flow-corpus.test.ts"],
    });
    const title = `${name}: ${flow.title ?? flow.origin ?? "regression flow"}`;
    const body = replay({ flow, repro, source: { type: "checked-in", file } });
    const finding = KNOWN_FAILING_CHECKED_IN_FLOWS[file];
    if (finding === undefined) test(title, body);
    else expectedFailure(finding, title, FINDING_SYMPTOMS[finding], body);
  }
}
