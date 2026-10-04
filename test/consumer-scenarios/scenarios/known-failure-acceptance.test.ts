/** Exact replays for known consumer failures that remain parked. */

import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { test } from "node:test";

import { FlowError, runFlowFile, unstableFixtureRefs } from "../support/fuzz.ts";
import { failureMarker } from "../support/failure-fingerprints.ts";
import { flowShape, parseFlowFile, type FlowFile } from "../support/flow-file.ts";

type ConsumerAcceptance = {
  fingerprint: string;
  issue: number;
  primary: string;
  reportSeed: number;
  title: string;
  flow: FlowFile;
  causeMessage?: string;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const registryPath =
  process.env["FOLIO_SCENARIO_KNOWN_FAILURE_REGISTRY"] ??
  fileURLToPath(new URL("../../known-failure-fingerprints.json", import.meta.url));
const registry: unknown = JSON.parse(readFileSync(registryPath, "utf8"));

const readAcceptanceEntries = (value: unknown): ConsumerAcceptance[] => {
  if (!isRecord(value) || !Array.isArray(value["known"])) {
    throw new TypeError("known failure registry must contain a known array");
  }
  const entries = value["known"];
  const acceptanceEntries: ConsumerAcceptance[] = [];
  for (const candidate of entries) {
    if (!isRecord(candidate) || !("acceptance" in candidate)) continue;
    const metadata = candidate["acceptance"];
    if (!isRecord(metadata)) {
      throw new TypeError("known failure acceptance metadata must be an object");
    }
    if (metadata["type"] === "browser") continue;
    if (metadata["type"] !== "consumer") {
      throw new TypeError("known failure acceptance metadata has an unknown type");
    }
    const { fingerprint, issueOrPr, firstSeen } = candidate;
    const { issue, primary, reportSeed, title, flow: flowValue, causeMessage } = metadata;
    if (
      typeof fingerprint !== "string" ||
      typeof issueOrPr !== "string" ||
      typeof firstSeen !== "string" ||
      typeof issue !== "number" ||
      typeof primary !== "string" ||
      typeof reportSeed !== "number" ||
      typeof title !== "string" ||
      (causeMessage !== undefined && typeof causeMessage !== "string")
    ) {
      throw new TypeError("known failure acceptance entry is malformed");
    }
    const primaryEntry = entries.find(
      (entry) =>
        isRecord(entry) &&
        entry["fingerprint"] === primary &&
        entry["issueOrPr"] === issueOrPr &&
        entry["firstSeen"] === firstSeen,
    );
    assert.ok(primaryEntry, `#${issue} primary fingerprint is absent from the registry`);
    acceptanceEntries.push({
      fingerprint,
      issue,
      primary,
      reportSeed,
      title,
      flow: parseFlowFile(flowValue),
      ...(causeMessage === undefined ? {} : { causeMessage }),
    });
  }
  return acceptanceEntries;
};

const acceptanceEntries = readAcceptanceEntries(registry);

assert.equal(
  new Set(acceptanceEntries.map(({ fingerprint }) => fingerprint)).size,
  acceptanceEntries.length,
  "each parked consumer fingerprint must have exactly one acceptance replay",
);

test("acceptance flows name fixture blocks only by the paraIds fixture models pin", async () => {
  const unstable: Record<string, string[]> = {};
  for (const { fingerprint, flow } of acceptanceEntries) {
    const refs = await unstableFixtureRefs(flow);
    if (refs.length > 0) unstable[fingerprint] = refs;
  }
  assert.deepEqual(unstable, {});
});

for (const {
  fingerprint,
  issue,
  primary,
  reportSeed,
  title,
  flow,
  causeMessage,
} of acceptanceEntries) {
  const testName = `consumer flow ${flow.fixture} / ${flow.mode}`;
  test(`known failure #${issue} ${fingerprint} (reported seed ${reportSeed}): ${title}`, async () => {
    let failure: unknown;
    try {
      await runFlowFile(flow);
    } catch (error) {
      failure = error;
    }
    assert.notEqual(failure, undefined, `#${issue} acceptance flow unexpectedly passed`);
    assert.ok(
      failure instanceof FlowError,
      `#${issue} acceptance replay failed outside the flow: ${String(failure)}`,
    );
    if (causeMessage !== undefined) {
      let cause: Error = failure;
      const seen = new Set<Error>([cause]);
      while (cause.cause instanceof Error && !seen.has(cause.cause)) {
        cause = cause.cause;
        seen.add(cause);
      }
      assert.equal(cause.message, causeMessage, `#${issue} exact failure cause changed`);
    }

    const marker = failureMarker({
      test: testName,
      seed: flow.seed,
      repro: "",
      failure,
      flow: flowShape(flow),
    });
    assert.equal(marker.primary, primary, `#${issue} no longer reproduces its recorded failure`);
    assert.equal(marker.fingerprint, fingerprint, `#${issue} minimized flow fingerprint changed`);
  });
}
