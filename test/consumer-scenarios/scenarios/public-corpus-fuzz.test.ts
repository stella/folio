/** Full consumer flows from pinned public documents, explicitly enabled in CI. */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { test } from "node:test";

import {
  failureMarker,
  failureRecord,
  logFailureMarker,
  shellQuote,
  writeFailureRecord,
} from "../support/failure-fingerprints.ts";
import { recordFailure, relationEnv } from "../support/fuzz-loop.ts";
import { FlowError, runFlow } from "../support/fuzz.ts";
import {
  PUBLIC_CORPUS_PREFIX,
  PUBLIC_CORPUS_DIRECTORY_ENV,
  publicCorpusFixtures,
} from "../support/public-corpus.ts";

const ENABLED = process.env["FOLIO_SCENARIO_PUBLIC_CORPUS_FUZZ"] === "1";
if (!ENABLED) {
  test(
    "full flows from public corpus documents",
    { skip: "Enable FOLIO_SCENARIO_PUBLIC_CORPUS_FUZZ=1 with a staged cache" },
    () => {},
  );
} else {
  const seed = Number(process.env["FOLIO_SCENARIO_SEED"]);
  const steps = Number(process.env["FOLIO_SCENARIO_PUBLIC_CORPUS_STEPS"] ?? "12");
  const output = process.env["FOLIO_SCENARIO_PUBLIC_CORPUS_OUTPUT"];
  const fixtureFilter = process.env["FOLIO_SCENARIO_PUBLIC_CORPUS_FIXTURE"];
  if (
    !Number.isSafeInteger(seed) ||
    seed < 0 ||
    seed >= 2 ** 32 ||
    !Number.isInteger(steps) ||
    steps < 8 ||
    steps > 16 ||
    !output
  ) {
    throw new TypeError(
      "Public corpus fuzz requires a uint32 seed, 8–16 steps, and FOLIO_SCENARIO_PUBLIC_CORPUS_OUTPUT",
    );
  }
  const fixtures = await publicCorpusFixtures();
  if (fixtureFilter && !fixtures.some((fixture) => fixture === fixtureFilter)) {
    throw new TypeError("Replay fixture is absent from the staged public corpus index");
  }
  for (const fixture of fixtures) {
    if (fixtureFilter && fixture !== fixtureFilter) continue;
    test(`public corpus full flow: ${fixture}`, { timeout: 180_000 }, async () => {
      const destination = path.join(output, fixture.slice(PUBLIC_CORPUS_PREFIX.length));
      await mkdir(destination, { recursive: true });
      const repro = [
        `FOLIO_SCENARIO_PUBLIC_CORPUS_FUZZ=1 FOLIO_SCENARIO_SEED=${seed} FOLIO_SCENARIO_PUBLIC_CORPUS_STEPS=${steps}`,
        `FOLIO_SCENARIO_PUBLIC_CORPUS_OUTPUT=test-results/public-corpus FOLIO_SCENARIO_PUBLIC_CORPUS_FIXTURE=${shellQuote(fixture)}`,
        relationEnv(),
        "bun scripts/consumer-scenarios.ts -- public-corpus-fuzz.test.ts",
      ]
        .filter(Boolean)
        .join(" ");
      const stagedDirectory = process.env[PUBLIC_CORPUS_DIRECTORY_ENV];
      if (!stagedDirectory) throw new TypeError("Public corpus cache directory is missing");
      const cache = path.dirname(stagedDirectory);
      const prepare = `FOLIO_CORPUS_CACHE=${shellQuote(cache)} bun run corpus:fetch && FOLIO_CORPUS_CACHE=${shellQuote(cache)} bun scripts/stage-public-corpus-flows.ts`;
      await writeFile(
        path.join(destination, "replay.txt"),
        `# Requires the same corpus/sources.lock.json; hydrate its cache first:\n` +
          `${prepare}\n${repro}\n`,
      );
      try {
        const result = await runFlow(seed, steps, "random", { fixture });
        await writeFile(
          path.join(destination, "completed-flow.json"),
          `${JSON.stringify(result.flow)}\n`,
        );
      } catch (error) {
        if (error instanceof FlowError) {
          await writeFile(
            path.join(destination, "original-flow.json"),
            `${JSON.stringify(error.flow)}\n`,
          );
          const record = await recordFailure(
            error,
            { seed, repro },
            { maxAttempts: 30, seconds: 60 },
          );
          writeFailureRecord(destination, record);
        } else {
          const marker = failureMarker({
            test: `public corpus fixture ${fixture}`,
            seed,
            repro,
            failure: error,
          });
          logFailureMarker(marker);
          writeFailureRecord(destination, failureRecord(marker, error));
        }
        throw error;
      }
    });
  }
}
