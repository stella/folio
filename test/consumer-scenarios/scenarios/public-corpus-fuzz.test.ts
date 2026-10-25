/** Full consumer flows from pinned public documents, explicitly enabled in CI. */
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { test } from "node:test";

import { openReviewer } from "../support/documents.ts";
import {
  failureMarker,
  failureRecord,
  logFailureMarker,
  shellQuote,
  writeFailureRecord,
} from "../support/failure-fingerprints.ts";
import type { FlowFile } from "../support/flow-file.ts";
import { recordFailure, relationInputs } from "../support/fuzz-loop.ts";
import { FlowError, runFlow, runFlowFile } from "../support/fuzz.ts";
import { assertReadersAgree, saveAndReopen } from "../support/invariants.ts";
import { startRelations } from "../support/metamorphic.ts";
import {
  loadPublicCorpusFixture,
  PUBLIC_CORPUS_DIRECTORY_ENV,
  PUBLIC_CORPUS_PREFIX,
  publicCorpusFixtures,
} from "../support/public-corpus.ts";
import { consumerScenarioReplay } from "../support/replay-inputs.ts";

const ENABLED = process.env["FOLIO_SCENARIO_PUBLIC_CORPUS_FUZZ"] === "1";
if (!ENABLED) {
  test(
    "full flows from public corpus documents",
    { skip: "Enable FOLIO_SCENARIO_PUBLIC_CORPUS_FUZZ=1 with a staged cache" },
    () => {},
  );
} else {
  test("pinned seed 889732909: merging before an empty nested bullet", async () => {
    await runFlowFile({
      version: 1,
      kind: "random",
      generation: "targeted",
      fixture:
        "public-corpus:003e08fc366ccc3f36f621b32f4c897f4e42b00853d0311a6b56b3e36423e1d6",
      mode: "tracked-changes",
      seed: 889732909,
      steps: [
        {
          action: "story batch",
          seed: 2089463052,
          operations: [
            { type: "mergeBlockWithNext", blockId: "08D286CB", separator: " " },
          ],
          atomic: false,
        },
      ],
      origin: "shrunk from random flow seed 889732909",
    });
  });

  test("pinned #1520 seed 1634001732: restyling a heading with inline breaks", async () => {
    await runFlowFile({
      version: 1,
      kind: "random",
      generation: "targeted",
      fixture:
        "public-corpus:00e76a5f41aec4c76133b89312337952016bc73ff12b37c29d5aa07026c35d51",
      mode: "direct",
      seed: 1634001732,
      steps: [
        {
          action: "suggest_changes",
          seed: 938224865,
          operations: [
            {
              type: "setBlockParagraphProperties",
              blockId: "021CE00A",
              properties: { styleId: "Heading2" },
            },
          ],
        },
      ],
      origin: "shrunk from random flow seed 1634001732",
    });
  });

  for (const [title, operation] of [
    [
      "inserting before a paragraph keeps its editing sessions on it",
      {
        type: "insertBeforeBlock",
        blockId: "441E8208",
        text: "Invoice schedule delivery party warranty party buyer.",
      },
    ],
    [
      "rejecting a split agrees with its save",
      { type: "splitBlock", blockId: "46227DB8", offset: 16 },
    ],
  ] as const) {
    test(`pinned #1762 seed 841387651: ${title}`, async () => {
      await runFlowFile({
        version: 1,
        kind: "random",
        generation: "targeted",
        fixture:
          "public-corpus:00e76a5f41aec4c76133b89312337952016bc73ff12b37c29d5aa07026c35d51",
        mode: "tracked-changes",
        seed: 841387651,
        steps: [
          {
            action: "core batch",
            seed: 2353247479,
            atomic: true,
            operations: [operation],
          },
        ],
        origin: "shrunk from random flow seed 841387651",
      });
    });
  }

  test("pinned #1811 seed 1988744118: an inserted list item does not repeat its anchor's page break", async () => {
    await runFlowFile({
      version: 1,
      kind: "random",
      generation: "targeted",
      fixture:
        "public-corpus:003e08fc366ccc3f36f621b32f4c897f4e42b00853d0311a6b56b3e36423e1d6",
      mode: "tracked-changes",
      seed: 1988744118,
      steps: [
        {
          action: "core batch",
          seed: 1101143660,
          atomic: true,
          operations: [
            {
              type: "insertAfterBlock",
              blockId: "15BA2F55",
              text: "Party invoice term written.",
              numbering: { kind: "newList", format: "bullet" },
            },
          ],
        },
      ],
      origin: "shrunk from random flow seed 1988744118",
    });
  });

  const READER_CASES = [
    {
      issue: 1366,
      fixture:
        "public-corpus:003e08fc366ccc3f36f621b32f4c897f4e42b00853d0311a6b56b3e36423e1d6",
    },
    {
      issue: 1367,
      fixture:
        "public-corpus:00e76a5f41aec4c76133b89312337952016bc73ff12b37c29d5aa07026c35d51",
    },
    {
      issue: 1369,
      fixture:
        "public-corpus:0108d86286c18600553c4b23d0bbe47d7e5737aca43d01eb24820c99d2cf1377",
    },
  ] as const;
  // PINNED seed 1772869001, reject-all step seed 2972599144.
  for (const { issue, fixture } of READER_CASES) {
    test(`pinned reader #${issue}, seed 1772869001`, async () => {
      const reviewer = await openReviewer(
        await loadPublicCorpusFixture(fixture),
      );
      reviewer.rejectAll();
      await assertReadersAgree(
        new Uint8Array(await reviewer.toBuffer()),
        `#${issue} reject all`,
      );
    });
  }

  const PACKAGE_IDENTITY_CASES = [
    {
      name: "#1368: rejecting all on a positional-ID document",
      flow: {
        version: 1,
        kind: "random",
        generation: "targeted",
        fixture:
          "public-corpus:0046e12f3a66e48dc3d27a2a534ae5cce848aa4497eaa076ebfb5c6fa807774d",
        mode: "suggested",
        seed: 1772869001,
        steps: [
          { action: "reject all", seed: 2972599144 },
        ] satisfies FlowFile["steps"],
        origin: "random flow seed 1772869001",
        title: "package identity #1368",
      } as const satisfies FlowFile,
    },
    {
      name: "#1370: rejecting all on a tracked positional-ID document",
      flow: {
        version: 1,
        kind: "random",
        generation: "targeted",
        fixture:
          "public-corpus:0058e2003402882807e50c1e66ac7e49eda88305c1aafd077d2efa9ae7d001fe",
        mode: "suggested",
        seed: 1772869001,
        steps: [
          { action: "reject all", seed: 2972599144 },
        ] satisfies FlowFile["steps"],
        origin: "random flow seed 1772869001",
        title: "package identity #1370",
      } as const satisfies FlowFile,
    },
  ] as const;

  for (const { name, flow } of PACKAGE_IDENTITY_CASES) {
    test(`pinned package identity ${name}`, async () => {
      const fixture = await loadPublicCorpusFixture(flow.fixture);
      const reviewer = await openReviewer(fixture);
      const relations = await startRelations({
        fixture,
        reviewer,
        mode: flow.mode,
        seed: flow.seed,
      });
      reviewer.rejectAll();
      const saved = await saveAndReopen(reviewer, name);
      await relations.afterStep(reviewer, saved, name);
    });
  }

  test("pinned numbering provenance #1453, seed 1634001732", async () => {
    await runFlowFile({
      version: 1,
      kind: "random",
      generation: "targeted",
      fixture:
        "public-corpus:003e08fc366ccc3f36f621b32f4c897f4e42b00853d0311a6b56b3e36423e1d6",
      mode: "direct",
      seed: 1634001732,
      steps: [
        {
          action: "suggest_changes",
          seed: 938224865,
          operations: [
            {
              type: "setBlockParagraphProperties",
              blockId: "43392F69",
              properties: { styleId: null },
            },
          ],
        },
      ],
      origin: "shrunk from random flow seed 1634001732",
    });
  });

  const seed = Number(process.env["FOLIO_SCENARIO_SEED"]);
  const steps = Number(
    process.env["FOLIO_SCENARIO_PUBLIC_CORPUS_STEPS"] ?? "12",
  );
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
    throw new TypeError(
      "Replay fixture is absent from the staged public corpus index",
    );
  }
  for (const fixture of fixtures) {
    if (fixtureFilter && fixture !== fixtureFilter) continue;
    test(
      `public corpus full flow: ${fixture}`,
      { timeout: 180_000 },
      async () => {
        const destination = path.join(
          output,
          fixture.slice(PUBLIC_CORPUS_PREFIX.length),
        );
        await mkdir(destination, { recursive: true });
        const repro = consumerScenarioReplay({
          inputs: {
            ...relationInputs(),
            publicCorpusFuzz: "1",
            seed: String(seed),
            publicCorpusSteps: String(steps),
            publicCorpusOutput: "test-results/public-corpus",
            publicCorpusFixture: fixture,
          },
          files: ["public-corpus-fuzz.test.ts"],
        });
        const stagedDirectory = process.env[PUBLIC_CORPUS_DIRECTORY_ENV];
        if (!stagedDirectory)
          throw new TypeError("Public corpus cache directory is missing");
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
      },
    );
  }
}
