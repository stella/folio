import { test, expect } from "@playwright/test";
import fc from "fast-check";
import {
  failureMarker,
  failureRecord,
  logFailureMarker,
  writeFailureRecord,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";

import {
  FolioDocxReviewer,
  type FolioDocumentOperationBatch,
} from "../../packages/core/src/server";
import { runFolioCli } from "../../packages/cli/src/cli";
import { fileVersionOf } from "../../packages/cli/src/document";
import { buildDocx, makeTempDir } from "../../packages/cli/src/__tests__/fixtures";
import { propertyConfig, propertyTestTimeout } from "../../test/property-testing";
import { openEditor } from "./parity-fixture";
import { assertGeneratedFlowReceipt } from "./flow-receipt";

// Authored paragraph ids address the original paragraphs through insertions and saves.
// Actions shrink independently; each replacement keeps its find token available.
const flowArbitrary = fc.record({
  paragraphs: fc.integer({ min: 2, max: 5 }),
  actions: fc.array(
    fc.record({
      kind: fc.constantFrom("replace", "insert"),
      target: fc.nat({ max: 4 }),
      text: fc.constantFrom("café", "東京", "e\u0301", "مرحبا", "👩🏽‍⚖️", "<&>"),
    }),
    { minLength: 1, maxLength: 8 },
  ),
});

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const semanticProjection = async (bytes: Uint8Array) => {
  expect(await validateDocxPackage(bytes)).toEqual({ valid: true });
  const reviewer = await FolioDocxReviewer.fromBuffer(new Uint8Array(bytes).buffer);
  return {
    // Exclude regenerated identity and package metadata, retain semantic structure.
    blocks: reviewer.snapshot().blocks.map(({ kind, text, displayLabel, listLevel, table }) => ({
      kind,
      text,
      displayLabel,
      listLevel,
      table,
    })),
    changes: reviewer.getChanges().map(({ type, text, author }) => ({ type, text, author })),
    comments: reviewer.getComments().map(({ text, author }) => ({ text, author })),
  };
};

const reactPort = Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200;
const vuePort = Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;

test("generated flow saves equivalent semantics in React, Vue, headless and CLI", async ({
  browser,
}, testInfo) => {
  test.setTimeout(propertyTestTimeout(240_000));
  const react = await browser.newPage();
  const vue = await browser.newPage();
  try {
    await openEditor(react, { name: "react", baseUrl: `http://localhost:${reactPort}` });
    await openEditor(vue, { name: "vue", baseUrl: `http://localhost:${vuePort}` });
    const verdict = await fc.check(
      fc.asyncProperty(flowArbitrary, async ({ paragraphs, actions }) => {
        const initial = Array.from({ length: paragraphs }, (_, index) => ({
          text: `token${index} original clause`,
          paraId: (0x10000001 + index).toString(16).toUpperCase(),
          ...(index === 0 ? { style: "Heading1" } : {}),
        }));
        const source = await buildDocx(initial);
        const expected = initial.map(({ text, paraId }) => ({ text, paraId }));
        const batches = actions.map(({ kind, target, text }, index) => {
          const targetIndex = target % paragraphs;
          const paraId = (0x10000001 + targetIndex).toString(16).toUpperCase();
          const id = `flow-${index}`;
          const token = `token${targetIndex}`;
          const currentIndex = expected.findIndex((entry) => entry.paraId === paraId);
          const current = expected.at(currentIndex);
          if (!current) throw new Error("Generated target disappeared");
          if (kind === "insert") {
            expected.splice(currentIndex + 1, 0, { text, paraId: `inserted-${index}` });
            return {
              version: 1,
              mode: "direct",
              operations: [{ id, type: "insertAfterBlock", blockId: paraId, text }],
            } as const satisfies FolioDocumentOperationBatch;
          }
          current.text = current.text.replace(token, `${token} ${text}`);
          return {
            version: 1,
            mode: "direct",
            operations: [
              {
                id,
                type: "replaceInBlock",
                blockId: paraId,
                find: token,
                replace: `${token} ${text}`,
              },
            ],
          } as const satisfies FolioDocumentOperationBatch;
        });
        const headless = await FolioDocxReviewer.fromBuffer(new Uint8Array(source).buffer);
        for (const batch of batches) {
          const operation = batch.operations[0];
          assertGeneratedFlowReceipt(headless.applyDocumentOperations(batch), operation);
        }
        const outputs = [{ name: "headless", bytes: new Uint8Array(await headless.toBuffer()) }];
        for (const [name, page] of [
          ["react", react],
          ["vue", vue],
        ] as const) {
          const output = await page.evaluate(
            async ({ source: inputBytes, batches: inputBatches }) => {
              const bridge = window.__folioParity;
              if (!bridge) throw new Error("Parity bridge unavailable");
              return await bridge.runGeneratedFlow(inputBytes, inputBatches);
            },
            { source: [...source], batches },
          );
          output.results.forEach((result, index) => {
            const batch = batches.at(index);
            if (!batch) throw new Error("Missing generated flow batch");
            assertGeneratedFlowReceipt(result, batch.operations[0]);
          });
          outputs.push({ name, bytes: new Uint8Array(output.bytes) });
        }
        const { dir, cleanup } = await makeTempDir();
        try {
          const file = path.join(dir, "flow.docx");
          await writeFile(file, source);
          for (const batch of batches) {
            const before = new Uint8Array(await readFile(file));
            const stdout: string[] = [];
            const stderr: string[] = [];
            const exit = await runFolioCli(
              [
                "suggest",
                file,
                "--input",
                JSON.stringify({ operations: batch.operations }),
                "--direct",
                "--in-place",
                "--expect-version",
                fileVersionOf(before),
                "--author",
                "Parity",
                "--date",
                "2026-01-02T03:04:05Z",
                "--allow-repack",
              ],
              {
                stdout: (text) => stdout.push(text),
                stderr: (text) => stderr.push(text),
                readStdin: async () => "",
                isTTY: false,
                env: {},
                cwd: dir,
              },
            );
            expect(exit, stderr.join("")).toBe(0);
            const envelope: unknown = JSON.parse(stdout.join(""));
            if (!isRecord(envelope) || !isRecord(envelope["data"])) {
              throw new Error("CLI returned no success receipt");
            }
            expect(envelope["ok"]).toBe(true);
            expect(envelope["data"]["status"]).toBe("committed");
            assertGeneratedFlowReceipt(envelope["data"]["result"], batch.operations[0]);
            expect(fileVersionOf(new Uint8Array(await readFile(file)))).not.toBe(
              fileVersionOf(before),
            );
          }
          outputs.push({ name: "cli", bytes: new Uint8Array(await readFile(file)) });
        } finally {
          await cleanup();
        }
        const baseline = await semanticProjection(new Uint8Array(await headless.toBuffer()));
        expect(baseline.blocks.map(({ text }) => text)).toEqual(expected.map(({ text }) => text));
        expect(baseline.changes).toEqual([]);
        for (const output of outputs) {
          const projection = await semanticProjection(output.bytes);
          expect(projection, output.name).toEqual(baseline);
          expect(
            projection.blocks.map(({ text }) => text),
            output.name,
          ).toEqual(expected.map(({ text }) => text));
        }
      }),
      propertyConfig({
        numRuns: 5,
        endOnFailure: false,
        // PINNED #1407: seed 501525673, path 1:0:1:2:2:2:2:2:2.
        examples: [
          [
            {
              paragraphs: 2,
              actions: [
                { kind: "insert", target: 0, text: "café" },
                { kind: "replace", target: 0, text: "café" },
              ],
            },
          ],
        ],
      }),
    );
    if (!verdict.failed) return;
    const failure = {
      seed: verdict.seed,
      path: verdict.counterexamplePath,
      trace: verdict.counterexample?.at(0),
      error: String(verdict.errorInstance),
    };
    const repro = `PROPERTY_TEST_SEED=${verdict.seed} PROPERTY_TEST_PATH=${verdict.counterexamplePath} bunx playwright test --project=parity-fuzzer tests/parity/cross-host-flow.spec.ts --workers=1`;
    const marker = failureMarker({
      test: "Four-host saved semantics",
      seed: verdict.seed,
      path: verdict.counterexamplePath,
      repro,
      failure: verdict.errorInstance,
    });
    logFailureMarker(marker);
    writeFailureRecord(
      "fuzz-artifacts/cross-host/findings",
      failureRecord(marker, verdict.errorInstance, {
        flow: verdict.counterexample?.at(0),
      }),
    );
    await testInfo.attach("cross-host-repro", {
      body: JSON.stringify(failure),
      contentType: "application/json",
    });
    throw new Error(
      `Cross-host differential failed: ${JSON.stringify(failure)}\n` +
        `PROPERTY_TEST_SEED=${verdict.seed} PROPERTY_TEST_PATH=${verdict.counterexamplePath} bunx playwright test --project=parity-fuzzer tests/parity/cross-host-flow.spec.ts --workers=1`,
    );
  } finally {
    await react.close();
    await vue.close();
  }
});
