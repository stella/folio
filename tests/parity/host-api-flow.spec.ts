import { test, expect } from "@playwright/test";
import type { Page } from "@playwright/test";
import fc from "fast-check";
import JSZip from "jszip";
import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";
import { readFile, writeFile } from "node:fs/promises";
import path from "node:path";
import {
  FolioDocxReviewer,
  type FolioDocumentOperationBatch,
} from "../../packages/core/src/server";
import { getDocumentText } from "../../packages/core/src/docx/documentParser";
import {
  failureMarker,
  logFailureMarker,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import { runFolioCli } from "../../packages/cli/src/cli";
import { fileVersionOf } from "../../packages/cli/src/document";
import { makeTempDir } from "../../packages/cli/src/__tests__/fixtures";
import { propertyConfig, propertyTestTimeout } from "../../test/property-testing";
import { buildScrollRootDocument } from "../support/scrollRootDocument";
import { ensureLiveView, openEditor } from "./parity-fixture";
import { HOST_NAVIGATION_CASES, hostApiFlowArbitrary, navigationWasEffective } from "./hostApiFlow";

const semanticProjection = async (bytes: Uint8Array) => {
  expect(await validateDocxPackage(bytes)).toEqual({ valid: true });
  const reviewer = await FolioDocxReviewer.fromBuffer(new Uint8Array(bytes).buffer);
  return {
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

const applyCliBatch = async (file: string, batch: FolioDocumentOperationBatch) => {
  const stdout: string[] = [];
  const stderr: string[] = [];
  const before = new Uint8Array(await readFile(file));
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
      cwd: path.dirname(file),
    },
  );
  expect(exit, stderr.join("")).toBe(0);
  expect(JSON.parse(stdout.join(""))).toMatchObject({
    ok: true,
    data: {
      status: "committed",
      result: {
        skipped: [],
        applied: [{ id: batch.operations.at(0)?.id }],
      },
    },
  });
  expect(fileVersionOf(new Uint8Array(await readFile(file)))).not.toBe(fileVersionOf(before));
};

const waitForLayout = async (page: Page) => {
  await ensureLiveView(page);
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getTotalPages() ?? 0))
    .toBeGreaterThan(3);
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
};

type NavigationCase = (typeof HOST_NAVIGATION_CASES)[number];
const navigate = async (page: Page, navigation: NavigationCase) => {
  const suggestion = navigation.type === "document" && navigation.method === "scrollToSuggestion";
  if (suggestion) {
    expect(await page.evaluate(() => window.__folioScrollParity?.prepareSuggestion())).toBe(true);
    await waitForLayout(page);
  }
  await page.evaluate(() => window.__folioScrollParity?.resetFlowScroll());
  const method = navigation.type === "document" ? navigation.method : undefined;
  const before = await page.evaluate((api) => window.__folioScrollParity?.readTarget(api), method);
  expect(before).toBeTruthy();
  if (!before) throw new Error("Navigation target unavailable before call");
  expect(before.top).toBeGreaterThan(before.viewportBottom);
  const outerBefore = await page.evaluate(() => window.scrollY);
  expect(
    await page.evaluate((action) => {
      switch (action.type) {
        case "document":
          return window.__folioScrollParity?.navigate(action.method);
        case "paged":
          return window.__folioScrollParity?.navigatePaged(action.method);
        default: {
          const exhaustive: never = action;
          return exhaustive;
        }
      }
    }, navigation),
  ).toBe(true);
  await expect(async () => {
    const after = await page.evaluate((api) => window.__folioScrollParity?.readTarget(api), method);
    if (!after) throw new Error("Navigation target unavailable after call");
    expect(
      navigationWasEffective({
        before,
        after,
        outerBefore,
        outerAfter: await page.evaluate(() => window.scrollY),
      }),
    ).toBe(true);
  }).toPass({ timeout: 5_000 });
  if (!suggestion) return;
  // Suggestions are ephemeral navigation setup. Reject them before comparing
  // saved content with the reference hosts, which received only direct edits.
  expect(await page.evaluate(() => window.__folioScrollParity?.rejectFlowSuggestion())).toBe(true);
  await waitForLayout(page);
};

test("host APIs interleaved with edits save equally across React, Vue, headless and CLI", async ({
  browser,
}, testInfo) => {
  test.setTimeout(propertyTestTimeout(240_000));
  const react = await browser.newPage();
  const vue = await browser.newPage();
  const hosts = [
    { name: "react", page: react },
    { name: "vue", page: vue },
  ] as const;
  const source = await buildScrollRootDocument();
  const zip = await JSZip.loadAsync(source);
  const part = zip.file("word/document.xml");
  if (!part) throw new Error("Missing fixture document part");
  zip.file(
    "word/document.xml",
    (await part.async("string")).replace("First page", "Replacement page"),
  );
  const replacement = await zip.generateAsync({ type: "uint8array" });
  const { dir, cleanup } = await makeTempDir();
  const file = path.join(dir, "host-flow.docx");
  try {
    for (const { name, page } of hosts) {
      await page.setViewportSize({ width: 1280, height: 600 });
      await page.emulateMedia({ reducedMotion: "reduce" });
      const port =
        name === "react"
          ? Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200
          : Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;
      await page.route("**/fixtures/host-flow.docx", (route) =>
        route.fulfill({
          body: Buffer.from(source),
          contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }),
      );
      await openEditor(page, { name, baseUrl: `http://localhost:${port}` }, "host-flow.docx");
      await page.evaluate(() => {
        document.documentElement.style.height = "2200px";
        window.scrollTo({ top: 120, behavior: "instant" });
      });
    }
    const verdict = await fc.check(
      fc.asyncProperty(hostApiFlowArbitrary, async (flow) => {
        let headless = await FolioDocxReviewer.fromBuffer(new Uint8Array(source).buffer);
        let expectedFirstText = "First page";
        const exercised = new Map(hosts.map(({ name }) => [name, new Set<string>()]));
        await writeFile(file, source);
        for (const { page } of hosts) {
          await page.evaluate(
            (bytes) => window.__folioScrollParity?.loadFlowDocument(bytes),
            [...source],
          );
          await waitForLayout(page);
        }
        for (const [index, navigation] of flow.navigation.entries()) {
          if (index === flow.replacementAfter) {
            headless = await FolioDocxReviewer.fromBuffer(new Uint8Array(replacement).buffer);
            expectedFirstText = "Replacement page";
            await writeFile(file, replacement);
            for (const { page } of hosts) {
              await page.evaluate(
                (bytes) => window.__folioScrollParity?.replaceFlowDocument(bytes),
                [...replacement],
              );
              await waitForLayout(page);
            }
          }
          const text = flow.edits.at(index);
          if (text === undefined) throw new Error("Generated flow missing edit");
          const id = `host-edit-${index}`;
          const batch = {
            version: 1,
            mode: "direct",
            operations: [
              {
                id,
                type: "replaceInBlock",
                blockId: "13300100",
                find: "page",
                replace: `page ${text}`,
              },
            ],
          } as const satisfies FolioDocumentOperationBatch;
          expectedFirstText = expectedFirstText.replace("page", `page ${text}`);
          expect(headless.applyDocumentOperations(batch)).toMatchObject({
            skipped: [],
            applied: [{ id }],
          });
          await applyCliBatch(file, batch);
          const baselineBytes = new Uint8Array(await headless.toBuffer());
          const baseline = await semanticProjection(baselineBytes);
          expect(baseline.blocks.at(0)?.text).toBe(expectedFirstText);
          for (const { name, page } of hosts) {
            expect(
              await page.evaluate(
                (input) => window.__folioScrollParity?.applyFlowBatch(input),
                batch,
              ),
            ).toMatchObject({ skipped: [], applied: [{ id }] });
            await waitForLayout(page);
            await navigate(page, navigation);
            exercised.get(name)?.add(`${navigation.type}:${navigation.method}`);
            const read = await page.evaluate(() => window.__folioScrollParity?.readFlowDocument());
            expect(
              read ? getDocumentText(read.documentBody).split("\n").at(0) : undefined,
              `${name}:getDocument`,
            ).toBe(expectedFirstText);
            expect(
              read ? getDocumentText(read.pagedBody).split("\n").at(0) : undefined,
              `${name}:paged.getDocument`,
            ).toBe(expectedFirstText);
            expect(read?.liveText, `${name}:live`).toContain(expectedFirstText);
            const saved = await page.evaluate(() => window.__folioScrollParity?.saveFlowDocument());
            if (!saved) throw new Error("Missing browser checkpoint output");
            expect(await semanticProjection(new Uint8Array(saved)), `${name}:${index}`).toEqual(
              baseline,
            );
            // Fresh headless parsing above is the save/reopen oracle at every step.
          }
          expect(
            await semanticProjection(new Uint8Array(await readFile(file))),
            `cli:${index}`,
          ).toEqual(baseline);
        }
        const declared = HOST_NAVIGATION_CASES.map(
          ({ type, method }) => `${type}:${method}`,
        ).sort();
        for (const { name } of hosts) {
          expect([...(exercised.get(name) ?? [])].sort(), `${name}:navigation coverage`).toEqual(
            declared,
          );
        }
      }),
      propertyConfig({ numRuns: 2, endOnFailure: false }),
    );
    if (!verdict.failed) return;
    const failure = {
      seed: verdict.seed,
      path: verdict.counterexamplePath,
      trace: verdict.counterexample?.at(0),
      error: String(verdict.errorInstance),
    };
    await testInfo.attach("host-api-repro", {
      body: JSON.stringify(failure),
      contentType: "application/json",
    });
    logFailureMarker(
      failureMarker({
        test: "host APIs interleaved with edits save equally across React, Vue, headless and CLI",
        seed: verdict.seed,
        path: verdict.counterexamplePath,
        repro: `PROPERTY_TEST_SEED=${verdict.seed} PROPERTY_TEST_PATH=${verdict.counterexamplePath} bunx playwright test --project=parity-fuzzer tests/parity/host-api-flow.spec.ts --workers=1`,
        failure: verdict.errorInstance,
        flow: fc.stringify(failure.trace),
      }),
    );
    throw new Error(
      `Host API flow failed: ${JSON.stringify(failure)}\n` +
        `PROPERTY_TEST_SEED=${verdict.seed} PROPERTY_TEST_PATH=${verdict.counterexamplePath} bunx playwright test --project=parity-fuzzer tests/parity/host-api-flow.spec.ts --workers=1`,
    );
  } finally {
    await react.close();
    await vue.close();
    await cleanup();
  }
});
