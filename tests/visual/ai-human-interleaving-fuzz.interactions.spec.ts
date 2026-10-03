import { expect, test, type Page } from "@playwright/test";
import fc from "fast-check";
import { createHash } from "node:crypto";
import {
  failureMarker,
  failureRecord,
  logFailureMarker,
  writeFailureRecord,
} from "../../test/consumer-scenarios/support/failure-fingerprints";

import { shapeArrayBuffer } from "../../packages/core/src/__tests__/documentShapes";
import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import { parseBrowserInputTraceConfig } from "./browserInputTrace";
import { interleavingTraceArbitrary, type InterleavingAction } from "./interleavingTrace";
import type {} from "./interleavingBridge";

const replayPath = process.env["PROPERTY_TEST_PATH"];
const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const bridgeUrl = `/@fs${new URL("./interleavingBridge.ts", import.meta.url).pathname}`;
const config = parseBrowserInputTraceConfig(
  process.env,
  process.env["FOLIO_FUZZ_LANE"] === "nightly" ? "nightly" : "pullRequest",
);

const project = (reviewer: FolioDocxReviewer) =>
  reviewer.snapshot().blocks.map(({ kind, text, table }) => ({ kind, text, table }));
const live = (page: Page) =>
  page.evaluate(() => {
    const ref = globalThis.__folioPlayground?.getEditorRef();
    const snapshot = ref?.createAIEditSnapshot();
    if (!ref || !snapshot) throw new Error("interleaving reader unavailable");
    return {
      blocks: snapshot.blocks.map(({ kind, text, table }) => ({ kind, text, table })),
      changes: ref.getTrackedChanges().map(({ id, type, text }) => ({ id, type, text })),
    };
  });
const save = async (page: Page) => {
  const bytes = await page.evaluate(async () => {
    const saved = await globalThis.__folioPlayground?.getEditorRef()?.save({ selective: false });
    if (!saved) throw new Error("interleaving save unavailable");
    return [...new Uint8Array(saved)];
  });
  return new Uint8Array(bytes).buffer;
};

/** Saving must preserve both the reader projection and every pending revision. */
const checkpoint = async (page: Page) => {
  const current = await live(page);
  const bytes = await save(page);
  const reopened = await FolioDocxReviewer.fromBuffer(bytes);
  expect(project(reopened)).toEqual(current.blocks);
  expect(reopened.getChanges().map(({ id, type, text }) => ({ id, type, text }))).toEqual(
    current.changes,
  );
  return { bytes, current, reopened };
};

const drive = async (page: Page, action: InterleavingAction) => {
  switch (action.kind) {
    case "suggest": {
      const pending = await page.evaluate((text) => {
        const suggest = globalThis.__folioInterleavingSuggest;
        if (!suggest) throw new Error("interleaving document-operation bridge unavailable");
        return suggest(text);
      }, action.text);
      expect(pending).toBeGreaterThan(0);
      return;
    }
    case "typing":
      await page.evaluate(() =>
        globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView()?.focus(),
      );
      await page.keyboard.type(action.text);
      return;
    case "undo":
      await page.keyboard.press(`${MODIFIER}+z`);
      return;
    case "redo":
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      return;
    case "accept":
    case "reject": {
      const before = await checkpoint(page);
      const change = before.current.changes.at(
        action.target % Math.max(1, before.current.changes.length),
      );
      if (!change) return;
      const expected =
        action.kind === "accept"
          ? before.reopened.acceptChange(change.id)
          : before.reopened.rejectChange(change.id);
      expect(expected).toBe(true);
      const resolved = await page.evaluate(
        ({ kind, id }) => {
          const ref = globalThis.__folioPlayground?.getEditorRef();
          if (!ref) throw new Error("interleaving resolution unavailable");
          return kind === "accept" ? ref.acceptAIEditOperation(id) : ref.rejectAIEditOperation(id);
        },
        { kind: action.kind, id: change.id },
      );
      expect(resolved).toBe(true);
      const after = await checkpoint(page);
      const expectedReopened = await FolioDocxReviewer.fromBuffer(await before.reopened.toBuffer());
      expect(after.current.blocks).toEqual(project(expectedReopened));
      expect(after.current.changes).toEqual(
        expectedReopened.getChanges().map(({ id, type, text }) => ({ id, type, text })),
      );
      return;
    }
    default: {
      const unreachable: never = action;
      return unreachable;
    }
  }
};

test.setTimeout(600_000);
for (const seed of config.seeds) {
  test(`seed ${seed}: AI and human interleaving preserves revisions, readers and fresh rendering`, async ({
    page,
  }, testInfo) => {
    page.setDefaultTimeout(10_000);
    page.setDefaultNavigationTimeout(30_000);
    let firstFailureRecorded = false;
    let firstFingerprint: string | null = null;
    const verdict = await fc.check(
      fc.asyncProperty(interleavingTraceArbitrary, async (trace) => {
        let stage = "load";
        try {
          const source = await shapeArrayBuffer(trace.shape);
          await page.goto("/");
          await page.waitForSelector(".layout-page");
          await page.evaluate(() =>
            globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView(),
          );
          await page.waitForFunction(
            () => !!globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView(),
          );
          await page.evaluate(
            async (bytes) => {
              const ref = globalThis.__folioPlayground?.getEditorRef();
              if (!ref) throw new Error("interleaving loader unavailable");
              await ref.loadDocumentBuffer(new Uint8Array(bytes));
            },
            [...new Uint8Array(source)],
          );
          stage = "install-document-operation-bridge";
          // Await module loading and installation so import errors reach fast-check
          // instead of leaving an unbounded wait for a missing global function.
          await page.evaluate(async (url) => {
            const { installInterleavingBridge } = await import(/* @vite-ignore */ url);
            installInterleavingBridge();
          }, bridgeUrl);
          await page.waitForFunction(
            () => typeof globalThis.__folioInterleavingSuggest === "function",
          );
          await checkpoint(page);
          for (const [index, action] of trace.actions.entries()) {
            stage = `action-${index}-${action.kind}`;
            await drive(page, action);
            await checkpoint(page);
          }
          // Compare the incrementally painted document with a freshly loaded saved
          // package: reload also rebuilds the editor state and page layout from scratch.
          stage = "fresh-render";
          await page.waitForTimeout(350);
          const painted = await page.locator(".layout-page-content").allTextContents();
          const final = await checkpoint(page);
          await page.evaluate(
            async (bytes) => {
              const ref = globalThis.__folioPlayground?.getEditorRef();
              if (!ref) throw new Error("interleaving loader unavailable");
              await ref.loadDocumentBuffer(new Uint8Array(bytes));
            },
            [...new Uint8Array(final.bytes)],
          );
          await expect
            .poll(() => page.locator(".layout-page-content").allTextContents())
            .toEqual(painted);
          const fresh = await checkpoint(page);
          expect(fresh.current).toEqual(final.current);
        } catch (error) {
          // Preserve the original failure before shrinking can hit teardown or a
          // different boundary. This is a test-runner boundary, not product flow.
          if (!firstFailureRecorded) {
            firstFailureRecorded = true;
            const message = error instanceof Error ? error.message : String(error);
            const failure = {
              seed,
              stage,
              trace,
              error: message,
              fingerprint: createHash("sha256")
                .update(`${stage}:${message.split("\n").at(0)}`)
                .digest("hex")
                .slice(0, 16),
            };
            firstFingerprint = failure.fingerprint;
            console.log(`INTERLEAVING_FAILURE ${JSON.stringify(failure)}`);
            await testInfo.attach("first-interleaving-failure", {
              body: JSON.stringify(failure, null, 2),
              contentType: "application/json",
            });
          }
          throw error;
        }
      }),
      {
        seed,
        ...(replayPath === undefined ? {} : { path: replayPath }),
        // #1342: seeds 11 and 29 shrank to this overlapping insertion trace.
        examples: [
          [
            {
              shape: "plain-markdown",
              actions: [
                { kind: "suggest", text: "alpha" },
                { kind: "typing", text: "alpha" },
                { kind: "suggest", text: "alpha" },
              ],
            },
          ],
        ],
        numRuns: config.runs,
        endOnFailure: false,
        interruptAfterTimeLimit: 540_000,
        markInterruptAsFailure: true,
      },
    );
    if (verdict.failed) {
      const repro = `FOLIO_FUZZ_SEEDS=${seed} FOLIO_FUZZ_RUNS=1 PROPERTY_TEST_PATH=${verdict.counterexamplePath} bunx playwright test --project=interleaving-fuzzer --workers=1`;
      const marker = {
        ...failureMarker({
          test: "AI and human revision interleaving",
          seed,
          path: verdict.counterexamplePath,
          repro,
          failure: verdict.errorInstance,
        }),
        ...(firstFingerprint === null ? {} : { fingerprint: firstFingerprint }),
      };
      logFailureMarker(marker);
      writeFailureRecord(
        "fuzz-artifacts/interleaving/findings",
        failureRecord(marker, verdict.errorInstance, {
          flow: verdict.counterexample?.at(0),
        }),
      );
      throw new Error(
        `seed=${seed} path=${verdict.counterexamplePath} trace=${JSON.stringify(verdict.counterexample?.at(0))}\n${String(verdict.errorInstance)}`,
        { cause: verdict.errorInstance },
      );
    }
  });
}
