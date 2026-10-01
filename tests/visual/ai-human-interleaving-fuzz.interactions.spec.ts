import { expect, test, type Page } from "@playwright/test";
import fc from "fast-check";

import { shapeArrayBuffer } from "../../packages/core/src/__tests__/documentShapes";
import { FolioDocxReviewer } from "../../packages/core/src/ai-edits/headless";
import { parseBrowserInputTraceConfig } from "./browserInputTrace";
import { interleavingTraceArbitrary, type InterleavingAction } from "./interleavingTrace";
import type {} from "./interleavingBridge";

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
        if (!suggest) throw new Error("interleaving tool bridge unavailable");
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
  }) => {
    const verdict = await fc.check(
      fc.asyncProperty(interleavingTraceArbitrary, async (trace) => {
        const source = await shapeArrayBuffer(trace.shape);
        await page.goto("/");
        await page.waitForSelector(".layout-page");
        await page.evaluate(() => globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView());
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
        await page.addScriptTag({
          type: "module",
          content: `import { installInterleavingBridge } from ${JSON.stringify(bridgeUrl)}; installInterleavingBridge();`,
        });
        await page.waitForFunction(
          () => typeof globalThis.__folioInterleavingSuggest === "function",
        );
        await checkpoint(page);
        for (const action of trace.actions) {
          await drive(page, action);
          await checkpoint(page);
        }
        // Compare the incrementally painted document with a freshly loaded saved
        // package: reload also rebuilds the editor state and page layout from scratch.
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
      }),
      { seed, numRuns: config.runs, endOnFailure: false },
    );
    if (verdict.failed) {
      throw new Error(
        `seed=${seed} path=${verdict.counterexamplePath} trace=${JSON.stringify(verdict.counterexample?.at(0))}\n${String(verdict.errorInstance)}`,
        { cause: verdict.errorInstance },
      );
    }
  });
}
