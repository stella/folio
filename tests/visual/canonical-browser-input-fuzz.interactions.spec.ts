import { expect, test, type Page } from "@playwright/test";
import fc from "fast-check";
import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";
import { appendFileSync } from "node:fs";

import { createDocx } from "../../packages/core/src/docx/rezip";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import {
  failureMarker,
  failureRecord,
  logFailureMarker,
  shellQuote,
  writeFailureRecord,
} from "../../test/consumer-scenarios/support/failure-fingerprints";
import {
  commonActionArbitraries,
  parseBrowserInputTraceConfig,
  type BrowserInputAction,
} from "./browserInputTrace";
import type {} from "../parity/canonicalBridge";
import type {} from "../parity/canonicalFuzzErrors";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const config = parseBrowserInputTraceConfig(
  process.env,
  process.env["FOLIO_FUZZ_LANE"] === "nightly" ? "nightly" : "pullRequest",
);
const traceArbitrary = fc.array(fc.oneof(...commonActionArbitraries), {
  minLength: 1,
  maxLength: 12,
});
const snapshot = async (page: Page) => {
  const current = await page.evaluate(() => globalThis.__folioCanonical?.snapshot());
  expect(current?.active).toBe(true);
  expect(current?.projectionMatchesCanonical).toBe(true);
  expect(current?.projectionJSON).toEqual(current?.canonicalProjectionJSON);
  if (!current?.document) throw new TypeError("Canonical document unavailable");
  return current;
};
const drainErrors = (page: Page) =>
  page.evaluate(() => {
    const errors = globalThis.__folioCanonicalFuzzErrors;
    if (!errors) throw new TypeError("Canonical error sink unavailable");
    return errors.splice(0);
  });
const drive = async (page: Page, action: BrowserInputAction) => {
  switch (action.kind) {
    case "typing":
      await page.keyboard.insertText(action.text);
      return;
    case "enter":
      await page.keyboard.press("Enter");
      return;
    case "backspace":
      await page.keyboard.press("Backspace");
      return;
    case "delete":
      await page.keyboard.press("Delete");
      return;
    case "undo":
      await page.keyboard.press(`${MODIFIER}+z`);
      return;
    case "redo":
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      return;
    case "cut":
      await page.keyboard.press(`${MODIFIER}+x`);
      return;
    case "imeReplacement": {
      const cdp = await page.context().newCDPSession(page);
      await cdp.send("Input.imeSetComposition", {
        text: action.text,
        selectionStart: action.text.length,
        selectionEnd: action.text.length,
      });
      await cdp.send("Input.insertText", { text: action.text });
      await cdp.detach();
      return;
    }
    case "pastePlain":
    case "pasteHtml":
    case "pasteWordHtml":
    case "pasteListHtml":
    case "pasteTable":
    case "pasteMultiBlock":
      await page.context().grantPermissions(["clipboard-read", "clipboard-write"]);
      await page.evaluate(async ({ plain, html }) => {
        const parts: Record<string, Blob> = {
          "text/plain": new Blob([plain], { type: "text/plain" }),
        };
        if (html) parts["text/html"] = new Blob([html], { type: "text/html" });
        await navigator.clipboard.write([new ClipboardItem(parts)]);
      }, action);
      await page.keyboard.press(`${MODIFIER}+v`);
      return;
    case "dragCellDelete":
    case "selectionDrag":
      throw new TypeError("Structural selections require a structural seed");
    default: {
      const unreachable: never = action;
      return unreachable;
    }
  }
};

for (const seed of config.seeds) {
  test(`canonical seed ${seed}: projection, exact history and save/reopen`, async ({
    page,
  }, info) => {
    test.setTimeout(600_000);
    const missing = createMissingOpBurndown();
    let applied = 0;
    let completed = 0;
    await page.goto("/?session=canonical");
    await page.waitForSelector(".layout-page");
    await page.evaluate(() => {
      globalThis.__folioCanonicalFuzzErrors = [];
    });
    const source = await createDocx(createEmptyDocument({ initialText: "alpha😀café東京" }));
    const verdict = await fc.check(
      fc.asyncProperty(traceArbitrary, async (actions) => {
        expect(
          await page.evaluate(
            async (bytes) => globalThis.__folioCanonical?.load(bytes),
            [...new Uint8Array(source)],
          ),
        ).toBe(true);
        await drainErrors(page);
        expect(await page.evaluate(() => globalThis.__folioCanonical?.select(1, 6))).toBe(true);
        await snapshot(page);
        for (const action of actions) {
          const before = await snapshot(page);
          await drive(page, action);
          const after = await snapshot(page);
          const errors = await drainErrors(page);
          for (const error of errors) {
            expect(error.type).toBe("CanonicalSessionRefusalError");
            missing.record(action.kind);
          }
          if (errors.length > 0) {
            // An explicit refusal cannot edit the model, projection or history.
            expect(after.document).toEqual(before.document);
            expect(after.projectionJSON).toEqual(before.projectionJSON);
            expect(after.canUndo).toBe(before.canUndo);
            expect(after.canRedo).toBe(before.canRedo);
            continue;
          }
          if (action.kind === "typing") {
            if (!before.textSelection) throw new TypeError("Missing input selection");
            expect(after.text).toBe(
              before.textSelection.before + action.text + before.textSelection.after,
            );
          }
          if (
            action.kind !== "undo" &&
            action.kind !== "redo" &&
            JSON.stringify(after.document) !== JSON.stringify(before.document)
          ) {
            applied++;
            await page.keyboard.press(`${MODIFIER}+z`);
            const undone = await snapshot(page);
            expect(undone.document).toEqual(before.document);
            expect(undone.projectionJSON).toEqual(before.projectionJSON);
            expect(undone.selection).toEqual(before.selection);
            await page.keyboard.press(`${MODIFIER}+Shift+z`);
            const redone = await snapshot(page);
            expect(redone.document).toEqual(after.document);
            expect(redone.projectionJSON).toEqual(after.projectionJSON);
            expect(redone.selection).toEqual(after.selection);
            expect(await drainErrors(page)).toEqual([]);
          }
          const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
          if (!saved) throw new TypeError("Canonical save unavailable");
          expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({ valid: true });
          const reopened = await parseDocx(new Uint8Array(saved), {
            preloadFonts: false,
            detectVariables: false,
          });
          expect(reopened.package.document.content).toEqual(
            after.document.package.document.content,
          );
        }
        // Reopen through the actual host as well as the DOCX reader.
        const final = await snapshot(page);
        const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
        if (!saved) throw new TypeError("Canonical save unavailable");
        expect(
          await page.evaluate(async (bytes) => globalThis.__folioCanonical?.load(bytes), saved),
        ).toBe(true);
        const reloaded = await snapshot(page);
        expect(reloaded.document.package.document.content).toEqual(
          final.document.package.document.content,
        );
        expect(reloaded.canUndo).toBe(false);
        completed++;
      }),
      {
        seed,
        numRuns: config.runs,
        endOnFailure: false,
        ...(process.env["FOLIO_FUZZ_PATH"] === undefined
          ? {}
          : { path: process.env["FOLIO_FUZZ_PATH"] }),
      },
    );
    await info.attach("canonical-missing-ops", {
      body: JSON.stringify({
        seed,
        applied,
        completed,
        missing: missing.rows(),
        verdict: verdict.failed ? "failure" : "completed",
      }),
      contentType: "application/json",
    });
    const summaryPath = process.env["GITHUB_STEP_SUMMARY"];
    if (summaryPath)
      appendFileSync(
        summaryPath,
        `\nCanonical seed ${seed}: ${completed} completed traces, ${applied} applied edits\n\n${missing.markdown()}\n`,
      );
    console.log(
      `Canonical seed ${seed}: ${completed} completed traces, ${applied} applied edits\n${missing.markdown()}`,
    );
    if (verdict.failed) {
      const flow = verdict.counterexample?.at(0);
      const failure = verdict.errorInstance;
      const detail =
        failure instanceof Error ? (failure.stack ?? failure.message) : fc.stringify(failure);
      const message = failure instanceof Error ? failure.message : detail;
      if (
        flow !== undefined &&
        failure !== undefined &&
        failure !== null &&
        message.trim() !== "" &&
        message !== "undefined"
      ) {
        const repro = `FOLIO_FUZZ_LANE=nightly FOLIO_FUZZ_SEEDS=${verdict.seed} FOLIO_FUZZ_RUNS=${config.runs}${verdict.counterexamplePath === null ? "" : ` FOLIO_FUZZ_PATH=${shellQuote(verdict.counterexamplePath)}`} bunx playwright test --project=browser-fuzzer tests/visual/canonical-browser-input-fuzz.interactions.spec.ts --workers=1`;
        const marker = failureMarker({
          test: "canonical browser input: projection, exact history and save/reopen",
          seed: verdict.seed,
          path: verdict.counterexamplePath,
          repro,
          failure,
          flow: flow.map(({ kind }) => kind).join(" → "),
        });
        logFailureMarker(marker);
        const artifact = writeFailureRecord(
          process.env["FOLIO_FUZZ_FAILURES_DIR"] ?? "fuzz-artifacts/canonical/findings",
          failureRecord(marker, failure, { flow }),
        );
        await info.attach("canonical-failure-record", {
          path: artifact,
          contentType: "application/json",
        });
      }
      throw new Error(
        `seed=${verdict.seed} path=${verdict.counterexamplePath} trace=${JSON.stringify(flow)}\n${detail}`,
        { cause: failure },
      );
    }
    expect(completed).toBeGreaterThan(0);
  });
}
