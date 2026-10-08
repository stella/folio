import { expect, type Page } from "@playwright/test";
import { Result } from "better-result";
import { captureCanonicalOracleFailure } from "../parity/canonicalOracleFailure";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import {
  canonicalBrowserRefusalRows,
  matchCanonicalRefusalRow,
  validateHarnessRefusalRows,
} from "../../test/canonical-refusal-rows";
import { BROWSER_INPUT_ACTION_DISPOSITIONS, type BrowserInputAction } from "./browserInputTrace";
import { driveCanonicalBrowserInput } from "./canonicalBrowserInputDriver";
import { assertCanonicalInputTimersSettled } from "./canonicalTimerProbe";
import type {} from "../parity/canonicalBridge";
import {
  isCanonicalSaveFallback,
  type CanonicalFuzzObservation,
  type CanonicalFuzzPhase,
} from "../parity/canonicalFuzzErrors";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const snapshot = async (page: Page) => {
  // An ended IME composition commits after the native flush settles; the
  // canonical document refuses snapshots until then (canSnapshot is false).
  await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
  const current = await page.evaluate(() => globalThis.__folioCanonical?.snapshot());
  expect(current?.active).toBe(true);
  expect(current?.composing, "native composition must end before canonical snapshots").toBe(false);
  expect(current?.projectionMatchesCanonical).toBe(true);
  expect(current?.projectionJSON).toEqual(current?.canonicalProjectionJSON);
  if (!current?.document) throw new TypeError("Canonical document unavailable");
  return { ...current, document: current.document };
};
const drainErrors = (page: Page) =>
  page.evaluate(() => {
    const errors = globalThis.__folioCanonicalFuzzErrors;
    if (!errors) throw new TypeError("Canonical error sink unavailable");
    return errors.splice(0);
  });

type CanonicalBrowserHistoryOptions = {
  page: Page;
  source: number[];
  actions: readonly BrowserInputAction[];
  missing: ReturnType<typeof createMissingOpBurndown>;
};

type CanonicalBrowserHistoryRunOptions = CanonicalBrowserHistoryOptions & {
  observations: CanonicalFuzzObservation[];
};

/** Establish the first loaded owner once, outside the per-case reset barrier. */
export const initializeCanonicalBrowserHistory = async (page: Page, source: number[]) => {
  await page.evaluate(() => {
    globalThis.__folioCanonicalFuzzErrors ??= [];
  });
  expect(await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), source)).toBe(
    true,
  );
  await expect
    .poll(
      async () =>
        typeof (await page.evaluate(() => globalThis.__folioCanonical?.nativeComposing())),
      { message: "canonical fixture must create its initial view" },
    )
    .toBe("boolean");
};

const runCanonicalBrowserHistory = async ({
  page,
  source,
  actions,
  missing,
  observations,
}: CanonicalBrowserHistoryRunOptions): Promise<number> => {
  let observation: CanonicalFuzzObservation = { phase: { type: "load" }, errors: [] };
  const beginPhase = async (phase: CanonicalFuzzPhase) => {
    observation = { phase, errors: [] };
    observations.push(observation);
    await page.evaluate((current) => {
      globalThis.__folioCanonicalFuzzPhase = current;
    }, phase);
  };
  const collectErrors = async () => {
    observation.errors = await drainErrors(page);
    return observation.errors;
  };
  await assertCanonicalInputTimersSettled(page);
  await beginPhase({ type: "load" });
  // Bootstrap happens outside the oracle. Ensure without reloading here so
  // the strict pre-load barrier still detects cross-case lifecycle failures.
  expect(await page.evaluate(() => globalThis.__folioCanonical?.ensureView())).toBe(true);
  await expect
    .poll(
      async () =>
        typeof (await page.evaluate(() => globalThis.__folioCanonical?.nativeComposing())),
      { message: "canonical view must exist after initialization" },
    )
    .toBe("boolean");
  expect(
    await page.evaluate(() => globalThis.__folioCanonical?.nativeComposing()),
    "case must start outside native composition",
  ).toBe(false);
  expect(await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), source)).toBe(
    true,
  );
  expect(await collectErrors()).toEqual([]);
  expect(await page.evaluate(() => globalThis.__folioCanonical?.select(1, 6))).toBe(true);
  const baseline = await snapshot(page);
  expect(baseline.canUndo).toBe(false);
  expect(baseline.canRedo).toBe(false);
  expect(baseline.selection).toEqual({ from: 1, to: 6 });
  let applied = 0;
  for (const [index, action] of actions.entries()) {
    const before = await snapshot(page);
    await beginPhase({ type: "input", index, action: action.kind });
    await driveCanonicalBrowserInput(page, action);
    const after = await snapshot(page);
    const errors = await collectErrors();
    const rows = canonicalBrowserRefusalRows(action);
    validateHarnessRefusalRows(rows);
    if (rows.length > 0) expect(errors.length).toBeGreaterThan(0);
    for (const error of errors) {
      expect(error.status).toBe("refusal");
      expect(error.type).toBe("CanonicalSessionRefusalError");
      if (error.status !== "refusal") throw new TypeError(error.message);
      const matchedRow = matchCanonicalRefusalRow({ rows, refusal: error });
      expect(matchedRow).toBeDefined();
      if (!matchedRow)
        throw new TypeError("Canonical browser refusal did not match a declared row.");
      missing.record(matchedRow.gap);
    }
    if (errors.length > 0) {
      expect(after.document).toEqual(before.document);
      expect(after.projectionJSON).toEqual(before.projectionJSON);
      expect(after.selection).toEqual(before.selection);
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
      BROWSER_INPUT_ACTION_DISPOSITIONS[action.kind] === "edit" &&
      JSON.stringify(after.document) !== JSON.stringify(before.document)
    ) {
      applied++;
      await beginPhase({ type: "undo", index, action: action.kind });
      await page.keyboard.press(`${MODIFIER}+z`);
      const undone = await snapshot(page);
      expect(await collectErrors()).toEqual([]);
      expect(undone.document).toEqual(before.document);
      expect(undone.projectionJSON).toEqual(before.projectionJSON);
      expect(undone.selection).toEqual(before.selection);
      await beginPhase({ type: "redo", index, action: action.kind });
      await page.keyboard.press(`${MODIFIER}+Shift+z`);
      const redone = await snapshot(page);
      expect(await collectErrors()).toEqual([]);
      expect(redone.document).toEqual(after.document);
      expect(redone.projectionJSON).toEqual(after.projectionJSON);
      expect(redone.selection).toEqual(after.selection);
    }
    await beginPhase({ type: "save", index, action: action.kind });
    const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
    expect((await collectErrors()).filter((error) => !isCanonicalSaveFallback(error))).toEqual([]);
    if (!saved) throw new TypeError("Canonical save unavailable");
    expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({ valid: true });
    const reopened = structuredClone(
      await parseDocx(new Uint8Array(saved), { preloadFonts: false, detectVariables: false }),
    );
    expect(reopened.package.document.content).toEqual(after.document.package.document.content);
    expect(reopened.package.numbering).toEqual(after.document.package.numbering);
  }
  const final = await snapshot(page);
  await assertCanonicalInputTimersSettled(page);
  await beginPhase({ type: "finalSave" });
  const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
  expect((await collectErrors()).filter((error) => !isCanonicalSaveFallback(error))).toEqual([]);
  if (!saved) throw new TypeError("Canonical save unavailable");
  expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({ valid: true });
  await beginPhase({ type: "reload" });
  expect(await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), saved)).toBe(
    true,
  );
  const reloaded = await snapshot(page);
  expect(await collectErrors()).toEqual([]);
  expect(reloaded.document.package.document.content).toEqual(
    final.document.package.document.content,
  );
  expect(reloaded.document.package.numbering).toEqual(final.document.package.numbering);
  expect(reloaded.canUndo).toBe(false);
  await assertCanonicalInputTimersSettled(page);
  return applied;
};

/** One history oracle for nightly properties and deterministic interaction replays. */
export const checkCanonicalBrowserHistory = async (options: CanonicalBrowserHistoryOptions) => {
  const observations: CanonicalFuzzObservation[] = [];
  const result = await Result.tryPromise({
    try: () => runCanonicalBrowserHistory({ ...options, observations }),
    catch: (cause: unknown) => cause,
  });
  if (result.isOk()) return result.value;
  throw await captureCanonicalOracleFailure({
    cause: result.error,
    observations,
    drainErrors: () => drainErrors(options.page),
  });
};
