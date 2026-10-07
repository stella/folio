import { expect, type Page } from "@playwright/test";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import { BROWSER_INPUT_ACTION_DISPOSITIONS, type BrowserInputAction } from "./browserInputTrace";
import { driveCanonicalBrowserInput } from "./canonicalBrowserInputDriver";
import type {} from "../parity/canonicalBridge";
import type {} from "../parity/canonicalFuzzErrors";

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

/** One history oracle for nightly properties and deterministic interaction replays. */
export const checkCanonicalBrowserHistory = async ({
  page,
  source,
  actions,
  missing,
}: CanonicalBrowserHistoryOptions): Promise<number> => {
  // The first case starts with a lazy view. Create it without loading or
  // resetting an existing owner, so this check still detects cross-case leaks.
  expect(await page.evaluate(() => globalThis.__folioCanonical?.ensureView())).toBe(true);
  await page.waitForFunction(() => {
    const composing = globalThis.__folioCanonical?.snapshot().composing;
    return composing !== undefined && composing !== null;
  });
  expect(
    await page.evaluate(() => globalThis.__folioCanonical?.snapshot().composing),
    "case must start outside native composition",
  ).toBe(false);
  expect(await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), source)).toBe(
    true,
  );
  await drainErrors(page);
  expect(await page.evaluate(() => globalThis.__folioCanonical?.select(1, 6))).toBe(true);
  const baseline = await snapshot(page);
  expect(baseline.canUndo).toBe(false);
  expect(baseline.canRedo).toBe(false);
  expect(baseline.selection).toEqual({ from: 1, to: 6 });
  let applied = 0;
  for (const action of actions) {
    const before = await snapshot(page);
    await driveCanonicalBrowserInput(page, action);
    const after = await snapshot(page);
    const errors = await drainErrors(page);
    for (const error of errors) {
      expect(error.type).toBe("CanonicalSessionRefusalError");
      missing.record(action.kind);
    }
    if (errors.length > 0) {
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
      BROWSER_INPUT_ACTION_DISPOSITIONS[action.kind] === "edit" &&
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
    const reopened = structuredClone(
      await parseDocx(new Uint8Array(saved), { preloadFonts: false, detectVariables: false }),
    );
    expect(reopened.package.document.content).toEqual(after.document.package.document.content);
  }
  const final = await snapshot(page);
  const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
  if (!saved) throw new TypeError("Canonical save unavailable");
  expect(await page.evaluate((bytes) => globalThis.__folioCanonical?.load(bytes), saved)).toBe(
    true,
  );
  const reloaded = await snapshot(page);
  expect(reloaded.document.package.document.content).toEqual(
    final.document.package.document.content,
  );
  expect(reloaded.canUndo).toBe(false);
  return applied;
};
