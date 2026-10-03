import { expect, test, type Page } from "@playwright/test";
import { createCanonicalSession } from "../../packages/core/src/controller/canonicalSession";
import { shapeArrayBuffer } from "../../packages/core/src/__tests__/documentShapes";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { validateDocxPackage } from "../../packages/docx-core/src/validate/docx";
import type { BrowserInputAction } from "./browserInputTrace";
import { driveCanonicalBrowserInput } from "./canonicalBrowserInputDriver";
import { canonicalBrowserAcceptances } from "./canonical-browser-acceptance-traces";
import { createMissingOpBurndown } from "../../test/canonical-missing-ops";
import type {} from "../parity/canonicalBridge";
import type {} from "../parity/canonicalFuzzErrors";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const snapshot = async (page: Page) => {
  await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
  const value = await page.evaluate(() => globalThis.__folioCanonical?.snapshot());
  expect(value?.active).toBe(true);
  expect(value?.projectionMatchesCanonical).toBe(true);
  expect(value?.projectionJSON).toEqual(value?.canonicalProjectionJSON);
  if (!value?.document) throw new TypeError("Canonical document unavailable");
  return { ...value, document: value.document };
};
const isHistoryAction = ({ kind }: BrowserInputAction) =>
  kind === "undo" || kind === "redo" || kind === "historyBurst";

const drainErrors = (page: Page) =>
  page.evaluate(() => globalThis.__folioCanonicalFuzzErrors?.splice(0) ?? []);

for (const { seed, trace } of canonicalBrowserAcceptances) {
  for (const mode of ["editing", "suggesting"] as const) {
    test(`canonical acceptance seed ${seed} / ${mode}`, async ({ page }, info) => {
      const missing = createMissingOpBurndown();
      await page.goto("/?session=canonical", { waitUntil: "networkidle" });
      await page.waitForSelector(".layout-page");
      await page.evaluate(() => {
        globalThis.__folioCanonicalFuzzErrors = [];
      });
      const source = await shapeArrayBuffer(trace.shape);
      const eligibility = createCanonicalSession(
        await parseDocx(source, { preloadFonts: false, detectVariables: false }),
      );
      // Temporary source limits include tables (#1474) and images; these traces
      // exercise the canonical path automatically as source support expands.
      if (eligibility.isErr()) {
        expect(eligibility.error.name).toBe("CanonicalSessionError");
        expect(eligibility.error.message).toBe(
          "Canonical sessions currently require plain paragraphs and note references without revisions.",
        );
        missing.record(`source:${trace.shape}`);
        await info.attach("canonical-missing-ops", {
          body: JSON.stringify({ seed, mode, missing: missing.rows() }),
          contentType: "application/json",
        });
        return;
      }
      await page.waitForFunction(() => globalThis.__folioCanonical != null);
      expect(
        await page.evaluate(
          async (bytes) => globalThis.__folioCanonical?.load(bytes),
          [...new Uint8Array(source)],
        ),
      ).toBe(true);
      await page.waitForLoadState("networkidle");
      await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
      expect(
        await page.evaluate((value) => globalThis.__folioCanonical?.setMode(value), mode),
      ).toBe(true);
      expect(await page.evaluate(() => globalThis.__folioCanonical?.select(1))).toBe(true);
      await drainErrors(page);
      for (const action of trace.actions) {
        if (action.kind === "selectionDrag" || action.kind === "dragCellDelete") {
          const structuralTarget =
            action.kind === "selectionDrag" ? action.target : ("table" as const);
          const selected = await page.evaluate(
            (target) => globalThis.__folioCanonical?.selectStructuralTarget(target),
            structuralTarget,
          );
          expect(selected).not.toBeNull();
          expect((await snapshot(page)).selectionJSON).toEqual(selected);
          if (action.kind === "selectionDrag") continue;
        }
        const before = await snapshot(page);
        if (action.kind === "dragCellDelete") await page.keyboard.press("Delete");
        else await driveCanonicalBrowserInput(page, action);
        const after = await snapshot(page);
        const errors = await drainErrors(page);
        if (errors.length > 0) {
          for (const error of errors) {
            expect(error.type).toBe("CanonicalSessionRefusalError");
            missing.record(action.kind);
          }
          expect(after.document).toEqual(before.document);
          expect(after.projectionJSON).toEqual(before.projectionJSON);
          expect(after.canUndo).toBe(before.canUndo);
          expect(after.canRedo).toBe(before.canRedo);
          continue;
        }
        if (action.kind === "imeReplacement" && action.completion === "cancel") {
          expect(after.document).toEqual(before.document);
          expect(after.projectionJSON).toEqual(before.projectionJSON);
          expect(after.selectionJSON).toEqual(before.selectionJSON);
        }
        if (
          !isHistoryAction(action) &&
          JSON.stringify(after.document) !== JSON.stringify(before.document)
        ) {
          await page.keyboard.press(`${MODIFIER}+z`);
          const undone = await snapshot(page);
          expect(undone.document).toEqual(before.document);
          expect(undone.projectionJSON).toEqual(before.projectionJSON);
          expect(undone.selectionJSON).toEqual(before.selectionJSON);
          await page.keyboard.press(`${MODIFIER}+Shift+z`);
          const redone = await snapshot(page);
          expect(redone.document).toEqual(after.document);
          expect(redone.projectionJSON).toEqual(after.projectionJSON);
          expect(await drainErrors(page)).toEqual([]);
        }
      }
      const final = await snapshot(page);
      const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
      if (!saved) throw new TypeError("Canonical save unavailable");
      expect(await validateDocxPackage(new Uint8Array(saved))).toEqual({ valid: true });
      const reopened = structuredClone(
        await parseDocx(new Uint8Array(saved), {
          preloadFonts: false,
          detectVariables: false,
        }),
      );
      expect(reopened.package.document.content).toEqual(final.document.package.document.content);
      await info.attach("canonical-missing-ops", {
        body: JSON.stringify({ seed, mode, missing: missing.rows() }),
        contentType: "application/json",
      });
    });
  }
}
