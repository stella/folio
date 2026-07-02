/**
 * Runtime smoke over the PACKED @stll/folio-* tarballs.
 *
 * The build gate (`packaged-consumer-build.ts`) proves a production `vite build`
 * RESOLVES everything the tarballs reference. This spec proves the built output
 * RUNS: it drives the served consumer app (`app.tsx`) in a real browser and
 * fails on any console error, catching the two runtime-only failure classes the
 * build gate is blind to:
 *
 *   1. A font-metrics worker that resolves at build time but fails to
 *      spawn/execute at runtime. The worker is a flag-gated, best-effort cache
 *      pre-warm — NOT on the critical layout path — so a rendered page proves
 *      nothing about it. The app enables the flag and exposes a real worker
 *      round-trip; a width only comes back if the shipped worker actually
 *      spawned, ran, and replied.
 *   2. UI strings missing from the shipped runtime catalog. use-intl reports a
 *      missing `t("...")` key as `IntlError: MISSING_MESSAGE` via its default
 *      `console.error` handler, which the console net below turns into a failure.
 *
 * Every console message and pageerror for the whole session is collected;
 * error-level output fails the run, and the `/MISSING_MESSAGE|IntlError/` and
 * `/worker/i` patterns must never appear at ANY level.
 */

import { expect, test } from "@playwright/test";
import type { ConsoleMessage } from "@playwright/test";

import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";

declare global {
  var __folioSmoke:
    | {
        getEditorRef: () => DocxEditorRef | null;
        measureRoundTrip: () => Promise<{ width: number; alive: boolean }>;
      }
    | undefined;
}

type CapturedMessage = { level: string; text: string };

const MISSING_MESSAGE_PATTERN = /MISSING_MESSAGE|IntlError/u;
const WORKER_PATTERN = /worker/iu;

test("packaged consumer mounts, lays out via the worker, and logs no errors", async ({ page }) => {
  const messages: CapturedMessage[] = [];
  const pageErrors: string[] = [];

  page.on("console", (message: ConsoleMessage) => {
    messages.push({ level: message.type(), text: message.text() });
  });
  page.on("pageerror", (error) => {
    pageErrors.push(error.stack ?? error.message);
  });

  await page.goto("/");

  // --- Surface 1: the editor mounts and lays out. ---------------------------
  await page.waitForSelector('[data-testid="folio-editor"]', { timeout: 20_000 });
  await page.waitForFunction(
    () => document.querySelectorAll(".layout-page").length >= 1,
    undefined,
    {
      timeout: 20_000,
    },
  );
  await page.evaluate(() => document.fonts.ready);

  // A painted page with real geometry: paged layout ran to completion.
  const pageBox = await page.locator(".layout-page").first().boundingBox();
  expect(pageBox, "a page must be painted").not.toBeNull();
  expect(pageBox?.width ?? 0).toBeGreaterThan(0);
  expect(pageBox?.height ?? 0).toBeGreaterThan(0);

  // --- Worker: a genuine round-trip through the shipped worker. -------------
  // Layout has already exercised the real worker-construction path via
  // `prefetchMeasurement` (the flag is on). This drives one more measurement and
  // asserts the worker returned a width: proof it spawned, ran, and replied. On
  // a broken worker this width never lands and `alive` flips false.
  const worker = await page.evaluate(() => globalThis.__folioSmoke?.measureRoundTrip());
  expect(worker, "smoke hook must be installed").toBeDefined();
  expect(worker?.alive, "worker proxy must not have errored out").toBe(true);
  expect(worker?.width ?? -1, "worker must return a measured width > 0").toBeGreaterThan(0);

  // --- Surface 2: the toolbar renders (chrome + catalog strings). -----------
  await expect(page.getByRole("toolbar").first()).toBeVisible();

  // --- Surface 3: a popover surface opens (context menu on a paragraph). -----
  // Force the deferred ProseMirror view to materialise so the document body is
  // interactive, then right-click to open the text editing menu.
  await page.evaluate(() => {
    globalThis.__folioSmoke?.getEditorRef()?.ensureEditorView({ focus: true });
  });
  await page.waitForFunction(
    () => !!globalThis.__folioSmoke?.getEditorRef()?.getEditorRef()?.getView(),
    undefined,
    { timeout: 10_000 },
  );

  const paragraph = page.locator(".layout-paragraph").first();
  await paragraph.scrollIntoViewIfNeeded();
  const paraBox = await paragraph.boundingBox();
  expect(paraBox, "a paragraph must be painted").not.toBeNull();
  if (paraBox) {
    const x = paraBox.x + paraBox.width / 2;
    const y = paraBox.y + paraBox.height / 2;
    await page.mouse.click(x, y);
    await page.mouse.click(x, y, { button: "right" });
  }
  const menu = page.locator('[role="menu"][aria-label="Text editing menu"]');
  await expect(menu).toBeVisible({ timeout: 5_000 });
  await page.keyboard.press("Escape");

  // Let any deferred error surface (async worker error events, late intl reads).
  await page.waitForTimeout(500);

  // --- Console net: the whole-session assertions. ---------------------------
  const errorLevel = messages.filter((m) => m.level === "error");
  expect(errorLevel, `console error(s): ${JSON.stringify(errorLevel)}`).toEqual([]);
  expect(pageErrors, `pageerror(s): ${JSON.stringify(pageErrors)}`).toEqual([]);

  const intlHits = messages.filter((m) => MISSING_MESSAGE_PATTERN.test(m.text));
  expect(intlHits, `missing-message log(s): ${JSON.stringify(intlHits)}`).toEqual([]);

  const workerHits = messages.filter((m) => WORKER_PATTERN.test(m.text));
  expect(workerHits, `worker log(s): ${JSON.stringify(workerHits)}`).toEqual([]);
});
