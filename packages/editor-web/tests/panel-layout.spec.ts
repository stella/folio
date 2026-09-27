/**
 * The side-panel layout oracle from the playground spec
 * (`tests/visual/panel-layout.interactions.spec.ts`), run against the built
 * VS Code bundle: a narrow and a medium pane, with and without review
 * content, under the webview's policy and theme mapping.
 */

import { expect, test } from "@playwright/test";

import {
  PANEL_LAYOUT_CASES,
  afterPanelDrawersDismissed,
  expectPanelDrawers,
  expectPanelsDoNotOverlap,
  readPanelState,
  waitForPanels,
} from "../../../tests/support/panelLayoutAssertions";
import { buildPanelLayoutDocument } from "../../../tests/support/panelLayoutDocument";
import { openWebview, postDocument, waitForSent } from "./webviewHarness";

const WEBVIEW_WIDTHS = new Set([480, 1100]);
const VIEWPORT_HEIGHT = 900;

for (const { width, review, expected } of PANEL_LAYOUT_CASES) {
  if (!WEBVIEW_WIDTHS.has(width)) continue;
  test(`webview ${String(width)}px, review ${review}: ${expected.tier} tier`, async ({ page }) => {
    const pageErrors: string[] = [];
    page.on("pageerror", (error) => pageErrors.push(error.message));
    await page.setViewportSize({ width, height: VIEWPORT_HEIGHT });
    await openWebview(page, "dark");
    await waitForSent(page, { type: "ready" });
    await postDocument(page, {
      type: "load",
      document: {
        bytes: await buildPanelLayoutDocument(review),
        fileVersion: "v1",
        fileName: "nda.docx",
      },
      author: "Test Author",
      mode: "editing",
      locale: "en",
    });
    await waitForSent(page, { type: "loaded", fileVersion: "v1" });
    await waitForPanels(page, review);

    expect(await readPanelState(page)).toEqual(expected);
    await expectPanelsDoNotOverlap(page, expected);
    await expectPanelDrawers(page, expected);
    const afterDismissal = afterPanelDrawersDismissed(expected);
    expect(await readPanelState(page)).toEqual(afterDismissal);
    expect(pageErrors).toEqual([]);
  });
}

const ZOOMED_WIDTHS = [
  { width: 1590, expected: { tier: "wide", outline: "column", comments: "column" } },
  { width: 1100, expected: { tier: "medium", outline: "rail", comments: "column" } },
  { width: 800, expected: { tier: "narrow", outline: "rail", comments: "drawer" } },
] as const;

for (const { width, expected } of ZOOMED_WIDTHS) {
  for (const review of ["changes-only", "comment-and-changes"] as const) {
    test(`webview ${String(width)}px at 80% CSS zoom, review ${review}`, async ({ page }) => {
      const pageErrors: string[] = [];
      page.on("pageerror", (error) => pageErrors.push(error.message));
      await page.setViewportSize({ width, height: 1550 });
      await openWebview(page, "dark");
      await page.addStyleTag({ content: "html { zoom: 80%; }" });
      await waitForSent(page, { type: "ready" });
      await postDocument(page, {
        type: "load",
        document: {
          bytes: await buildPanelLayoutDocument(review),
          fileVersion: "v1",
          fileName: "nda.docx",
        },
        author: "Test Author",
        mode: "editing",
        locale: "en",
      });
      await waitForSent(page, { type: "loaded", fileVersion: "v1" });
      await waitForPanels(page, review);

      if (review === "changes-only") {
        expect((await readPanelState(page))?.comments).toBe("hidden");
        await page.getByTestId("toolbar-comments-toggle").click();
      }
      await expect.poll(() => readPanelState(page)).toEqual(expected);
      if (expected.comments === "column") {
        await page.locator('[data-folio-comments-surface="column"]').waitFor();
      }
      await expectPanelsDoNotOverlap(page, expected);
      await page.getByTestId("toolbar-toggle-ruler").click();
      const rulerAndViewport = await page.evaluate(() => {
        const ruler = document.querySelector<HTMLElement>('[data-testid="folio-horizontal-ruler"]');
        const scroll = document.querySelector<HTMLElement>("[data-folio-scroll]");
        if (!ruler || !scroll) return null;
        const rulerBox = ruler.getBoundingClientRect();
        const scrollBox = scroll.getBoundingClientRect();
        return {
          rulerLeft: rulerBox.left,
          rulerRight: rulerBox.right,
          scrollLeft: scrollBox.left,
          scrollRight: scrollBox.right,
        };
      });
      expect(rulerAndViewport).not.toBeNull();
      if (rulerAndViewport) {
        expect(
          Math.abs(rulerAndViewport.rulerLeft - rulerAndViewport.scrollLeft),
        ).toBeLessThanOrEqual(1);
        expect(
          Math.abs(rulerAndViewport.rulerRight - rulerAndViewport.scrollRight),
        ).toBeLessThanOrEqual(1);
      }
      expect(pageErrors).toEqual([]);
    });
  }
}
