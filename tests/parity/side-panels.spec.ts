/**
 * Side-panel layout of both editor adapters: at each editor width,
 * with and without review content, the page, the document outline and the
 * comments occupy disjoint columns, each panel takes the presentation its
 * tier calls for, and the drawers open over the page and close again.
 *
 * The same oracle also runs against the VS Code bundle in
 * `packages/editor-web/tests/panel-layout.spec.ts`.
 */

import { expect } from "@playwright/test";
import { ensureLiveView, forEachAdapter, openEditor, type AdapterFixture } from "./parity-fixture";
import type { Page } from "@playwright/test";

import {
  PANEL_LAYOUT_CASES,
  afterPanelDrawersDismissed,
  expectPanelDrawers,
  expectPanelsDoNotOverlap,
  readPanelState,
  waitForPanels,
} from "../support/panelLayoutAssertions";
import {
  buildPanelLayoutDocument,
  type PanelLayoutReview,
  type PanelLayoutSections,
} from "../support/panelLayoutDocument";

const VIEWPORT_HEIGHT = 900;

const loadDocument = async (
  page: Page,
  adapter: AdapterFixture,
  review: PanelLayoutReview,
  sections: PanelLayoutSections = "portrait",
) => {
  const bytes = await buildPanelLayoutDocument(review, sections);
  await page.route("**/fixtures/panel-layout.docx", (route) =>
    route.fulfill({
      contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      body: Buffer.from(bytes),
    }),
  );
  await openEditor(page, adapter, "panel-layout.docx");
  await page.evaluate(() => document.fonts.ready);
  await ensureLiveView(page);
  await waitForPanels(page, review);
};

for (const { width, review, expected } of PANEL_LAYOUT_CASES) {
  forEachAdapter(
    `${String(width)}px, review ${review}: ${expected.tier} tier`,
    async (adapter, { page }) => {
      await page.setViewportSize({ width, height: VIEWPORT_HEIGHT });
      await loadDocument(page, adapter, review);

      expect(await readPanelState(page)).toEqual(expected);
      await expectPanelsDoNotOverlap(page, expected);
      if (review !== "none") {
        await expect(page.getByTestId("toolbar-comments-count")).toHaveText("1");
      }

      await expectPanelDrawers(page, expected);
      const afterDismissal = afterPanelDrawersDismissed(expected);
      expect(await readPanelState(page)).toEqual(afterDismissal);
      await expectPanelsDoNotOverlap(page, afterDismissal);
    },
  );
}

forEachAdapter("tracked changes alone leave comments closed", async (adapter, { page }) => {
  await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
  await loadDocument(page, adapter, "changes-only");
  await expect(page.getByTestId("toolbar-comments-toggle")).toHaveAttribute(
    "aria-pressed",
    "false",
  );
  await expect(page.getByTestId("toolbar-comments-count")).toHaveCount(0);
  expect(await readPanelState(page)).toEqual({
    tier: "wide",
    outline: "column",
    comments: "hidden",
  });
});

forEachAdapter(
  "the widest page in mixed sections controls panel room",
  async (adapter, { page }) => {
    await page.setViewportSize({ width: 1400, height: VIEWPORT_HEIGHT });
    await loadDocument(page, adapter, "comment-and-changes", "landscape-then-portrait");

    await expect
      .poll(() =>
        page
          .locator(".layout-page")
          .evaluateAll((pages) =>
            Math.max(...pages.map((element) => element.getBoundingClientRect().width)),
          ),
      )
      .toBeGreaterThan(1000);
    const expected = { tier: "narrow", outline: "column", comments: "drawer" } as const;
    await expect.poll(() => readPanelState(page)).toEqual(expected);
    await expectPanelsDoNotOverlap(page, expected);
  },
);

forEachAdapter(
  "a dismissed comments drawer stays closed after the editor widens",
  async (adapter, { page }) => {
    await page.setViewportSize({ width: 800, height: VIEWPORT_HEIGHT });
    await loadDocument(page, adapter, "comment-and-changes");
    const toggle = page.getByTestId("toolbar-comments-toggle");
    await toggle.click();
    await expect(page.locator('[data-folio-comments-surface="drawer"]')).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.locator('[data-folio-comments-surface="drawer"]')).toBeHidden();
    await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
    const expected = { tier: "wide", outline: "column", comments: "hidden" } as const;
    await expect.poll(() => readPanelState(page)).toEqual(expected);
    await expectPanelsDoNotOverlap(page, expected);
  },
);

forEachAdapter("the outline column is keyboard navigable", async (adapter, { page }) => {
  await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
  await loadDocument(page, adapter, "none");
  const items = page.locator('[data-folio-outline-surface="column"] .folio-outline-item');
  await expect(items).toHaveCount(4);
  // One tab stop for the whole list.
  await expect(items.and(page.locator('[tabindex="0"]'))).toHaveCount(1);

  await items.first().focus();
  await page.keyboard.press("ArrowDown");
  await expect(items.nth(1)).toBeFocused();
  await page.keyboard.press("End");
  await expect(items.nth(3)).toBeFocused();
  await page.keyboard.press("Enter");
  await expect(items.nth(3)).toHaveAttribute("aria-current", "true");
  // The jump scrolled the document to the heading.
  await expect
    .poll(() => page.locator("[data-folio-scroll]").evaluate((element) => element.scrollTop))
    .toBeGreaterThan(0);
});

forEachAdapter("toggling the comments column re-centres the page", async (adapter, { page }) => {
  await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
  await loadDocument(page, adapter, "comment-and-changes");
  const toggle = page.getByTestId("toolbar-comments-toggle");
  await expect(toggle).toHaveAttribute("aria-pressed", "true");

  await toggle.click();
  await expect(toggle).toHaveAttribute("aria-pressed", "false");
  const closed = { tier: "wide", outline: "column", comments: "hidden" } as const;
  expect(await readPanelState(page)).toEqual(closed);
  await expectPanelsDoNotOverlap(page, closed);

  await toggle.click();
  await expect(toggle).toHaveAttribute("aria-pressed", "true");
  const reopened = { tier: "wide", outline: "column", comments: "column" } as const;
  expect(await readPanelState(page)).toEqual(reopened);
  await expectPanelsDoNotOverlap(page, reopened);
});
