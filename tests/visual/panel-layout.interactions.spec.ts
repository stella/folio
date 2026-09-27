/**
 * Side-panel layout of the React editor (playground): at each editor width,
 * with and without review content, the page, the document outline and the
 * comments occupy disjoint columns, each panel takes the presentation its
 * tier calls for, and the drawers open over the page and close again.
 *
 * No layout test existed before this one: the outline and the comments were
 * positioned independently (the outline pinned to the end edge over the
 * page, the comments beside the page), and nothing measured them against
 * each other, so the two could overlap the page and each other in a narrow
 * pane. The same oracle runs against the VS Code bundle in
 * `packages/editor-web/tests/panel-layout.spec.ts`.
 */

import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
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

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const VIEWPORT_HEIGHT = 900;

const loadDocument = async (
  page: Page,
  review: PanelLayoutReview,
  sections: PanelLayoutSections = "portrait",
) => {
  const bytes = [...(await buildPanelLayoutDocument(review, sections))];
  await page.goto("/");
  await page.waitForFunction(() => globalThis.__folioPlayground?.getEditorRef() != null);
  await page.evaluate(async (array) => {
    await globalThis.__folioPlayground?.getEditorRef()?.loadDocumentBuffer(new Uint8Array(array));
  }, bytes);
  await page.waitForFunction(() => document.querySelectorAll(".layout-page").length >= 1);
  await page.evaluate(() => document.fonts.ready);
  // The outline reads headings from the body view, created lazily; a host
  // (the VS Code webview does) creates it up front.
  await page.evaluate(() =>
    globalThis.__folioPlayground?.getEditorRef()?.ensureEditorView({ focus: false }),
  );
  await waitForPanels(page, review);
};

test.describe("side panel layout", () => {
  for (const { width, review, expected } of PANEL_LAYOUT_CASES) {
    test(`${String(width)}px, review ${review}: ${expected.tier} tier`, async ({ page }) => {
      await page.setViewportSize({ width, height: VIEWPORT_HEIGHT });
      await loadDocument(page, review);

      expect(await readPanelState(page)).toEqual(expected);
      await expectPanelsDoNotOverlap(page, expected);
      if (review !== "none") {
        await expect(page.getByTestId("toolbar-comments-count")).toHaveText("1");
      }
      await expectPanelDrawers(page, expected);
      const afterDismissal = afterPanelDrawersDismissed(expected);
      expect(await readPanelState(page)).toEqual(afterDismissal);
      await expectPanelsDoNotOverlap(page, afterDismissal);
    });
  }

  test("the widest page in mixed sections controls panel room", async ({ page }) => {
    await page.setViewportSize({ width: 1400, height: VIEWPORT_HEIGHT });
    await loadDocument(page, "comment-and-changes", "landscape-then-portrait");

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
  });

  test("a dismissed comments drawer stays closed after the editor widens", async ({ page }) => {
    await page.setViewportSize({ width: 800, height: VIEWPORT_HEIGHT });
    await loadDocument(page, "comment-and-changes");
    const toggle = page.getByTestId("toolbar-comments-toggle");
    await toggle.click();
    await expect(page.locator('[data-folio-comments-surface="drawer"]')).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.locator('[data-folio-comments-surface="drawer"]')).toBeHidden();
    await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
    const expected = { tier: "wide", outline: "column", comments: "hidden" } as const;
    await expect.poll(() => readPanelState(page)).toEqual(expected);
    await expectPanelsDoNotOverlap(page, expected);
  });

  test("the outline column is keyboard navigable", async ({ page }) => {
    await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
    await loadDocument(page, "none");
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

  test("toggling the comments column re-centres the page", async ({ page }) => {
    await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
    await loadDocument(page, "comment-and-changes");
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
});
