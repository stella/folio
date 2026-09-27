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
  PANEL_LAYOUT_HEADINGS,
  type PanelLayoutReview,
  type PanelLayoutSections,
} from "../support/panelLayoutDocument";

declare global {
  var __folioPlayground: { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const VIEWPORT_HEIGHT = 900;

type LoadDocumentOptions = {
  review: PanelLayoutReview;
  sections?: PanelLayoutSections;
  source?: "imperative" | "buffer-prop";
};

const loadDocument = async (
  page: Page,
  { review, sections = "portrait", source = "imperative" }: LoadDocumentOptions,
) => {
  const bytes = await buildPanelLayoutDocument(review, sections);
  if (source === "buffer-prop") {
    await page.route("**/fixtures/panel-layout.docx", (route) =>
      route.fulfill({ body: Buffer.from(bytes) }),
    );
    // The fixture path starts with no document and loads through documentBuffer.
    await page.goto("/?file=panel-layout.docx");
  } else {
    await page.goto("/");
    await page.waitForFunction(() => globalThis.__folioPlayground?.getEditorRef() != null);
    await page.evaluate(
      async (array) => {
        await globalThis.__folioPlayground
          ?.getEditorRef()
          ?.loadDocumentBuffer(new Uint8Array(array));
      },
      [...bytes],
    );
  }
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
      await loadDocument(page, { review });

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

  // Previously only an already-mounted editor was exercised. Both loading paths
  // must measure mixed sections instead of using the final section's page width.
  for (const source of ["imperative", "buffer-prop"] as const) {
    test(`the widest page in mixed sections controls panel room via ${source}`, async ({
      page,
    }) => {
      await page.setViewportSize({ width: 1400, height: VIEWPORT_HEIGHT });
      await loadDocument(page, {
        review: "comment-and-changes",
        sections: "landscape-then-portrait",
        source,
      });

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
  }

  test("a dismissed comments drawer stays closed after the editor widens", async ({ page }) => {
    await page.setViewportSize({ width: 800, height: VIEWPORT_HEIGHT });
    await loadDocument(page, { review: "comment-and-changes" });
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
    await loadDocument(page, { review: "none" });
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

  test("Tab reaches the outline after focused headings are removed", async ({ page }) => {
    await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
    await loadDocument(page, { review: "none" });
    const outline = page.locator('[data-folio-outline-surface="column"]');
    const items = outline.locator(".folio-outline-item");
    await expect(items).toHaveCount(4);
    await items.last().focus();

    // Keep focus in the outline while a document edit removes its last entries.
    await page.evaluate(
      (removedHeadings) => {
        const view = globalThis.__folioPlayground?.getEditorRef()?.getEditorRef()?.getView();
        if (!view) throw new Error("Missing editor view");
        const transaction = view.state.tr;
        view.state.doc.descendants((node, pos) => {
          if (node.isTextblock && removedHeadings.some((heading) => heading === node.textContent)) {
            transaction.delete(
              transaction.mapping.map(pos),
              transaction.mapping.map(pos + node.nodeSize),
            );
            return false;
          }
          return true;
        });
        view.dispatch(transaction);
      },
      PANEL_LAYOUT_HEADINGS.slice(2).map(({ text }) => text),
    );

    await expect(items).toHaveCount(2);
    await expect(items.and(page.locator('[tabindex="0"]'))).toHaveCount(1);
    await outline.evaluate((element) => {
      const before = document.createElement("button");
      before.textContent = "Before outline";
      element.before(before);
      before.focus();
    });
    await page.keyboard.press("Tab");
    await expect(items.last()).toBeFocused();
    await page.keyboard.press("ArrowUp");
    await expect(items.first()).toBeFocused();
  });

  test("toggling the comments column re-centres the page", async ({ page }) => {
    await page.setViewportSize({ width: 1500, height: VIEWPORT_HEIGHT });
    await loadDocument(page, { review: "comment-and-changes" });
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
