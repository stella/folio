/**
 * React panel regressions for document loading and changing outline entries.
 * Shared panel geometry and drawer behavior live in tests/parity/side-panels.spec.ts.
 */

import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";

import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import {
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
      const expected = { tier: "narrow", outline: "rail", comments: "drawer" } as const;
      await expect.poll(() => readPanelState(page)).toEqual(expected);
      await expectPanelsDoNotOverlap(page, expected);
    });
  }

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
});
