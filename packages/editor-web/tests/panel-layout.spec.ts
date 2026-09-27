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
