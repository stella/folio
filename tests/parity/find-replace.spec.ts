import { test } from "@playwright/test";

import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("find: the dialog selects a live document match", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  await page.keyboard.press("Control+f");
  const dialog = page.getByTestId("find-replace-dialog");
  await expect(dialog).toBeVisible();
  await dialog.locator("input").first().fill("editor");

  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getSelectedText().toLowerCase() ?? ""))
    .toBe("editor");
});

test("find/replace: Vue replaces the selected live match", async ({ page }) => {
  const vue = { name: "vue", baseUrl: "http://localhost:4201" } as const;
  await openEditor(page, vue);
  await ensureLiveView(page);

  await page.keyboard.press("Control+f");
  const dialog = page.getByTestId("find-replace-dialog");
  await dialog.getByTestId("find-input").fill("editor");
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getSelectedText().toLowerCase() ?? ""))
    .toBe("editor");

  await dialog.getByTestId("find-replace-toggle").click();
  await dialog.getByTestId("replace-input").fill("folioReplacement");
  await dialog.getByTestId("replace-current").click();

  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""))
    .toContain("folioReplacement");
});
