import { test } from "@playwright/test";

import { ensureLiveView, expect, forEachAdapter, openEditor, VUE_ADAPTER } from "./parity-fixture";

forEachAdapter(
  "find: the dialog highlights matches without changing selection",
  async (adapter, { page }) => {
    await openEditor(page, adapter);
    await ensureLiveView(page);
    await page.locator(".layout-page").first().click();
    expect(await page.evaluate(() => window.__folioParity?.selectFirstWord())).toBe(true);
    const before = await page.evaluate(() => ({
      selection: window.__folioParity?.getSelectedText(),
      text: window.__folioParity?.getDocumentText(),
    }));
    expect(before.selection).toBeTruthy();
    expect(before.selection?.toLowerCase()).not.toBe("editor");

    await page.keyboard.press("ControlOrMeta+f");
    const dialog = page.getByTestId("find-replace-dialog");
    await expect(dialog).toBeVisible();
    await dialog.locator("input").first().fill("editor");

    const current = page.locator('[data-find-match-index="0"][data-current="true"]').first();
    await expect(current).toBeVisible();
    await expect(dialog).toContainText("1 / 3");
    await dialog.locator("input").first().press("Enter");
    await expect(
      page.locator('[data-find-match-index="1"][data-current="true"]').first(),
    ).toBeVisible();
    await expect(page.locator('[data-find-match-index="0"][data-current="true"]')).toHaveCount(0);
    await expect(dialog).toContainText("2 / 3");
    expect(
      await page.evaluate(() => ({
        selection: window.__folioParity?.getSelectedText(),
        text: window.__folioParity?.getDocumentText(),
      })),
    ).toEqual(before);
  },
);

test("find/replace: Vue replaces the current highlighted match [vue]", async ({ page }) => {
  await openEditor(page, VUE_ADAPTER);
  await ensureLiveView(page);
  await page.locator(".layout-page").first().click();

  await page.keyboard.press("ControlOrMeta+f");
  const dialog = page.getByTestId("find-replace-dialog");
  await dialog.getByTestId("find-input").fill("editor");
  await expect(
    page.locator('[data-find-match-index="0"][data-current="true"]').first(),
  ).toBeVisible();
  await expect(dialog).toContainText("1 / 3");

  await dialog.getByTestId("find-replace-toggle").click();
  await dialog.getByTestId("replace-input").fill("folioReplacement");
  await dialog.getByTestId("replace-current").click();

  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""))
    .toContain("folioReplacement");
});
