import { test } from "@playwright/test";

import { ensureLiveView, expect, openEditor, type AdapterFixture } from "../parity/parity-fixture";

const VUE_ADAPTER = {
  name: "vue",
  baseUrl: "http://localhost:4201",
} as const satisfies AdapterFixture;

test("Vue applies and clears a watermark from the Insert menu", async ({ page }) => {
  await openEditor(page, VUE_ADAPTER);
  await ensureLiveView(page);

  await page.getByRole("button", { name: "Insert", exact: true }).click();
  await page.getByRole("button", { name: "Watermark", exact: true }).click();

  const dialog = page.getByRole("dialog", { name: "Watermark" });
  await expect(dialog).toBeVisible();
  await dialog.locator("select").selectOption("text");
  await dialog.getByLabel("Text", { exact: true }).fill("DRAFT");
  await dialog.getByRole("button", { name: "Apply" }).click();

  await expect(dialog).toBeHidden();
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getTextWatermark() ?? null))
    .toBe("DRAFT");
  expect(await page.evaluate(() => window.__folioParity?.hasPendingChanges() ?? false)).toBe(true);

  await page.getByRole("button", { name: "Insert", exact: true }).click();
  await page.getByRole("button", { name: "Watermark", exact: true }).click();
  await dialog.locator("select").selectOption("none");
  await dialog.getByRole("button", { name: "Apply" }).click();

  await expect(dialog).toBeHidden();
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getTextWatermark() ?? null))
    .toBeNull();
});
