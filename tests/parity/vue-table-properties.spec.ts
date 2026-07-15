import { test } from "@playwright/test";

import { ensureLiveView, expect, openEditor, type AdapterFixture } from "./parity-fixture";

const VUE_ADAPTER = {
  name: "vue",
  baseUrl: "http://localhost:4201",
} as const satisfies AdapterFixture;

test("Vue applies table properties from the table options menu", async ({ page }) => {
  await openEditor(page, VUE_ADAPTER);
  await ensureLiveView(page);
  expect(await page.evaluate(() => window.__folioParity?.insertTable(2, 2) ?? false)).toBe(true);

  await page.getByTitle("Table options").click();
  await page.getByRole("button", { name: "Table properties" }).click();

  const dialog = page.getByRole("dialog", { name: "Table Properties" });
  await expect(dialog).toBeVisible();
  await dialog.locator("select").nth(0).selectOption("pct");
  await dialog.locator('input[type="number"]').fill("2500");
  await dialog.locator("select").nth(1).selectOption("center");
  await dialog.getByRole("button", { name: "Apply" }).click();

  await expect(dialog).toBeHidden();
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getCurrentTableProperties() ?? null))
    .toEqual({ width: 2500, widthType: "pct", justification: "center" });
});
