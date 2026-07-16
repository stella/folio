import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("edits dropdown and date content controls", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);
  expect(await page.evaluate(() => window.__folioParity?.setupContentControls() ?? false)).toBe(
    true,
  );

  const dropdown = page.locator('.paged-editor__pages [data-sdt-type="dropdown"]').first();
  await expect(dropdown).toBeVisible();
  await dropdown.click();
  await page.getByRole("menuitem", { name: "New York" }).click();
  await expect(dropdown).toContainText("New York");

  const date = page.locator('.paged-editor__pages [data-sdt-type="date"]').first();
  await expect(date).toBeVisible();
  await date.click();
  await page.locator('input[type="date"]').fill("2026-06-02");
  await expect(date).toContainText("2026-06-02");

  expect(await page.evaluate(() => window.__folioParity?.save() ?? 0)).toBeGreaterThan(0);
});
