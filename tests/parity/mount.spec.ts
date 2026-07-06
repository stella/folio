import { expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("smoke: editor mounts and lays out pages", async (adapter, { page }) => {
  await openEditor(page, adapter);
  const pageCount = await page.evaluate(() => window.__folioParity?.getTotalPages() ?? 0);
  expect(pageCount).toBeGreaterThan(0);
});
