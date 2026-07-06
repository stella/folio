import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("smoke: applies bold through the bridge", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const applied = await page.evaluate(() => window.__folioParity?.boldFirstWord() ?? false);
  expect(applied).toBe(true);
});
