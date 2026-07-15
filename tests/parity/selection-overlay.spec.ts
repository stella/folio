import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("selection: paints a selected text range", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const selected = await page.evaluate(() => window.__folioParity?.selectFirstWord() ?? false);
  expect(selected).toBe(true);

  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.countSelectionRects() ?? 0))
    .toBeGreaterThan(0);
});
