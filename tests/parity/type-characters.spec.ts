import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("smoke: inserts text through the bridge", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const inserted = await page.evaluate(() => window.__folioParity?.insertText(" smoke") ?? false);
  expect(inserted).toBe(true);

  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""))
    .toContain("smoke");
});
