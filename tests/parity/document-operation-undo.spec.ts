import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("document operations: committed batch can be undone", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  expect(await page.evaluate(() => window.__folioParity?.applyAndUndoDocumentOperation())).toBe(
    true,
  );
});
