import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

// Both adapters route Insert > Table through the same core `insertTableInView`
// helper (React via the toolbar Insert group, Vue via the title-bar MenuBar's
// TableGridInline picker). Driving that helper through the shared bridge proves
// the insert produces the same structural change (one new `table` node) in each.
forEachAdapter("insert: adds a table node through the bridge", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const before = await page.evaluate(() => window.__folioParity?.countTables() ?? -1);
  const inserted = await page.evaluate(() => window.__folioParity?.insertTable(2, 3) ?? false);
  const after = await page.evaluate(() => window.__folioParity?.countTables() ?? -1);

  expect(inserted).toBe(true);
  expect(after).toBe(before + 1);
});
