import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

// hasPendingChanges tracks whether the live editor has edits not yet serialized:
// clean on load, dirty after an edit, clean again after a save. Both adapters
// must agree (React derives it from the change tracker + comments-dirty flag;
// Vue from equivalent doc-dirty + comments-dirty flags with the same set/clear
// points).
forEachAdapter("hasPendingChanges: clean -> edit -> save", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  // Freshly loaded document: nothing pending.
  expect(await page.evaluate(() => window.__folioParity?.hasPendingChanges() ?? true)).toBe(false);

  // A doc edit makes changes pending.
  const inserted = await page.evaluate(() => window.__folioParity?.insertText(" pending") ?? false);
  expect(inserted).toBe(true);
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.hasPendingChanges() ?? false))
    .toBe(true);

  // A successful save clears the pending flag.
  const byteLength = await page.evaluate(() => window.__folioParity?.save() ?? 0);
  expect(byteLength).toBeGreaterThan(1000);
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.hasPendingChanges() ?? true))
    .toBe(false);
});
