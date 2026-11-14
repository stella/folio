import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

// hasPendingChanges tracks whether the live editor has edits not yet serialized:
// clean on load, dirty after an edit, clean again after a save. Both adapters
// must agree (React derives it from the change tracker + comments-dirty flag;
// Vue from equivalent doc-dirty + comments-dirty flags with the same set/clear
// points).
forEachAdapter(
  "hasPendingChanges: clean -> edit -> save",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    await ensureLiveView(page);

    // Freshly loaded document: nothing pending.
    expect(
      await page.evaluate(
        () => window.__folioParity?.hasPendingChanges() ?? true,
      ),
    ).toBe(false);

    // A typed edit makes changes pending.
    await page.locator(".layout-paragraph").first().click();
    await page.keyboard.press("End");
    await page.keyboard.type(" pending");
    await expect
      .poll(() =>
        page.evaluate(() => window.__folioParity?.hasPendingChanges() ?? false),
      )
      .toBe(true);

    // A successful save clears the pending flag.
    const byteLength = await page.evaluate(
      () => window.__folioParity?.save() ?? 0,
    );
    expect(byteLength).toBeGreaterThan(1000);
    await expect
      .poll(() =>
        page.evaluate(() => window.__folioParity?.hasPendingChanges() ?? true),
      )
      .toBe(false);
  },
);
