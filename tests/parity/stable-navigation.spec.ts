import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

forEachAdapter(
  "navigation: stable block targets reveal on the resolved page",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    await ensureLiveView(page);

    const result = await page.evaluate(() =>
      window.__folioParity?.navigateToFirstBlock(),
    );

    expect(result?.shown).toBe(true);
    expect(result?.targetPage).toBeGreaterThan(0);
    expect(result?.currentPage).toBe(result?.targetPage);
  },
);
