import { expect, forEachAdapter, openCanonicalEditor } from "./parity-fixture";

forEachAdapter(
  "smoke: editor mounts and lays out pages",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    const pageCount = await page.evaluate(
      () => window.__folioParity?.getTotalPages() ?? 0,
    );
    expect(pageCount).toBeGreaterThan(0);
  },
);
