import { expect, forEachAdapter, openCanonicalEditor } from "./parity-fixture";

forEachAdapter("smoke: save returns docx bytes", async (adapter, { page }) => {
  await openCanonicalEditor(page, adapter);
  const byteLength = await page.evaluate(
    () => window.__folioParity?.save() ?? 0,
  );
  expect(byteLength).toBeGreaterThan(1000);
});
