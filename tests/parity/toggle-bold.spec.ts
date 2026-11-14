import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

forEachAdapter(
  "smoke: applies bold through the bridge",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    await ensureLiveView(page);

    const applied = await page.evaluate(
      () => window.__folioParity?.boldFirstWord() ?? false,
    );
    expect(applied).toBe(true);
  },
);
