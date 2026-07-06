import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

// Both adapters install the anonymization decoration plugin (always on, inert
// until terms are pushed) and mount an overlay that projects the plugin's match
// ranges onto container-relative highlight rects. Pushing a term matching the
// first word through the bridge proves the full path in each adapter: plugin
// match recompute -> overlay re-projection -> painted `.folio-anonymization-term`
// rect. Because both playgrounds parse the same fixture with the same core
// plugin + projection, a highlight appears in React and Vue alike.
forEachAdapter("anonymization: pushed term paints a highlight rect", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const pushed = await page.evaluate(() => window.__folioParity?.anonymizeFirstWord() ?? false);
  expect(pushed).toBe(true);

  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.countAnonymizationRects() ?? 0), {
      timeout: 15_000,
    })
    .toBeGreaterThan(0);
});
