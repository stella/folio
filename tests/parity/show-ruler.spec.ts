import { expect, forEachAdapter, openEditor } from "./parity-fixture";

// Both playgrounds mount the editor with `showRuler` on, so the horizontal and
// vertical rulers must render in both adapters. Guards the `showRuler` /
// `rulerUnit` prop pairing (React ≡ Vue): the rulers share the
// `docx-horizontal-ruler` / `docx-vertical-ruler` root classes across adapters.
forEachAdapter("showRuler renders the rulers", async (adapter, { page }) => {
  await openEditor(page, adapter);

  await expect(page.locator(".docx-horizontal-ruler").first()).toBeVisible();
  await expect(page.locator(".docx-vertical-ruler").first()).toBeVisible();
});
