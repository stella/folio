import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter(
  "projects autocomplete suggestions onto painted pages",
  async (adapter, { page }) => {
    await openEditor(page, adapter);
    await ensureLiveView(page);

    expect(
      await page.evaluate(() => window.__folioParity?.startAutocomplete(" suggested text")),
    ).toBe(true);

    const overlay = page.locator("[data-folio-autocomplete-overlay]");
    await expect(overlay.locator(".folio-autocomplete-ghost")).toBeVisible();
    await expect(overlay.locator(".folio-autocomplete-ghost")).toContainText("suggested text");
    await expect(overlay.locator(".folio-autocomplete-caret")).toHaveClass(
      /\bfolio-autocomplete-caret--streaming\b/u,
    );

    expect(await page.evaluate(() => window.__folioParity?.finishAutocomplete())).toBe(true);
    await expect(overlay.locator(".folio-autocomplete-caret")).not.toHaveClass(
      /\bfolio-autocomplete-caret--streaming\b/u,
    );

    expect(await page.evaluate(() => window.__folioParity?.clearAutocomplete())).toBe(true);
    await expect(overlay).toHaveCount(0);
  },
);
