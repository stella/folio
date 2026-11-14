import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

// The ref's autocomplete API shows ghost text at the caret in the canonical
// session. Typing its next characters consumes them while the ghost follows
// the caret, a word-wise accept inserts the next word as typed text, and
// Escape dismisses it. Both adapters paint the same overlay.
forEachAdapter(
  "projects autocomplete suggestions onto painted pages",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    await ensureLiveView(page);

    // Put the caret at the end of the first painted run of body text.
    const run = page
      .locator(".paged-editor__pages .layout-page-content span[data-pm-start]")
      .first();
    const box = await run.boundingBox();
    if (!box) throw new Error("No painted body text");
    await page.mouse.click(box.x + box.width - 0.2, box.y + box.height / 2);

    expect(
      await page.evaluate(() =>
        window.__folioParity?.startAutocomplete(" suggested text"),
      ),
    ).toBe(true);
    const overlay = page.locator("[data-folio-autocomplete-overlay]");
    const ghost = overlay.locator(".folio-autocomplete-ghost");
    await expect(ghost).toBeVisible();
    await expect(ghost).toContainText("suggested text");
    await expect(overlay.locator(".folio-autocomplete-caret")).toHaveClass(
      /\bfolio-autocomplete-caret--streaming\b/u,
    );

    await page.keyboard.type(" s");
    await expect
      .poll(() =>
        page.evaluate(() => window.__folioParity?.readAutocompleteText()),
      )
      .toBe("uggested text");
    await expect(ghost).toContainText("uggested text");

    expect(
      await page.evaluate(() => window.__folioParity?.finishAutocomplete()),
    ).toBe(true);
    await expect(overlay.locator(".folio-autocomplete-caret")).not.toHaveClass(
      /\bfolio-autocomplete-caret--streaming\b/u,
    );

    expect(
      await page.evaluate(() => window.__folioParity?.acceptAutocompleteWord()),
    ).toBe(true);
    expect(
      await page.evaluate(() => window.__folioParity?.readAutocompleteText()),
    ).toBe("text");
    await expect
      .poll(() =>
        page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""),
      )
      .toContain(" suggested ");

    await page.keyboard.press("Escape");
    await expect(overlay).toHaveCount(0);
    expect(
      await page.evaluate(() => window.__folioParity?.readAutocompleteText()),
    ).toBeNull();
  },
);
