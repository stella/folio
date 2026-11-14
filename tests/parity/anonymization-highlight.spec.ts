import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openEditor,
} from "./parity-fixture";

// The `anonymizationTerms` prop highlights a term wherever it occurs in the
// shown text. In the canonical session both adapters read the occurrences the
// editor derives from the text and paint them through the shared range
// projection; the occurrences come back as text positions, and setting terms
// leaves the document unchanged.
forEachAdapter(
  "anonymization: a term from the prop paints a highlight rect",
  async (adapter, { page }) => {
    await openEditor(page, adapter, undefined, { session: "canonical" });
    await ensureLiveView(page);

    const word = await page.evaluate(
      () => window.__folioParity?.anonymizeFirstWord() ?? null,
    );
    expect(word).not.toBeNull();

    await expect
      .poll(
        () =>
          page.evaluate(
            () => window.__folioParity?.countAnonymizationRects() ?? 0,
          ),
        {
          timeout: 15_000,
        },
      )
      .toBeGreaterThan(0);
    const starts = await page.evaluate(
      () => window.__folioParity?.getAnonymizationMatchStarts() ?? [],
    );
    expect(starts.length).toBeGreaterThan(0);
    for (const start of starts) expect(start).toMatch(/^\S+:\d+$/u);
    expect(
      await page.evaluate(
        () => window.__folioParity?.hasPendingChanges() ?? true,
      ),
    ).toBe(false);
  },
);
