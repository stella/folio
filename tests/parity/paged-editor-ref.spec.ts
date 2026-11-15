import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
  openEditor,
} from "./parity-fixture";

// DocxEditorRef.getPositionPage resolves a public text position (story,
// paragraph id, offset) to the rendered page that shows it, in the canonical
// session, and both adapters agree.
forEachAdapter(
  "getPositionPage resolves paragraph positions to rendered pages",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    await ensureLiveView(page);

    await expect
      .poll(() =>
        page.evaluate(
          () => window.__folioParity?.getParagraphEdgePages() ?? null,
        ),
      )
      .not.toBeNull();
    const edges = await page.evaluate(
      () => window.__folioParity?.getParagraphEdgePages() ?? null,
    );
    const totalPages = await page.evaluate(
      () => window.__folioParity?.getTotalPages() ?? 0,
    );
    expect(edges?.first).toBe(1);
    expect(edges?.last).toBeGreaterThanOrEqual(1);
    expect(edges?.last).toBeLessThanOrEqual(totalPages);

    const unknown = await page.evaluate(() => {
      const bridge = window.__folioParity;
      return bridge === undefined
        ? "no bridge"
        : bridge.getUnknownPositionPage();
    });
    expect(unknown).toBeNull();
  },
);

forEachAdapter(
  "block geometry is snapshot-aligned and scroll-root relative",
  async (adapter, { page }) => {
    await openEditor(page, adapter);
    await ensureLiveView(page);

    const geometry = await page.evaluate(() =>
      window.__folioParity?.readBlockGeometry(),
    );
    expect(geometry).toBeDefined();
    expect(geometry?.missingIsNull).toBe(true);
    expect(geometry?.hasScrollRoot).toBe(true);
    expect(geometry?.rects.length).toBeGreaterThan(1);

    const rects = geometry?.rects ?? [];
    // Snapshot order is document order, not vertical order: table cells can
    // share a top coordinate, and later page columns can reset it.
    for (const rect of rects) {
      expect(rect.blockId).toBe(rect.snapshotBlockId);
      expect(Number.isInteger(rect.page)).toBe(true);
      expect(rect.page).toBeGreaterThan(0);
      expect(Number.isFinite(rect.top)).toBe(true);
      expect(rect.top).toBeGreaterThanOrEqual(0);
      expect(Number.isFinite(rect.height)).toBe(true);
      expect(rect.height).toBeGreaterThan(0);
    }
  },
);
