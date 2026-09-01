import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

// DocxEditorRef.getEditorRef() returns a PagedEditorRef-shaped handle in both
// adapters. React sources it from the real PagedEditor component; Vue has no
// ported PagedEditor, so useDocxEditorRefApi.ts synthesizes an equivalent
// handle from the same primitives (see the "getEditorRef" pairedNote in
// scripts/parity/parity.contract.json). Exercise two of its methods
// end-to-end — dispatch (through the nested ref, not the raw ProseMirror
// view.dispatch other specs use) and getPageNumberForPmPos — and confirm both
// adapters agree.
forEachAdapter("getEditorRef(): dispatch + getPageNumberForPmPos", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  const before = await page.evaluate(() => window.__folioParity?.getDocumentText() ?? "");

  const inserted = await page.evaluate(
    () => window.__folioParity?.insertTextViaPagedEditorRef("via-paged-ref-marker") ?? false,
  );
  expect(inserted).toBe(true);

  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""))
    .toContain("via-paged-ref-marker");

  const after = await page.evaluate(() => window.__folioParity?.getDocumentText() ?? "");
  expect(after).not.toBe(before);

  // The selection anchor sits somewhere on the (now laid-out) first page.
  const pageNumber = await page.evaluate(
    () => window.__folioParity?.getPageNumberForSelection() ?? 0,
  );
  expect(pageNumber).toBeGreaterThanOrEqual(1);
});

forEachAdapter(
  "block geometry is snapshot-aligned and scroll-root relative",
  async (adapter, { page }) => {
    await openEditor(page, adapter);
    await ensureLiveView(page);

    const geometry = await page.evaluate(() => window.__folioParity?.readBlockGeometry());
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
