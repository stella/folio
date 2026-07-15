import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

forEachAdapter("fires host callbacks for clipboard events", async (adapter, { page }) => {
  await openEditor(page, adapter);
  await ensureLiveView(page);

  for (const kind of ["copy", "cut", "paste"] as const) {
    expect(
      await page.evaluate(
        (clipboardKind) => window.__folioParity?.dispatchClipboardEvent(clipboardKind) ?? 0,
        kind,
      ),
    ).toBe(1);
  }
});
