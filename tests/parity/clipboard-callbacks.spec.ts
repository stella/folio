import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

forEachAdapter(
  "fires host callbacks for clipboard events",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    await ensureLiveView(page);
    const paragraph = page.locator(".layout-paragraph").first();
    // Each shortcut acts on a selection inside the page, as a user's does.
    for (const [kind, key] of [
      ["copy", "c"],
      ["cut", "x"],
      ["paste", "v"],
    ] as const) {
      await paragraph.click();
      await page.keyboard.press("Home");
      await page.keyboard.press("Shift+End");
      await page.keyboard.press(`ControlOrMeta+${key}`);
      await expect
        .poll(() =>
          page.evaluate(
            (clipboardKind) =>
              window.__folioParity?.getClipboardCallbackCount(clipboardKind) ??
              0,
            kind,
          ),
        )
        .toBe(1);
    }
  },
);
