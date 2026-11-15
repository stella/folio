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

forEachAdapter(
  "fires host callbacks for the context menu's cut and plain-text paste",
  async (adapter, { page }) => {
    await page
      .context()
      .grantPermissions(["clipboard-read", "clipboard-write"], {
        origin: new URL(adapter.baseUrl).origin,
      });
    await openCanonicalEditor(page, adapter);
    await ensureLiveView(page);
    const paragraph = page.locator(".layout-paragraph").first();
    const word = paragraph.locator("span[data-pm-start]").first();
    // The paste leaves text for the cut to take.
    await page.evaluate(() => navigator.clipboard.writeText("Zeta"));
    for (const [kind, item] of [
      ["paste", /Paste (as Plain Text|without formatting)/iu],
      ["cut", /^\s*Cut(?![a-z])/u],
    ] as const) {
      await paragraph.click();
      await page.keyboard.press("Home");
      await page.keyboard.press("Shift+End");
      // Right-clicking inside the selection keeps it.
      await word.click({ button: "right" });
      // React's menu is a labelled ARIA menu; Vue's is a list of buttons.
      const menu = page
        .locator('[role="menu"][aria-label="Text editing menu"], .ctx-menu')
        .first();
      await menu.waitFor({ state: "visible" });
      await menu
        .locator('[role="menuitem"], button')
        .filter({ hasText: item })
        .first()
        .click();
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
