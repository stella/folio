import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

forEachAdapter(
  "smoke: typed text reaches the document",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    await ensureLiveView(page);

    await page.locator(".layout-paragraph").first().click();
    await page.keyboard.press("End");
    await page.keyboard.type(" smoke");

    await expect
      .poll(() =>
        page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""),
      )
      .toContain("smoke");
  },
);
