import { expect, forEachAdapter } from "./parity-fixture";

forEachAdapter("paints the configured page margin guides", async (adapter, { page }) => {
  const query = new URLSearchParams({
    file: "sample.docx",
    marginGuides: "1",
    marginGuideColor: "rgb(12, 34, 56)",
  });
  await page.goto(`${adapter.baseUrl}/?${query.toString()}`);

  const guide = page.locator(".layout-page-margin-guide").first();
  const content = page.locator(".layout-page-content").first();
  await expect(guide).toBeVisible({ timeout: 30_000 });
  await expect(guide).toHaveCSS("border-color", "rgb(12, 34, 56)");
  await expect(guide).toHaveCSS("pointer-events", "none");

  const guideBox = await guide.boundingBox();
  const contentBox = await content.boundingBox();
  expect(guideBox).not.toBeNull();
  expect(contentBox).not.toBeNull();
  expect(guideBox?.x).toBeCloseTo(contentBox?.x ?? 0, 1);
  expect(guideBox?.y).toBeCloseTo(contentBox?.y ?? 0, 1);
  expect(guideBox?.width).toBeCloseTo(contentBox?.width ?? 0, 1);
  expect(guideBox?.height).toBeCloseTo(contentBox?.height ?? 0, 1);
});
