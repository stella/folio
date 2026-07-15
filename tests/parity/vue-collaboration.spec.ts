import { test, expect } from "@playwright/test";

declare global {
  // eslint-disable-next-line typescript/consistent-type-definitions
  interface Window {
    __folioVueCollaboration?: {
      getSharedText: () => string;
      showRemoteSelection: () => boolean;
      wasSeeded: () => boolean;
    };
  }
}

test("Vue collaboration seeds and synchronizes the shared document", async ({ page }) => {
  await page.goto("http://localhost:4201/?collaboration=1");
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.hasView() ?? false), {
      timeout: 30_000,
    })
    .toBe(true);
  await expect
    .poll(() => page.evaluate(() => window.__folioVueCollaboration?.wasSeeded() ?? false))
    .toBe(true);

  expect(await page.evaluate(() => window.__folioParity?.insertText("shared text"))).toBe(true);
  await expect
    .poll(() => page.evaluate(() => window.__folioVueCollaboration?.getSharedText() ?? ""))
    .toContain("shared text");

  await expect
    .poll(() => page.evaluate(() => window.__folioVueCollaboration?.showRemoteSelection() ?? false))
    .toBe(true);
  await expect.poll(() => page.locator(".folio-remote-selection-rect").count()).toBeGreaterThan(0);
  await expect(page.locator(".folio-remote-selection-label")).toContainText("Remote collaborator");
});
