import { test } from "@playwright/test";

import { expect } from "./parity-fixture";

const VUE_PLAYGROUND = "http://localhost:4201";

test("Vue inherits locale and reactive system color mode from the host", async ({ page }) => {
  await page.emulateMedia({ colorScheme: "dark" });
  await page.goto(`${VUE_PLAYGROUND}/?file=sample.docx&locale=de&colorMode=system`);

  const editor = page.locator(".docx-editor-vue");
  await expect(editor).toHaveClass(/\bdark\b/u);
  await expect(page.getByRole("button", { name: "Datei", exact: true })).toBeVisible();

  const pageContent = page.locator(".layout-page-content").first();
  await expect(pageContent).toBeVisible();
  await pageContent.click({ button: "right", position: { x: 40, y: 40 } });
  await expect(page.locator(".ctx-menu")).toHaveClass(/\bdark\b/u);

  await page.emulateMedia({ colorScheme: "light" });
  await expect(editor).not.toHaveClass(/\bdark\b/u);
});
