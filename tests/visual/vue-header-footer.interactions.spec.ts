import { expect, test } from "@playwright/test";

test("Vue header editing uses the shared persistent view and saves outside the body", async ({
  page,
}) => {
  await page.goto("http://localhost:4201/?file=sample.docx");
  await expect(page.locator(".layout-page-header").first()).toBeVisible({ timeout: 30_000 });

  const header = page.locator(".layout-page-header").first();
  await header.dblclick();
  await expect(page.locator(".hf-inline-editor")).toBeVisible();

  const hiddenHeader = page.locator(
    '.paged-editor__hidden-hf-pm [data-hf-kind="header"] .ProseMirror',
  );
  await expect(hiddenHeader).toHaveCount(1);
  await page.waitForTimeout(200);
  await page.keyboard.type("VueHeaderMark", { delay: 15 });

  await expect(hiddenHeader).toContainText("VueHeaderMark");
  await expect(header).toContainText("VueHeaderMark");
  await expect(page.locator('[data-testid="hf-caret"]')).toHaveCount(1);

  await page.getByRole("button", { name: /Options/u }).click();
  await page.getByRole("button", { name: "Close header editing" }).click();
  await expect(page.locator(".hf-inline-editor")).toHaveCount(0);
  await expect(hiddenHeader).toContainText("VueHeaderMark");

  const bodyText = await page.evaluate(() => window.__folioParity?.getDocumentText() ?? "");
  expect(bodyText).not.toContain("VueHeaderMark");
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.save() ?? Promise.resolve(0)))
    .toBeGreaterThan(0);
});
