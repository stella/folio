import { expect, test } from "@playwright/test";
import { ensureLiveView } from "../parity/parity-fixture";

for (const query of ["", "?paragraphs=12", "?file=sample.docx&paragraphs=12"]) {
  test(`playground initializes its document and keeps the scroll bridge through rerenders (${query || "default"})`, async ({
    page,
  }) => {
    await page.goto(`/${query}`);
    await expect
      .poll(() => page.evaluate(() => window.__folioParity?.getTotalPages() ?? 0))
      .toBeGreaterThan(0);
    await ensureLiveView(page);
    const initialText = await page.evaluate(() => window.__folioParity?.getDocumentText());
    if (query === "?paragraphs=12") {
      expect(initialText?.length).toBeGreaterThan(0);
      expect(initialText?.match(/Performance paragraph /gu)).toHaveLength(12);
    } else {
      if (query.includes("file=")) expect(initialText?.length).toBeGreaterThan(0);
      expect(initialText).not.toContain("Performance paragraph ");
    }
    expect(await page.evaluate(() => !!globalThis.__folioScrollParity)).toBe(true);
    await page.evaluate(() => {
      globalThis.__startupScrollBridge = globalThis.__folioScrollParity;
    });
    await page.getByRole("button", { name: "Track Changes", exact: true }).click();
    expect(
      await page.evaluate(
        () => globalThis.__startupScrollBridge === globalThis.__folioScrollParity,
      ),
    ).toBe(true);
    await page.getByRole("button", { name: "New", exact: true }).click();
    await expect
      .poll(() => page.evaluate(() => window.__folioParity?.getTotalPages() ?? 0))
      .toBeGreaterThan(0);
    await ensureLiveView(page);
    await expect.poll(() => page.evaluate(() => window.__folioParity?.getDocumentText())).toBe("");
  });
}

test("collaboration startup mounts its editor branch", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/?collaboration=1");
  await expect(page.locator(".pg-collab-header")).toContainText("folio collaboration demo");
  await expect(page.getByTestId("folio-editor")).toBeVisible();
  await expect.poll(() => page.locator(".layout-page").count()).toBeGreaterThan(0);
  await expect(page.locator(".pg-collab-loading")).toHaveCount(0);
  await expect(page.getByTestId("playground-controls")).toHaveCount(0);
  await expect(page.getByRole("button", { name: "Share link", exact: true })).toBeVisible();
  expect(errors).toEqual([]);
});

declare global {
  var __startupScrollBridge: typeof globalThis.__folioScrollParity;
}
