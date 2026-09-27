/**
 * A review view paints the pages of the text it shows.
 *
 * `markup-views.docx` holds tracked changes in justified paragraphs; its two
 * companions hold the same text written without revisions (see
 * tests/visual/fixtures/build-markup-views.ts). Choosing a view in the markup
 * menu must paint exactly the lines, line breaks and paragraph boxes of the
 * companion that reads the same: Original the rejected text, No Markup and
 * Simple Markup the accepted text. Hiding runs inside the All Markup layout
 * leaves All Markup's line breaks, stretched justified lines, empty lines and
 * blank paragraphs, and fails every comparison here.
 */

import { test, type Page } from "@playwright/test";

import { expect, forEachAdapter, openEditor } from "../parity/parity-fixture";
import type { AdapterFixture } from "../parity/parity-fixture";

const TRACKED_FIXTURE = "markup-views.docx";

const VIEWS = [
  { label: "Original", companion: "markup-views-original.docx" },
  { label: "No Markup", companion: "markup-views-final.docx" },
  { label: "Simple Markup", companion: "markup-views-final.docx" },
] as const;

type PaintedParagraph = { top: number; height: number; lines: string[] };

/** Every painted body paragraph: its box and the text of each of its lines. */
const paintedBody = (page: Page): Promise<PaintedParagraph[]> =>
  page.evaluate(() =>
    Array.from(
      document.querySelectorAll<HTMLElement>(
        ".paged-editor__pages .layout-page-content .layout-paragraph",
      ),
      (paragraph) => ({
        top: Math.round(Number.parseFloat(paragraph.style.top) * 10) / 10,
        height: Math.round(Number.parseFloat(paragraph.style.height) * 10) / 10,
        lines: Array.from(
          paragraph.querySelectorAll<HTMLElement>(".layout-line"),
          (line) => line.textContent ?? "",
        ),
      }),
    ),
  );

const chooseView = async (page: Page, adapter: AdapterFixture, label: string): Promise<void> => {
  if (adapter.name === "react") {
    await page.getByRole("combobox").filter({ hasText: "All Markup" }).click();
    await page.getByRole("option", { name: label, exact: true }).click();
    return;
  }
  await page.locator(".review-controls__display").click();
  const option = page.locator(".review-controls__option", { hasText: label });
  // The toolbar scrolls horizontally when a menu option is brought into view.
  // Keep the same option node attached across that scroll and two paint frames.
  const scrollResult = await option.evaluate(async (element) => {
    const toolbar = element.closest(".basic-toolbar");
    if (!(toolbar instanceof HTMLElement)) return null;
    const before = toolbar.scrollLeft;
    toolbar.scrollLeft += before > 0 ? -1 : 1;
    const after = toolbar.scrollLeft;
    await new Promise<void>((resolve) =>
      requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
    );
    return { before, after, connected: element.isConnected };
  });
  expect(scrollResult).not.toBeNull();
  expect(scrollResult?.after).not.toBe(scrollResult?.before);
  expect(scrollResult).toMatchObject({ connected: true });
  await option.click();
};

for (const { label, companion } of VIEWS) {
  forEachAdapter(`${label} paints the pages of the text it shows`, async (adapter, { page }) => {
    await openEditor(page, adapter, companion);
    await expect.poll(() => paintedBody(page).then((body) => body.length)).toBeGreaterThan(0);
    const expected = await paintedBody(page);

    await openEditor(page, adapter, TRACKED_FIXTURE);
    await expect.poll(() => paintedBody(page)).not.toEqual(expected);
    await chooseView(page, adapter, label);

    await expect.poll(() => paintedBody(page)).toEqual(expected);
  });
}

test("document scroll closes the Vue markup menu", async ({ page }) => {
  const vuePort = Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201;
  await openEditor(page, { name: "vue", baseUrl: `http://localhost:${vuePort}` }, TRACKED_FIXTURE);

  await page.locator(".review-controls__display").click();
  await expect(page.locator(".review-controls__panel")).toBeVisible();

  await page.evaluate(() => document.dispatchEvent(new Event("scroll")));

  await expect(page.locator(".review-controls__panel")).toHaveCount(0);
});
