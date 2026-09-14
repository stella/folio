import type { Page } from "@playwright/test";

import { ensureLiveView, expect, forEachAdapter, openEditor } from "./parity-fixture";

const FIXTURE = "tracked-insertion-boundary.docx";
const ORIGINAL_DOCUMENT_TEXT = "The buyer promptly pays.";
const TRACKED_INSERTION_TEXT = " promptly";

const clickTrackedInsertionEnd = async (page: Page): Promise<void> => {
  const insertion = page.locator(".paged-editor__pages .docx-insertion").first();
  await expect(insertion).toHaveText(TRACKED_INSERTION_TEXT);

  const box = await insertion.boundingBox();
  if (!box) throw new Error("tracked insertion was not painted");

  await page.mouse.click(box.x + box.width - 0.2, box.y + box.height / 2);
};

const expectBoundaryResult = async (page: Page, insertedText: string): Promise<void> => {
  await expect
    .poll(() => page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""))
    .toBe(`The buyer promptly${insertedText} pays.`);
  await expect(page.locator(".paged-editor__pages .docx-insertion")).toHaveCount(1);
  await expect(page.locator(".paged-editor__pages .docx-insertion").first()).toHaveText(
    TRACKED_INSERTION_TEXT,
  );
};

forEachAdapter(
  "physical typing at a tracked insertion edge stays outside the revision",
  async (adapter, { page }) => {
    await openEditor(page, adapter, FIXTURE);
    await ensureLiveView(page);
    await expect
      .poll(() => page.evaluate(() => window.__folioParity?.getDocumentText() ?? ""))
      .toBe(ORIGINAL_DOCUMENT_TEXT);
    await clickTrackedInsertionEnd(page);

    await page.keyboard.type("LOCAL");

    await expectBoundaryResult(page, "LOCAL");
  },
);

forEachAdapter(
  "non-keyboard Unicode input at a tracked insertion edge stays outside the revision",
  async (adapter, { page }) => {
    await openEditor(page, adapter, FIXTURE);
    await ensureLiveView(page);
    await clickTrackedInsertionEnd(page);
    await page.evaluate(() => {
      document.documentElement.dataset["trackedBoundaryKeydowns"] = "0";
      document.addEventListener("keydown", () => {
        const current = Number(document.documentElement.dataset["trackedBoundaryKeydowns"]);
        document.documentElement.dataset["trackedBoundaryKeydowns"] = String(current + 1);
      });
    });

    await page.keyboard.insertText("日本語");

    await expectBoundaryResult(page, "日本語");
    await expect(page.locator("html")).toHaveAttribute("data-tracked-boundary-keydowns", "0");
  },
);

forEachAdapter(
  "IME composition at a tracked insertion edge commits once outside the revision",
  async (adapter, { page }) => {
    await openEditor(page, adapter, FIXTURE);
    await ensureLiveView(page);
    await clickTrackedInsertionEnd(page);

    const cdp = await page.context().newCDPSession(page);
    await cdp.send("Input.imeSetComposition", {
      text: "日",
      selectionStart: 1,
      selectionEnd: 1,
    });
    await cdp.send("Input.imeSetComposition", {
      text: "日本語",
      selectionStart: 3,
      selectionEnd: 3,
    });
    await cdp.send("Input.insertText", { text: "日本語" });

    await expectBoundaryResult(page, "日本語");
  },
);
