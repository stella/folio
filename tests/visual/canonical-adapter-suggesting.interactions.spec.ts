import { expect, test, type Page } from "@playwright/test";
import { createDocx } from "../../packages/core/src/docx/rezip";
import { parseDocx } from "../../packages/core/src/docx/parser";
import { createEmptyDocument } from "../../packages/core/src/utils/createDocument";
import { createEmptyHeaderFooter } from "../../packages/core/src/utils/headerFooter";
import type { Document } from "../../packages/core/src/types/document";
import type {} from "../parity/canonicalBridge";

const MODIFIER = process.platform === "darwin" ? "Meta" : "Control";
const adapters = [
  { name: "React", port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200 },
  { name: "Vue", port: Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201 },
] as const;

type SetHeaderFooterTextOptions = {
  document: Document;
  position: "header" | "footer";
  text: string;
};

const setHeaderFooterText = ({ document, position, text }: SetHeaderFooterTextOptions): string => {
  const parts = position === "header" ? document.package.headers : document.package.footers;
  const identity = [...(parts?.keys() ?? [])].at(0);
  if (!identity || !parts) throw new TypeError(`Missing ${position} fixture.`);
  const part = parts.get(identity);
  if (!part) throw new TypeError(`Missing ${position} part.`);
  parts.set(identity, {
    ...part,
    content: [{ type: "paragraph", content: [{ type: "run", content: [{ type: "text", text }] }] }],
  });
  return identity;
};

const canonicalSnapshot = async (page: Page) => {
  await page.waitForFunction(() => globalThis.__folioCanonical?.canSnapshot());
  const current = await page.evaluate(() => globalThis.__folioCanonical?.snapshot());
  expect(current?.active).toBe(true);
  expect(current?.projectionMatchesCanonical).toBe(true);
  if (!current?.document) throw new TypeError("Canonical document unavailable.");
  return { ...current, document: current.document };
};

const hasTrackedInsertion = (value: unknown, author: string): boolean => {
  if (Array.isArray(value)) return value.some((entry) => hasTrackedInsertion(entry, author));
  if (value === null || typeof value !== "object") return false;
  if (
    "type" in value &&
    value.type === "insertion" &&
    "info" in value &&
    value.info !== null &&
    typeof value.info === "object" &&
    "author" in value.info &&
    value.info.author === author
  ) {
    return true;
  }
  return Object.values(value).some((entry) => hasTrackedInsertion(entry, author));
};

for (const adapter of adapters) {
  test(`canonical track changes control edits body, header and footer in ${adapter.name}`, async ({
    page,
  }) => {
    test.setTimeout(90_000);
    const sourceDocument = createEmptyDocument({ initialText: "Body seed" });
    const withHeader = createEmptyHeaderFooter(sourceDocument, "header", false);
    if (!withHeader) throw new TypeError("Could not create header fixture.");
    const document = createEmptyHeaderFooter(withHeader, "footer", false);
    if (!document) throw new TypeError("Could not create footer fixture.");
    const headerRId = setHeaderFooterText({
      document,
      position: "header",
      text: "Header seed",
    });
    const footerRId = setHeaderFooterText({
      document,
      position: "footer",
      text: "Footer seed",
    });
    const bytes = [...new Uint8Array(await createDocx(document))];

    await page.goto(`http://localhost:${adapter.port}/?session=canonical`);
    await page.waitForSelector(".layout-page");
    expect(await page.evaluate((input) => globalThis.__folioCanonical?.load(input), bytes)).toBe(
      true,
    );
    const initial = await canonicalSnapshot(page);

    const trackChanges = page.getByRole("button", {
      name: "Toggle Track Changes",
      exact: true,
    });
    await expect(trackChanges).toBeVisible();
    await trackChanges.click();
    await expect(trackChanges).toHaveAttribute("aria-pressed", "true");

    const body = page.locator(".layout-paragraph").first();
    await body.click();
    await page.keyboard.press("End");
    await page.keyboard.type(" body tracked");
    const bodyAfterTyping = await canonicalSnapshot(page);
    const bodyModel = bodyAfterTyping.document.package.document.content;
    expect(hasTrackedInsertion(bodyModel, "Folio User")).toBe(true);
    await page.keyboard.press(`${MODIFIER}+z`);
    expect((await canonicalSnapshot(page)).document.package.document.content).toEqual(
      initial.document.package.document.content,
    );
    await page.keyboard.press(`${MODIFIER}+Shift+z`);
    expect((await canonicalSnapshot(page)).document.package.document.content).toEqual(bodyModel);

    const originalBodyJSON = JSON.stringify(bodyModel);
    const header = page.locator(".layout-page-header").first();
    await header.dblclick();
    await expect(page.locator(".hf-inline-editor")).toBeVisible();
    const headerEditor = page.locator(
      `.paged-editor__hidden-hf-pm [data-hf-r-id="${headerRId}"] .ProseMirror`,
    );
    await expect(headerEditor).toBeAttached();
    await page.keyboard.press("End");
    await page.keyboard.type(" header tracked");
    await expect(header).toContainText("header tracked");
    await page.keyboard.press(`${MODIFIER}+z`);
    await expect(header).not.toContainText("header tracked");
    await page.keyboard.press(`${MODIFIER}+Shift+z`);
    await expect(header).toContainText("header tracked");
    expect(JSON.stringify((await canonicalSnapshot(page)).document.package.document.content)).toBe(
      originalBodyJSON,
    );
    await page.getByRole("button", { name: "Options", exact: false }).click();
    await page.getByRole("button", { name: "Close header editing" }).click();

    const footer = page.locator(".layout-page-footer").first();
    await footer.scrollIntoViewIfNeeded();
    await footer.dblclick();
    await expect(page.locator(".hf-inline-editor")).toBeVisible();
    const footerEditor = page.locator(
      `.paged-editor__hidden-hf-pm [data-hf-r-id="${footerRId}"] .ProseMirror`,
    );
    await expect(footerEditor).toBeAttached();
    await page.keyboard.press("End");
    await page.keyboard.type(" footer tracked");
    await expect(footer).toContainText("footer tracked");
    await page.keyboard.press(`${MODIFIER}+z`);
    await expect(footer).not.toContainText("footer tracked");
    await page.keyboard.press(`${MODIFIER}+Shift+z`);
    await expect(footer).toContainText("footer tracked");
    const beforeSave = await canonicalSnapshot(page);
    expect(JSON.stringify(beforeSave.document.package.document.content)).toBe(originalBodyJSON);

    const saved = await page.evaluate(() => globalThis.__folioCanonical?.save());
    if (!saved) throw new TypeError("Canonical save unavailable.");
    const reopened = await parseDocx(new Uint8Array(saved).buffer, {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(hasTrackedInsertion(reopened.package.document.content, "Folio User")).toBe(true);
    const savedHeader = reopened.package.headers?.get(headerRId);
    const savedFooter = reopened.package.footers?.get(footerRId);
    expect(savedHeader?.content).toEqual(
      expect.arrayContaining([expect.objectContaining({ type: "paragraph" })]),
    );
    expect(hasTrackedInsertion(savedHeader?.content, "Folio User")).toBe(true);
    expect(hasTrackedInsertion(savedFooter?.content, "Folio User")).toBe(true);
    expect(JSON.stringify(reopened.package.document.content)).toBe(originalBodyJSON);
  });
}
