import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HTML_PAYLOAD = "<html><body>Chunk markup</body></html>";
const TEXT_PAYLOAD = "CHUNK-READER-TEXT";
const CHUNK_HTML = '<w:altChunk r:id="rHtml"/>';
const CHUNK_TEXT = '<w:altChunk r:id="rText"/>';
const paragraph = (text: string): string =>
  `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const chunkDocument = async (): Promise<Buffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="htm" ContentType="text/html"/><Default Extension="txt" ContentType="text/plain"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rHtml" Type="${R}/aFChunk" Target="afchunk.htm"/><Relationship Id="rText" Type="${R}/aFChunk" Target="afchunk.txt"/></Relationships>`,
  );
  zip.set("word/afchunk.htm", HTML_PAYLOAD);
  zip.set("word/afchunk.txt", TEXT_PAYLOAD);
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraph("BEFORE-CHUNKS")}${CHUNK_HTML}${paragraph("BETWEEN-CHUNKS")}${CHUNK_TEXT}${paragraph("AFTER-CHUNKS")}<w:sectPr/></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

const load = async (page: Page, url: string): Promise<void> => {
  await page.goto(url);
  await page.locator("#file-input").setInputFiles({
    name: "chunks.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await chunkDocument(),
  });
  await expect(page.locator(".layout-page").first()).toContainText(
    "AFTER-CHUNKS",
  );
};

const topOfText = (page: Page, text: string): Promise<number> =>
  page
    .locator(".layout-page")
    .first()
    .evaluate((root, wanted) => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        if (walker.currentNode.textContent?.includes(wanted)) {
          const range = document.createRange();
          range.selectNodeContents(walker.currentNode);
          return range.getBoundingClientRect().top;
        }
      }
      throw new TypeError(`Text not on the page: ${wanted}`);
    }, text);

for (const { painter, query } of [
  { painter: "default", query: "" },
  { painter: "display-list", query: "?pageRenderer=display-list" },
]) {
  test(`an altChunk shows a placeholder in flow on the ${painter} painter`, async ({
    page,
  }) => {
    await load(page, `/${query}`);

    const before = await topOfText(page, "BEFORE-CHUNKS");
    const marker = await topOfText(page, "Content not displayed");
    const between = await topOfText(page, "BETWEEN-CHUNKS");
    const readerText = await topOfText(page, TEXT_PAYLOAD);
    const after = await topOfText(page, "AFTER-CHUNKS");
    expect(before).toBeLessThan(marker);
    expect(marker).toBeLessThan(between);
    expect(between).toBeLessThan(readerText);
    expect(readerText).toBeLessThan(after);
  });
}

test("typing at a placeholder leaves it and the chunk untouched", async ({
  page,
}) => {
  await load(page, "/");
  const pageEl = page.locator(".layout-page").first();
  const box = await pageEl.evaluate((root) => {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      if (walker.currentNode.textContent?.includes("Content not displayed")) {
        const range = document.createRange();
        range.selectNodeContents(walker.currentNode);
        const { left, top, height } = range.getBoundingClientRect();
        return { x: left + 10, y: top + height / 2 };
      }
    }
    throw new TypeError("No marker on the page");
  });
  await page.mouse.click(box.x, box.y);
  await page.keyboard.type("TYPED");

  await expect(pageEl).toContainText("Content not displayed");
  await expect(pageEl).not.toContainText(
    /Content not displayedTYPED|TYPEDContent/u,
  );

  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioPlayground?.getEditorRef()?.save();
    return buffer ? Array.from(new Uint8Array(buffer)) : null;
  });
  expect(saved).not.toBeNull();
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  const documentXml = await zip.entry("word/document.xml")?.text();
  expect(documentXml).toContain(CHUNK_HTML);
  expect(documentXml).toContain(CHUNK_TEXT);
  expect(documentXml).not.toContain("Content not displayed");
  expect(await zip.entry("word/afchunk.htm")?.text()).toBe(HTML_PAYLOAD);
  expect(await zip.entry("word/afchunk.txt")?.text()).toBe(TEXT_PAYLOAD);
});

test("the marker follows the editor's language", async ({ page }) => {
  await load(page, "/");
  await page
    .locator("select")
    .filter({ hasText: "English (en)" })
    .selectOption("cs");
  await expect(page.locator(".layout-page").first()).toContainText(
    "Obsah se nezobrazuje",
  );
});
