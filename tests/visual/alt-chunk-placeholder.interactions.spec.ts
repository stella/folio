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
const TEXT_PAYLOAD = "CHUNK-READER-TEXT\nCHUNK-SECOND-LINE";
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

/** A header whose chunk has far more lines than a page holds. */
const longHeaderChunkDocument = async (): Promise<Buffer> => {
  const lines = Array.from(
    { length: 400 },
    (_, index) => `HEADER-LINE-${String(index)}`,
  ).join("\n");
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="txt" ContentType="text/plain"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rHeader" Type="${R}/header" Target="header1.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/header1.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rText" Type="${R}/aFChunk" Target="afchunk.txt"/></Relationships>`,
  );
  zip.set("word/afchunk.txt", lines);
  zip.set(
    "word/header1.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="${W}" xmlns:r="${R}">${CHUNK_TEXT}<w:p/></w:hdr>`,
  );
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${paragraph("BODY-TEXT")}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
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
  { painter: "default", query: "?session=canonical" },
  {
    painter: "display-list",
    query: "?session=canonical&pageRenderer=display-list",
  },
]) {
  test(`an altChunk shows a placeholder in flow on the ${painter} painter`, async ({
    page,
  }) => {
    await load(page, `/${query}`);

    const before = await topOfText(page, "BEFORE-CHUNKS");
    const marker = await topOfText(page, "Content not displayed");
    const between = await topOfText(page, "BETWEEN-CHUNKS");
    const readerText = await topOfText(page, "CHUNK-READER-TEXT");
    const after = await topOfText(page, "AFTER-CHUNKS");
    expect(before).toBeLessThan(marker);
    expect(marker).toBeLessThan(between);
    expect(between).toBeLessThan(readerText);
    expect(readerText).toBeLessThan(after);
  });
}

/** Clicks well inside the text, then types: the click must not land in it. */
const typeAtText = async (
  page: Page,
  text: string,
  typed: string,
): Promise<void> => {
  const box = await page
    .locator(".layout-page")
    .first()
    .evaluate((root, wanted) => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) {
        if (walker.currentNode.textContent?.includes(wanted)) {
          const range = document.createRange();
          range.selectNodeContents(walker.currentNode);
          const { left, width, top, height } = range.getBoundingClientRect();
          return { x: left + width / 2, y: top + height / 2 };
        }
      }
      throw new TypeError(`Text not on the page: ${wanted}`);
    }, text);
  await page.mouse.click(box.x, box.y);
  await page.keyboard.type(typed);
};

for (const { painter, query } of [
  { painter: "default", query: "?session=canonical" },
  {
    painter: "display-list",
    query: "?session=canonical&pageRenderer=display-list",
  },
]) {
  test(`a click in a placeholder types at the next block on the ${painter} painter`, async ({
    page,
  }) => {
    await load(page, `/${query}`);
    const pageEl = page.locator(".layout-page").first();

    await typeAtText(page, "Content not displayed", "ONE");
    await expect(pageEl).toContainText("ONEBETWEEN-CHUNKS");
    await expect(pageEl).toContainText("Content not displayed");

    await typeAtText(page, "CHUNK-SECOND-LINE", "TWO");
    await expect(pageEl).toContainText("TWOAFTER-CHUNKS");
    await expect(pageEl).toContainText("CHUNK-SECOND-LINE");
  });
}

test("typing at a placeholder leaves the saved chunk untouched", async ({
  page,
}) => {
  await load(page, "/?session=canonical");
  await typeAtText(page, "Content not displayed", "TYPED");
  await expect(page.locator(".layout-page").first()).toContainText(
    "TYPEDBETWEEN-CHUNKS",
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

test("a long chunk in a header leaves the body its page", async ({ page }) => {
  await page.goto("/?session=canonical");
  await page.locator("#file-input").setInputFiles({
    name: "header-chunk.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await longHeaderChunkDocument(),
  });
  const first = page.locator(".layout-page").first();
  await expect(first).toContainText("BODY-TEXT");
  await expect(first).toContainText("Content not displayed");
  await expect(first).not.toContainText("HEADER-LINE-0");
});

test("the marker follows the editor's language", async ({ page }) => {
  await load(page, "/?session=canonical");
  await page
    .locator("select")
    .filter({ hasText: "English (en)" })
    .selectOption("cs");
  await expect(page.locator(".layout-page").first()).toContainText(
    "Obsah se nezobrazuje",
  );
});
