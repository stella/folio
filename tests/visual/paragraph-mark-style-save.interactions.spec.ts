import { ZipArchive } from "@stll/docx-core/zip";

/**
 * Typing in a paragraph whose mark and runs share a bold character style
 * leaves the runs bold in the saved document.
 */
import { expect, type Page, test } from "../support/unclassifiedProbe";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const STRONG = '<w:rStyle w:val="Strong"/>';

const fixture = async (): Promise<Buffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="${R}/styles" Target="styles.xml"/></Relationships>`,
  );
  zip.set(
    "word/styles.xml",
    `${HEAD}<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style></w:styles>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body>` +
      `<w:p><w:pPr><w:rPr>${STRONG}</w:rPr></w:pPr><w:r><w:rPr>${STRONG}</w:rPr><w:t>STRONGHEAD</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Body.</w:t></w:r></w:p>` +
      `</w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

const savedDocument = async (page: Page) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioPlayground?.getEditorRef()?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  expect(saved).not.toBeNull();
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  return (await zip.entry("word/document.xml")?.text()) ?? "";
};

test("a run keeps its character style's bold after typing beside it", async ({
  page,
}) => {
  await page.goto("/");
  await page.locator("#file-input").setInputFiles({
    name: "strong-heading.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await fixture(),
  });
  const heading = page.locator(".layout-page").getByText("STRONGHEAD").first();
  await expect(heading).toBeVisible();
  await heading.click();
  await page.keyboard.press("End");
  await page.keyboard.type("X");

  const xml = await savedDocument(page);
  expect(xml).toContain("STRONGHEADX");
  expect(xml).not.toMatch(/<w:b w:val="(?:0|false)"\/>/u);
});
