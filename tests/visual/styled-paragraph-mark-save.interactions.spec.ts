/**
 * Typing in a paragraph with a style, whose mark repeats the runs' own
 * properties, leaves those properties on the runs in the saved document.
 */
import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const RAISED = '<w:position w:val="-2"/><w:sz w:val="28"/>';

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
    `${HEAD}<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style><w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/></w:style></w:styles>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body>` +
      `<w:p><w:pPr><w:pStyle w:val="Body"/><w:rPr>${RAISED}</w:rPr></w:pPr><w:r><w:rPr>${RAISED}</w:rPr><w:t>RAISEDTEXT</w:t></w:r></w:p>` +
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

test("a run keeps the properties its paragraph mark repeats after typing beside it", async ({
  page,
}) => {
  await page.goto("/?session=canonical");
  await page.locator("#file-input").setInputFiles({
    name: "styled-mark.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await fixture(),
  });
  const heading = page.locator(".layout-page").getByText("RAISEDTEXT").first();
  await expect(heading).toBeVisible();
  await heading.click();
  await page.keyboard.press("End");
  await page.keyboard.type("X");

  const xml = await savedDocument(page);
  expect(xml).toContain("RAISEDTEXTX");
  const run =
    /<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t[^>]*>RAISEDTEXTX<\/w:t><\/w:r>/u.exec(
      xml,
    );
  expect(run?.[1] ?? "").toContain('<w:position w:val="-2"/>');
  expect(run?.[1] ?? "").toContain('<w:sz w:val="28"/>');
});
