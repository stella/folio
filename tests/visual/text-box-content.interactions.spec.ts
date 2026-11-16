import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS = `xmlns:w="${W}" xmlns:r="${R}" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"`;
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const box = (inner: string) =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="4000000" cy="2000000"/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="2000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const documentWith = async (inner: string) => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="html" ContentType="text/html"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdC" Type="${R}/aFChunk" Target="chunk.html"/></Relationships>`,
  );
  zip.set("word/chunk.html", "<html><body>Chunk</body></html>");
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document ${NS}><w:body><w:p><w:r><w:t>Host</w:t></w:r>${box(inner)}</w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

const load = async (page: Page, inner: string) => {
  await page.goto("/?session=legacy");
  await page.locator("#file-input").setInputFiles({
    name: "box.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await documentWith(inner),
  });
  await expect(page.locator(".layout-page").first()).toContainText("TAIL");
};

const savedDocument = async (page: Page) => {
  const saved = await page.evaluate(async () => {
    const buffer = await (
      globalThis as unknown as {
        __folioPlayground?: {
          getEditorRef: () => { save: () => Promise<ArrayBuffer> } | null;
        };
      }
    ).__folioPlayground
      ?.getEditorRef()
      ?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  expect(saved).not.toBeNull();
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  return (await zip.entry("word/document.xml")?.text()) ?? "";
};

const typeInHost = async (page: Page) => {
  await page.getByText("Host", { exact: false }).first().click();
  await page.keyboard.press("Home");
  await page.keyboard.type("Z");
  await page.waitForTimeout(300);
};

const CONTROL = `<w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sdt><w:sdtPr><w:id w:val="9"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>Held</w:t></w:r></w:p></w:sdtContent></w:sdt>`;
const CHUNK = `<w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:altChunk r:id="rIdC"/>`;

test("text in a content control in a text box shows, and an edit of its host keeps it", async ({
  page,
}) => {
  await load(page, CONTROL);
  await expect(page.locator(".layout-page").first()).toContainText("Held");
  await typeInHost(page);
  const xml = await savedDocument(page);
  expect(xml).toContain("Held");
  expect(xml).toContain("<w:sdt>");
  expect(xml).toMatch(/<w:t>[^<]*Z[^<]*<\/w:t>/u);
});

test("an altChunk in a text box shows its marker, and an edit of its host keeps it", async ({
  page,
}) => {
  await load(page, CHUNK);
  await expect(page.locator(".layout-page").first()).toContainText(
    "Content not displayed",
  );
  await typeInHost(page);
  const xml = await savedDocument(page);
  expect(xml).toContain("<w:altChunk");
  expect(xml).toMatch(/<w:t>[^<]*Z[^<]*<\/w:t>/u);
});

test("text typed in a content control in a text box saves inside the control", async ({
  page,
}) => {
  await load(page, CONTROL);
  await page.locator(".layout-page").getByText("Held").first().click();
  await page.keyboard.type("Q");
  await page.waitForTimeout(300);
  await expect(page.locator(".layout-page").first()).toContainText("Q");
  const xml = await savedDocument(page);
  const control = /<w:sdt>.*<\/w:sdt>/su.exec(xml)?.[0] ?? "";
  expect(control).toMatch(/<w:t>[^<]*Q[^<]*<\/w:t>/u);
  expect(xml).toContain("Plain");
});

test("text wrapped in a control in a text box survives an edit with Track Changes on", async ({
  page,
}) => {
  await load(page, CONTROL);
  await page
    .getByRole("button", { name: "Track Changes", exact: true })
    .click();
  await typeInHost(page);
  const xml = await savedDocument(page);
  expect(xml).toContain("Held");
  expect(xml).toContain("<w:sdt>");
  expect(xml).toMatch(/<w:ins [^>]*>.*Z.*<\/w:ins>/su);
});
