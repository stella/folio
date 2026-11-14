import { ZipArchive } from "@stll/docx-core/zip";

/**
 * A citation field nested in another field's result keeps its text when the
 * editor saves after a structural edit elsewhere in the document.
 */
import { expect, type Page, test } from "../support/unclassifiedProbe";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const run = (inner: string) => `<w:r>${inner}</w:r>`;
const text = (value: string) => run(`<w:t xml:space="preserve">${value}</w:t>`);
const fieldChar = (type: string) => run(`<w:fldChar w:fldCharType="${type}"/>`);
const instruction = (value: string) =>
  run(`<w:instrText xml:space="preserve">${value}</w:instrText>`);

const fixture = async (): Promise<Buffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body>` +
      `<w:p>${text("OPENING")}</w:p>` +
      `<w:p>${text("Cited ")}${fieldChar("begin")}${instruction(" ADDIN EN.CITE outer ")}${fieldChar("separate")}` +
      `${text("(Outer, 2003")}${fieldChar("begin")}${instruction(" ADDIN EN.CITE inner ")}${fieldChar("separate")}` +
      `${text("; Inner, 1966)")}${fieldChar("end")}${fieldChar("end")}${text(". After.")}</w:p>` +
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

test("a nested citation survives a save after a new paragraph", async ({
  page,
}) => {
  await page.goto("/");
  await page.locator("#file-input").setInputFiles({
    name: "nested-field.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await fixture(),
  });
  const opening = page.locator(".layout-page").getByText("OPENING").first();
  await expect(opening).toBeVisible();
  await expect(page.locator(".layout-page").first()).toContainText(
    "; Inner, 1966)",
  );
  await opening.click();
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("Added.");

  const xml = await savedDocument(page);
  const cited = xml.slice(xml.indexOf("Cited"), xml.indexOf(". After."));
  expect(
    [...cited.matchAll(/w:fldCharType="(\w+)"/gu)].map(([, type]) => type),
  ).toEqual(["begin", "separate", "begin", "separate", "end", "end"]);
  expect(cited).toContain("; Inner, 1966)");
  expect(xml).toContain("Added.");
});
