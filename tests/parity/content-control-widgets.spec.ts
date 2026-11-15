import { readFileSync } from "node:fs";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  ensureLiveView,
  expect,
  forEachAdapter,
  openCanonicalEditor,
} from "./parity-fixture";

const paragraph = (paraId: string, text: string) =>
  `<w:p w14:paraId="${paraId}" w14:textId="77777777"><w:r><w:t>${text}</w:t></w:r></w:p>`;

/** The sample package with its body replaced by a dropdown and a date control. */
const contentControlsDocument = async (): Promise<number[]> => {
  const zip = await ZipArchive.read(
    readFileSync(new URL("../visual/fixtures/sample.docx", import.meta.url)),
  );
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  const bodyStart = xml.indexOf("<w:body>") + "<w:body>".length;
  const sectionStart = xml.lastIndexOf("<w:sectPr");
  if (bodyStart < "<w:body>".length || sectionStart < bodyStart)
    throw new TypeError("The sample body is missing");
  const controls = [
    '<w:sdt><w:sdtPr><w:tag w:val="state"/><w:id w:val="101"/><w:dropDownList><w:listItem w:displayText="California" w:value="ca"/><w:listItem w:displayText="New York" w:value="ny"/></w:dropDownList></w:sdtPr>',
    `<w:sdtContent>${paragraph("3CC00001", "California")}</w:sdtContent></w:sdt>`,
    '<w:sdt><w:sdtPr><w:tag w:val="effective"/><w:id w:val="102"/><w:date><w:dateFormat w:val="yyyy-MM-dd"/></w:date></w:sdtPr>',
    `<w:sdtContent>${paragraph("3CC00002", "2026-01-15")}</w:sdtContent></w:sdt>`,
    paragraph("3CC00003", "After the controls"),
  ].join("");
  zip.set(
    "word/document.xml",
    `${xml.slice(0, bodyStart)}${controls}${xml.slice(sectionStart)}`,
  );
  return [...(await zip.generate())];
};

forEachAdapter(
  "edits dropdown and date content controls",
  async (adapter, { page }) => {
    await openCanonicalEditor(page, adapter);
    const bytes = await contentControlsDocument();
    await page.evaluate((data) => window.__folioParity?.loadDocx(data), bytes);
    await ensureLiveView(page);

    const dropdown = page
      .locator('.paged-editor__pages [data-sdt-type="dropdown"]')
      .first();
    await expect(dropdown).toBeVisible();
    await dropdown.click();
    await page.getByRole("menuitem", { name: "New York" }).click();
    await expect(dropdown).toContainText("New York");

    const date = page
      .locator('.paged-editor__pages [data-sdt-type="date"]')
      .first();
    await expect(date).toBeVisible();
    await date.click();
    await page.locator('input[type="date"]').fill("2026-06-02");
    await expect(date).toContainText("2026-06-02");

    expect(
      await page.evaluate(() => window.__folioParity?.save() ?? 0),
    ).toBeGreaterThan(0);
  },
);
