import { expect, type Page } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  clearRefusals,
  expectNoRefusals,
  reactPort,
  vuePort,
} from "../parity/canonicalInteractionAssertions";
import { test } from "../parity/canonicalTest";

/**
 * A comment whose range opens between blocks, beside a content control or
 * between two paragraphs, is deleted through its card in both hosts on the
 * canonical session: the card goes, no marker of it is saved, and undo
 * brings the comment back with its range where it was.
 */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const COMMENTED =
  '<w:p><w:r><w:t xml:space="preserve">Please remit funds.</w:t></w:r><w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>';
const PLACEMENTS = {
  "beside a content control":
    '<w:commentRangeStart w:id="1" w:displacedByCustomXml="next"/>' +
    `<w:sdt><w:sdtPr><w:alias w:val="Payment"/><w:id w:val="7"/></w:sdtPr><w:sdtContent>${COMMENTED}</w:sdtContent></w:sdt>`,
  "between paragraphs": `<w:commentRangeStart w:id="1"/>${COMMENTED}`,
} as const;

const HOSTS = [
  { name: "React", port: reactPort, card: ".docx-comment-card" },
  { name: "Vue", port: vuePort, card: ".comment-card" },
] as const;

const documentWith = async (commented: string) => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R}/comments" Target="comments.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>Opening words</w:t></w:r></w:p>${commented}<w:p><w:r><w:t>Closing words</w:t></w:r></w:p></w:body></w:document>`,
  );
  zip.set(
    "word/comments.xml",
    `${HEAD}<w:comments xmlns:w="${W}"><w:comment w:id="1" w:author="Reviewer" w:date="2024-01-01T00:00:00Z"><w:p><w:r><w:t>Check the account</w:t></w:r></w:p></w:comment></w:comments>`,
  );
  return [...(await zip.generate())];
};

/** The saved main part's comment markers, in order. */
const savedMarkers = async (page: Page) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioCanonical?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  if (!saved) throw new TypeError("The canonical document did not save.");
  const xml =
    (await (
      await ZipArchive.read(Uint8Array.from(saved).buffer)
    )
      .entry("word/document.xml")
      ?.text()) ?? "";
  expect(xml).toContain("Please remit funds.");
  return (
    xml.match(/<w:comment(?:RangeStart|RangeEnd|Reference) [^>]*\/>/gu) ?? []
  );
};

test("a comment opening between blocks is deleted through its card and comes back on undo", async ({
  page,
}) => {
  test.setTimeout(120_000);
  for (const host of HOSTS)
    for (const [placement, commented] of Object.entries(PLACEMENTS)) {
      const label = `${host.name}, ${placement}`;
      await page.goto(`http://localhost:${host.port}/?session=canonical`);
      await page.waitForSelector(".layout-page");
      await clearRefusals(page);
      expect(
        await page.evaluate(
          (data) => globalThis.__folioCanonical?.loadAnySession(data),
          await documentWith(commented),
        ),
        label,
      ).toBe(true);
      await expect(page.locator(".layout-page").first()).toContainText(
        "Closing words",
      );
      const markers = await savedMarkers(page);
      expect(markers, label).toHaveLength(3);

      const card = page.locator(`${host.card}[data-comment-id="1"]`);
      await expect(card, label).toBeVisible();
      await card.click();
      await card.getByTitle("More options").click();
      await page
        .getByRole("button", { name: "Delete", exact: true })
        .filter({ visible: true })
        .first()
        .click();
      await expect(card, label).toHaveCount(0);
      await expectNoRefusals(page);
      expect(await savedMarkers(page), label).toEqual([]);

      await page.getByRole("button", { name: "Undo", exact: true }).click();
      await expect(card, label).toHaveCount(1);
      expect(await savedMarkers(page), label).toEqual(markers);
    }
});
