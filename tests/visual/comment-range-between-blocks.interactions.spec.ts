import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

/**
 * A comment whose range opens between blocks, beside a content control or
 * between two paragraphs, keeps one start and one end over the same text
 * through typing, with tracking on and off, a save and a reopen.
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
  return Buffer.from(await zip.generate());
};

const save = async (page: Page): Promise<number[]> => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioPlayground?.getEditorRef()?.save();
    return buffer ? [...new Uint8Array(buffer)] : null;
  });
  expect(saved).not.toBeNull();
  return saved ?? [];
};

const documentXml = async (bytes: readonly number[]) =>
  (await (
    await ZipArchive.read(Uint8Array.from(bytes).buffer)
  )
    .entry("word/document.xml")
    ?.text()) ?? "";

/** The comment's markers and the text its range covers, read from the saved part. */
const commentOf = (xml: string) => {
  const starts = xml.match(/<w:commentRangeStart w:id="1"[^>]*\/>/gu) ?? [];
  const ends = xml.match(/<w:commentRangeEnd w:id="1"[^>]*\/>/gu) ?? [];
  const from = xml.search(/<w:commentRangeStart w:id="1"/u);
  const to = xml.search(/<w:commentRangeEnd w:id="1"/u);
  const covered =
    from === -1 || to === -1
      ? ""
      : [
          ...xml
            .slice(from, to)
            .matchAll(/<w:(?:t|delText)(?: [^>]*)?>([^<]*)</gu),
        ]
          .map(([, text]) => text)
          .join("");
  return { starts: starts.length, ends: ends.length, covered };
};

const word = (page: Page, text: string) =>
  page.locator(".layout-page").getByText(text).first();

for (const [placement, commented] of Object.entries(PLACEMENTS)) {
  for (const tracking of [false, true]) {
    test(`a comment opening ${placement} keeps its range, tracking ${tracking ? "on" : "off"}`, async ({
      page,
    }) => {
      await page.goto("/");
      await page.locator("#file-input").setInputFiles({
        name: "commented.docx",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        buffer: await documentWith(commented),
      });
      await expect(page.locator(".layout-page").first()).toContainText(
        "Closing words",
      );
      if (tracking) {
        await page
          .getByRole("button", { name: "Track Changes", exact: true })
          .click();
      }
      // Inside the commented text, then beside it in the paragraph before.
      await word(page, "Please remit funds.").click();
      await page.keyboard.press("Home");
      for (let step = 0; step < "Please".length; step += 1)
        await page.keyboard.press("ArrowRight");
      await page.keyboard.type(" now");
      await word(page, "Opening words").click();
      await page.keyboard.press("End");
      await page.keyboard.type(" here");
      await page.waitForTimeout(300);

      const first = await save(page);
      const saved = commentOf(await documentXml(first));
      expect(saved.starts, "one start").toBe(1);
      expect(saved.ends, "one end").toBe(1);
      expect(saved.covered).toContain("Please");
      expect(saved.covered).toContain(" now");
      expect(saved.covered).toContain(" remit funds.");
      expect(saved.covered).not.toContain("Opening");
      expect(saved.covered).not.toContain("here");

      // Reopened, it reads and saves the same range.
      await page.evaluate(async (bytes) => {
        await globalThis.__folioPlayground
          ?.getEditorRef()
          ?.loadDocumentBuffer(new Uint8Array(bytes));
      }, first);
      await expect(page.locator(".layout-page").first()).toContainText(
        "Closing words",
      );
      const reopened = commentOf(await documentXml(await save(page)));
      expect(reopened).toEqual(saved);
    });
  }
}

for (const [placement, commented] of Object.entries(PLACEMENTS)) {
  for (const session of ["legacy", "canonical"] as const)
    test(`${session}: a comment opening ${placement} deleted through its card leaves no marker`, async ({
      page,
    }) => {
      await page.goto(`/?session=${session}`);
      await page.locator("#file-input").setInputFiles({
        name: "commented.docx",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        buffer: await documentWith(commented),
      });
      await expect(page.locator(".layout-page").first()).toContainText(
        "Closing words",
      );
      const card = page.locator('.docx-comment-card[data-comment-id="1"]');
      await expect(card).toBeVisible();
      await card.click();
      await card.getByTitle("More options").click();
      await page
        .getByRole("menu")
        .getByRole("button", { name: "Delete" })
        .click();
      await expect(card).toHaveCount(0);
      await page.waitForTimeout(300);

      const xml = await documentXml(await save(page));
      expect(
        xml.match(/<w:comment(?:RangeStart|RangeEnd|Reference) [^>]*\/>/gu) ??
          [],
      ).toEqual([]);
      expect(xml).toContain("Please remit funds.");
    });
}
