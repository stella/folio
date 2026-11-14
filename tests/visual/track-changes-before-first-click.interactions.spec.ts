import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const documentWith = async (first: string) => {
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
    `${HEAD}<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t xml:space="preserve">${first}</w:t></w:r></w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

const open = async (page: Page, first = "Alpha beta gamma") => {
  await page.locator("#file-input").setInputFiles({
    name: "plain.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await documentWith(first),
  });
  await expect(page.locator(".layout-page").first()).toContainText("TAIL");
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

const word = (page: Page, text: string) =>
  page.locator(".layout-page").getByText(text).first();

/** How a host or a reader turns tracking on, none of them touching the text. */
const ROUTES = [
  {
    route: "the mode prop",
    turnOn: (page: Page) =>
      page.getByRole("button", { name: "Track Changes", exact: true }).click(),
  },
  {
    route: "the toolbar",
    turnOn: (page: Page) =>
      page.getByRole("button", { name: "Toggle Track Changes" }).click(),
  },
] as const;

const expectTrackingOn = (page: Page) =>
  expect(
    page.getByRole("button", { name: "Toggle Track Changes" }),
  ).toHaveAttribute("aria-pressed", "true");

for (const { route, turnOn } of ROUTES) {
  test.describe(`tracking turned on through ${route} before the first click`, () => {
    test.beforeEach(async ({ page }) => {
      await page.goto("/?session=canonical");
      await open(page);
      await turnOn(page);
      await expectTrackingOn(page);
    });

    test("typed text is a tracked insertion", async ({ page }) => {
      await word(page, "Alpha beta gamma").click();
      await page.keyboard.type("Z");
      await page.waitForTimeout(300);
      expect(await savedDocument(page)).toMatch(
        /<w:ins [^>]*>(?:(?!<\/w:ins>).)*<w:t[^>]*>Z<\/w:t>/su,
      );
    });

    test("deleted text is a tracked deletion", async ({ page }) => {
      await word(page, "Alpha beta gamma").dblclick();
      await page.keyboard.press("Delete");
      await page.waitForTimeout(300);
      const xml = await savedDocument(page);
      expect(xml).toMatch(/<w:del [^>]*>.*<w:delText/su);
      // What is deleted stays in the file until it is accepted.
      expect(`${xml.replaceAll(/<[^>]+>/gu, "")}`).toContain("Alpha");
      expect(`${xml.replaceAll(/<[^>]+>/gu, "")}`).toContain("gamma");
    });

    test("pasted text is a tracked insertion", async ({ page, context }) => {
      await context.grantPermissions(["clipboard-read", "clipboard-write"]);
      await page.evaluate(() => navigator.clipboard.writeText("Pasted words"));
      await word(page, "Alpha beta gamma").click();
      await page.keyboard.press("ControlOrMeta+v");
      await page.waitForTimeout(500);
      expect(await savedDocument(page)).toMatch(
        /<w:ins [^>]*>(?:(?!<\/w:ins>).)*Pasted words/su,
      );
    });
  });
}

test("an editor mounted with tracking on tracks the first edit", async ({
  page,
}) => {
  await page.goto("/?session=canonical&mode=suggesting");
  await open(page);
  await expectTrackingOn(page);
  await word(page, "Alpha beta gamma").click();
  await page.keyboard.type("Z");
  await page.waitForTimeout(300);
  expect(await savedDocument(page)).toMatch(
    /<w:ins [^>]*>(?:(?!<\/w:ins>).)*<w:t[^>]*>Z<\/w:t>/su,
  );
});

test("tracking stays on for a document opened while it is on", async ({
  page,
}) => {
  await page.goto("/?session=canonical");
  await open(page);
  await page
    .getByRole("button", { name: "Track Changes", exact: true })
    .click();
  await expectTrackingOn(page);
  await open(page, "Second document here");
  await expect(page.locator(".layout-page").first()).toContainText("Second");
  await word(page, "Second document here").click();
  await page.keyboard.type("Z");
  await page.waitForTimeout(300);
  expect(await savedDocument(page)).toMatch(
    /<w:ins [^>]*>(?:(?!<\/w:ins>).)*<w:t[^>]*>Z<\/w:t>/su,
  );
});

test("tracking turned off before the first click leaves edits direct", async ({
  page,
}) => {
  await page.goto("/?session=canonical");
  await open(page);
  const toggle = page.getByRole("button", { name: "Toggle Track Changes" });
  await toggle.click();
  await expectTrackingOn(page);
  await toggle.click();
  await word(page, "Alpha beta gamma").click();
  await page.keyboard.type("Z");
  await page.waitForTimeout(300);
  expect(await savedDocument(page)).not.toContain("<w:ins ");
});
