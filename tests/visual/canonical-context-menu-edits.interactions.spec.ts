import { expect, type Locator, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import type {} from "../parity/canonicalBridge";
import { evaluateCanonicalLoad } from "../parity/canonicalLoadEvaluation";

const adapters = [
  { name: "React", port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200 },
  {
    name: "Vue",
    port: Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201,
  },
] as const;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
/** A 1 x 1 PNG. */
const PNG =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

const picture = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="952500" cy="952500"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="dot.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

/** "Alpha Beta Gamma", then a paragraph holding one inline picture, then "TAIL". */
const source = async (): Promise<number[]> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="${R}/image" Target="media/dot.png"/></Relationships>`,
  );
  zip.set(
    "word/media/dot.png",
    Uint8Array.from(atob(PNG), (char) => char.charCodeAt(0)),
  );
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:pic="${PIC}"><w:body><w:p><w:r><w:t xml:space="preserve">Alpha Beta Gamma</w:t></w:r></w:p><w:p>${picture}</w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p></w:body></w:document>`,
  );
  return [...new Uint8Array(await zip.generate())];
};

/** Where "Beta" sits in the first paragraph. */
const BETA = { from: 7, to: 11 } as const;

const open = async (page: Page, port: number) => {
  await page.context().grantPermissions(["clipboard-read", "clipboard-write"], {
    origin: `http://localhost:${port}`,
  });
  await page.goto(`http://localhost:${port}/?session=canonical`);
  await page.waitForSelector(".layout-page");
  expect(await evaluateCanonicalLoad(page, await source())).toBe(true);
  await expect(page.locator(".layout-page").first()).toContainText("TAIL");
};

const savedBody = async (page: Page) => {
  const saved = await page.evaluate(
    async () => (await globalThis.__folioCanonical?.save()) ?? null,
  );
  expect(saved).not.toBeNull();
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  return (await zip.entry("word/document.xml")?.text()) ?? "";
};

/** Each saved paragraph's kept text, without deleted text. */
const savedParagraphs = (xml: string) =>
  [...xml.matchAll(/<w:p[ >].*?<\/w:p>/gsu)].map(([paragraph]) =>
    [...paragraph.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)]
      .map(([, text]) => text)
      .join(""),
  );

const select = async (page: Page, from: number, to: number) => {
  expect(
    await page.evaluate(
      ([start, end]) => globalThis.__folioCanonical?.select(start, end),
      [from, to] as const,
    ),
  ).toBe(true);
};

/** Open the menu over the first paragraph, keeping the selection. */
const openMenu = async (page: Page, target: Locator): Promise<Locator> => {
  await target.scrollIntoViewIfNeeded();
  const box = await target.boundingBox();
  if (!box) throw new Error("no painted target");
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2, {
    button: "right",
  });
  // React's menu is a labelled ARIA menu; Vue's is a list of buttons.
  const menu = page
    .locator('[role="menu"][aria-label="Text editing menu"], .ctx-menu')
    .first();
  await menu.waitFor({ state: "visible", timeout: 5000 });
  return menu;
};

const choose = async (menu: Locator, item: RegExp) => {
  await menu
    .locator('[role="menuitem"], button')
    .filter({ hasText: item })
    .first()
    .click();
};

const firstParagraph = (page: Page) =>
  page.locator(".layout-paragraph").filter({ hasText: "Alpha" }).first();

for (const adapter of adapters) {
  test.describe(`${adapter.name} canonical context menu edits`, () => {
    test("Delete removes the selection", async ({ page }) => {
      await open(page, adapter.port);
      await select(page, BETA.from - 1, BETA.to);
      await choose(
        await openMenu(page, firstParagraph(page)),
        /^\s*Delete\s*$/u,
      );
      await expect(firstParagraph(page)).not.toContainText("Beta");
      expect(savedParagraphs(await savedBody(page))[0]).toBe("Alpha Gamma");
    });

    test("Delete while suggesting keeps the text as a deletion", async ({
      page,
    }) => {
      await open(page, adapter.port);
      expect(
        await page.evaluate(() =>
          globalThis.__folioCanonical?.setMode("suggesting"),
        ),
      ).toBe(true);
      await select(page, BETA.from, BETA.to);
      await choose(
        await openMenu(page, firstParagraph(page)),
        /^\s*Delete\s*$/u,
      );
      await expect
        .poll(async () => /<w:del [^>]*>.*?Beta/su.test(await savedBody(page)))
        .toBe(true);
    });

    test("Cut removes the selection", async ({ page }) => {
      await open(page, adapter.port);
      await select(page, BETA.from - 1, BETA.to);
      await choose(await openMenu(page, firstParagraph(page)), /^\s*Cut\s*$/u);
      await expect(firstParagraph(page)).not.toContainText("Beta");
      expect(savedParagraphs(await savedBody(page))[0]).toBe("Alpha Gamma");
    });

    test("Paste as plain text replaces the selection", async ({ page }) => {
      await open(page, adapter.port);
      await page.evaluate(() => navigator.clipboard.writeText("Zeta"));
      await select(page, BETA.from, BETA.to);
      await choose(
        await openMenu(page, firstParagraph(page)),
        /Paste (as Plain Text|without formatting)/iu,
      );
      await expect(firstParagraph(page)).toContainText("Alpha Zeta Gamma");
      expect(savedParagraphs(await savedBody(page))[0]).toBe(
        "Alpha Zeta Gamma",
      );
    });
  });
}

// Only Vue's menu carries the picture actions.
const vue = adapters[1];
test.describe("Vue canonical picture menu edits", () => {
  for (const item of [/^\s*Delete image\s*$/u, /^\s*Cut\s*$/u]) {
    test(`${item.source} removes the selected picture`, async ({ page }) => {
      await open(page, vue.port);
      const image = page.locator(".layout-run-image, .layout-image").first();
      await image.waitFor({ state: "visible" });
      const box = await image.boundingBox();
      if (!box) throw new Error("no painted picture");
      await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
      await choose(await openMenu(page, image), item);
      await expect(
        page.locator(".layout-run-image, .layout-image"),
      ).toHaveCount(0);
      const xml = await savedBody(page);
      expect(xml).not.toContain("<w:drawing");
      expect(savedParagraphs(xml)).toContain("TAIL");
    });
  }
});
