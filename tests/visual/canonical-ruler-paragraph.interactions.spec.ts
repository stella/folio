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

/** One paragraph with a left tab stop at one inch, then "TAIL". */
const source = async (): Promise<number[]> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1440"/></w:tabs></w:pPr><w:r><w:t xml:space="preserve">Alpha Beta Gamma</w:t></w:r></w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p></w:body></w:document>`,
  );
  return [...new Uint8Array(await zip.generate())];
};

const open = async (page: Page, port: number) => {
  await page.goto(`http://localhost:${port}/?session=canonical`);
  await page.waitForSelector(".layout-page");
  expect(await evaluateCanonicalLoad(page, await source())).toBe(true);
  await expect(page.locator(".layout-page").first()).toContainText("TAIL");
  // The ruler shows the paragraph that holds the caret.
  expect(
    await page.evaluate(() => globalThis.__folioCanonical?.select(2, 2)),
  ).toBe(true);
};

/** The first saved paragraph's properties. */
const savedFirstParagraphProperties = async (page: Page) => {
  const saved = await page.evaluate(
    async () => (await globalThis.__folioCanonical?.save()) ?? null,
  );
  expect(saved).not.toBeNull();
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  const paragraph = /<w:p[ >].*?<\/w:p>/su.exec(xml)?.[0] ?? "";
  return /<w:pPr>.*?<\/w:pPr>/su.exec(paragraph)?.[0] ?? "";
};

/** A twips value of the first paragraph's `w:ind`, by either attribute name. */
const savedIndent = async (page: Page, side: "left" | "right") => {
  const properties = await savedFirstParagraphProperties(page);
  const names = side === "left" ? "(?:left|start)" : "(?:right|end)";
  const match = new RegExp(`<w:ind\\b[^>]*w:${names}="(\\d+)"`, "u").exec(
    properties,
  );
  return match ? Number(match[1]) : 0;
};

const ruler = (page: Page) => page.getByTestId("folio-horizontal-ruler");

/** Drag a marker sideways by `dx` pixels. */
const drag = async (page: Page, marker: Locator, dx: number) => {
  const box = await marker.boundingBox();
  if (!box) throw new Error("the ruler marker is not painted");
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  await page.mouse.move(x, y);
  await page.mouse.down();
  await page.mouse.move(x + dx / 2, y, { steps: 3 });
  await page.mouse.move(x + dx, y, { steps: 3 });
  await page.mouse.up();
};

for (const adapter of adapters) {
  test.describe(`${adapter.name} canonical ruler paragraph edits`, () => {
    test("dragging the left indent indents the paragraph", async ({ page }) => {
      await open(page, adapter.port);
      await drag(
        page,
        ruler(page).getByRole("slider", { name: "Left indent" }),
        48,
      );
      await expect.poll(() => savedIndent(page, "left")).toBeGreaterThan(0);
    });

    test("dragging the right indent indents the paragraph from the right", async ({
      page,
    }) => {
      await open(page, adapter.port);
      await drag(
        page,
        ruler(page).getByRole("slider", { name: "Right indent" }),
        -48,
      );
      await expect.poll(() => savedIndent(page, "right")).toBeGreaterThan(0);
    });

    test("double-clicking a tab stop removes it", async ({ page }) => {
      await open(page, adapter.port);
      expect(await savedFirstParagraphProperties(page)).toContain("<w:tabs>");
      await ruler(page).getByText("L", { exact: true }).first().dblclick();
      await expect
        .poll(() => savedFirstParagraphProperties(page))
        .not.toContain("<w:tabs>");
    });
  });
}
