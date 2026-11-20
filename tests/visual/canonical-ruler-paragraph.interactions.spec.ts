import { expect, type Locator, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";
import { twipsToPixels } from "@stll/folio-core/utils/units";

import type {} from "../parity/canonicalBridge";
import { evaluateCanonicalLoad } from "../parity/canonicalLoadEvaluation";
import { PLAYGROUND_PORTS } from "../parity/playgroundHosts";
import type {} from "./interleavingBridge";

const adapters = [
  { name: "React", port: PLAYGROUND_PORTS.react },
  {
    name: "Vue",
    port: PLAYGROUND_PORTS.vue,
  },
] as const;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const LEFT_MARGIN_TWIPS = 1440;
const TAB_POSITION_TWIPS = 1440;

/** One paragraph with a left tab stop at one inch, then "TAIL". */
const source = async (): Promise<number[]> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body><w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="${TAB_POSITION_TWIPS}"/></w:tabs></w:pPr><w:r><w:t xml:space="preserve">Alpha Beta Gamma</w:t></w:r></w:p><w:p><w:r><w:t>TAIL</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:left="${LEFT_MARGIN_TWIPS}" w:right="1440" w:top="1440" w:bottom="1440"/></w:sectPr></w:body></w:document>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="${R}/header" Target="header1.xml"/></Relationships>`,
  );
  zip.set(
    "word/header1.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="${W}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>`,
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
    test("a header ruler edit does not change the body paragraph", async ({
      page,
    }) => {
      await open(page, adapter.port);
      const bodyProperties = await savedFirstParagraphProperties(page);
      await page.locator(".layout-page-header").first().dblclick();
      const editor = page.locator('[data-hf-r-id="rIdHeader"] .ProseMirror');
      await expect(editor).toBeFocused();
      await drag(
        page,
        ruler(page).getByLabel("Left indent", { exact: true }),
        48,
      );
      // Secondary canonical stories refuse unclassified formatting. In
      // particular, the body paragraph must not receive the header's edit.
      expect(await savedFirstParagraphProperties(page)).toBe(bodyProperties);
      await expect(editor).toHaveText("Header");
    });

    test("keyboard indentation keeps focus, saves both steps, and undoes each edit", async ({
      page,
    }) => {
      await open(page, adapter.port);
      const indent = ruler(page).getByRole("slider", {
        name: "Left indent",
        exact: true,
      });
      await expect(indent).toHaveAttribute("aria-valuenow", "0");
      await indent.focus();
      await page.keyboard.press("ArrowRight");
      await expect(indent).toBeFocused();
      await expect(indent).toHaveAttribute("aria-valuenow", "180");
      await expect.poll(() => savedIndent(page, "left")).toBe(180);
      await expect(indent).toBeFocused();
      // Use the actual focused element again; locator.press would refocus it
      // and hide a command-routing focus regression.
      await page.keyboard.press("ArrowRight");
      await expect(indent).toBeFocused();
      await expect(indent).toHaveAttribute("aria-valuenow", "360");
      await expect.poll(() => savedIndent(page, "left")).toBe(360);
      expect(
        await page.evaluate(() =>
          globalThis.__folioCanonical?.executeNamedCommand("undo", []),
        ),
      ).toBe(true);
      await expect.poll(() => savedIndent(page, "left")).toBe(180);
      expect(
        await page.evaluate(() =>
          globalThis.__folioCanonical?.executeNamedCommand("undo", []),
        ),
      ).toBe(true);
      await expect.poll(() => savedIndent(page, "left")).toBe(0);
    });

    test("dragging the left indent indents the paragraph", async ({ page }) => {
      await open(page, adapter.port);
      await drag(
        page,
        ruler(page).getByLabel("Left indent", { exact: true }),
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
        ruler(page).getByLabel("Right indent", { exact: true }),
        -48,
      );
      await expect.poll(() => savedIndent(page, "right")).toBeGreaterThan(0);
    });

    for (const zoom of [1, 1.5]) {
      test(`double-clicking a tab stop removes it at ${zoom * 100}% zoom`, async ({
        page,
      }) => {
        await open(page, adapter.port);
        await page.evaluate((value) => {
          const ref = globalThis.__folioPlayground?.getEditorRef();
          if (!ref) throw new Error("The canonical editor ref is unavailable.");
          ref.setZoom(value);
        }, zoom);
        expect(await savedFirstParagraphProperties(page)).toContain("<w:tabs>");
        const tab = ruler(page).getByText("L", { exact: true }).first();
        await expect
          .poll(async () => {
            const [rulerBox, tabBox, leftIndentBox] = await Promise.all([
              ruler(page)
                .getByRole("group", { name: "Horizontal ruler", exact: true })
                .boundingBox(),
              tab.boundingBox(),
              ruler(page)
                .getByLabel("Left indent", { exact: true })
                .boundingBox(),
            ]);
            if (!rulerBox || !tabBox || !leftIndentBox) return null;
            return {
              margin: Math.round(
                leftIndentBox.x + leftIndentBox.width / 2 - rulerBox.x,
              ),
              tab: Math.round(tabBox.x + tabBox.width / 2 - rulerBox.x),
            };
          })
          .toEqual({
            margin: twipsToPixels(LEFT_MARGIN_TWIPS) * zoom,
            tab: twipsToPixels(LEFT_MARGIN_TWIPS + TAB_POSITION_TWIPS) * zoom,
          });
        await tab.dblclick();
        await expect
          .poll(() => savedFirstParagraphProperties(page))
          .not.toContain("<w:tabs>");
      });
    }
  });
}
