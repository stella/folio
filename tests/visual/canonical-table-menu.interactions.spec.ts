import { expect, type Locator, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";
import type {} from "../parity/canonicalBridge";
import { evaluateCanonicalLoad } from "../parity/canonicalLoadEvaluation";
import type {} from "../parity/parity-fixture";

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const adapters = [
  { name: "React", port: Number(process.env["FOLIO_PLAYGROUND_PORT"]) || 4200 },
  {
    name: "Vue",
    port: Number(process.env["FOLIO_PLAYGROUND_VUE_PORT"]) || 4201,
  },
] as const;

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const cell = (text: string) =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;

/** A 2 x 2 table between two paragraphs. */
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
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>HEAD</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell("AAA")}${cell("BBB")}</w:tr><w:tr>${cell("CCC")}${cell("DDD")}</w:tr></w:tbl><w:p><w:r><w:t>TAIL</w:t></w:r></w:p></w:body></w:document>`,
  );
  return [...new Uint8Array(await zip.generate())];
};

const open = async (page: Page, port: number) => {
  await page.goto(`http://localhost:${port}/?session=canonical`);
  await page.waitForSelector(".layout-page");
  expect(await evaluateCanonicalLoad(page, await source())).toBe(true);
  await expect(page.locator(".layout-page").first()).toContainText("TAIL");
};

/** Each saved table as its rows' cell texts. */
const savedTables = async (page: Page) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioPlayground?.getEditorRef()?.save();
    return buffer ? Array.from(new Uint8Array(buffer)) : null;
  });
  expect(saved).not.toBeNull();
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  return [...xml.matchAll(/<w:tbl>.*?<\/w:tbl>/gsu)].map(([table]) =>
    [...table.matchAll(/<w:tr[ >].*?<\/w:tr>/gsu)].map(([row]) =>
      [...row.matchAll(/<w:tc>.*?<\/w:tc>/gsu)].map(([tableCell]) =>
        [...tableCell.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)]
          .map(([, text]) => text)
          .join(""),
      ),
    ),
  );
};

const openMenuOn = async (page: Page, text: string): Promise<Locator> => {
  const target = page
    .locator(".layout-table-cell")
    .filter({ hasText: text })
    .first();
  await target.scrollIntoViewIfNeeded();
  const box = await target.boundingBox();
  if (!box) throw new Error(`no painted cell for ${text}`);
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  await page.mouse.click(x, y);
  await page.mouse.click(x, y, { button: "right" });
  // React's menu is a labelled ARIA menu; Vue's is a list of buttons.
  const menu = page
    .locator('[role="menu"][aria-label="Text editing menu"], .ctx-menu')
    .first();
  await menu.waitFor({ state: "visible", timeout: 5000 });
  return menu;
};

const choose = async (page: Page, text: string, item: string) => {
  const menu = await openMenuOn(page, text);
  await menu
    .locator('[role="menuitem"], button')
    .filter({ hasText: new RegExp(`^\\s*${item}\\s*$`, "u") })
    .first()
    .click();
  await page.waitForTimeout(250);
};

/** Every menu action, from the 2 x 2 table, with the table it leaves. */
const ACTIONS = [
  {
    item: "Insert row above",
    on: "CCC",
    table: [
      ["AAA", "BBB"],
      ["", ""],
      ["CCC", "DDD"],
    ],
  },
  {
    item: "Insert row below",
    on: "AAA",
    table: [
      ["AAA", "BBB"],
      ["", ""],
      ["CCC", "DDD"],
    ],
  },
  { item: "Delete row", on: "AAA", table: [["CCC", "DDD"]] },
  {
    item: "Insert column left",
    on: "BBB",
    table: [
      ["AAA", "", "BBB"],
      ["CCC", "", "DDD"],
    ],
  },
  {
    item: "Insert column right",
    on: "AAA",
    table: [
      ["AAA", "", "BBB"],
      ["CCC", "", "DDD"],
    ],
  },
  { item: "Delete column", on: "AAA", table: [["BBB"], ["DDD"]] },
] as const;

for (const adapter of adapters) {
  test.describe(`the table menu on the canonical engine in ${adapter.name}`, () => {
    for (const { item, on, table } of ACTIONS) {
      test(`${item} changes the table`, async ({ page }) => {
        const refusals: string[] = [];
        page.on("console", (message) => {
          if (message.text().includes("unavailable in this session"))
            refusals.push(message.text());
        });
        await open(page, adapter.port);
        await choose(page, on, item);
        expect(await savedTables(page)).toEqual([table]);
        expect(refusals).toEqual([]);
      });
    }

    test("Insert Table adds a table", async ({ page }) => {
      await open(page, adapter.port);
      await page
        .locator(".layout-paragraph")
        .filter({ hasText: "TAIL" })
        .click();
      if (adapter.name === "React") {
        const insert = page.locator('[data-testid="toolbar-insert-table"]');
        await expect(insert).toBeVisible();
        await insert.click();
      } else {
        // Vue's Insert > Table picker runs the same core insertTableInView.
        expect(
          await page.evaluate(
            () => window.__folioParity?.insertTable(2, 3) ?? false,
          ),
        ).toBe(true);
      }
      await page.waitForTimeout(250);
      const tables = await savedTables(page);
      expect(tables).toHaveLength(2);
    });
  });
}
