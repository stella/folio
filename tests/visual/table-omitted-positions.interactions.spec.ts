import { expect, type Locator, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

import type { DocxEditorRef } from "../../packages/react/src/components/DocxEditor.props";

declare global {
  var __folioPlayground:
    { getEditorRef: () => DocxEditorRef | null } | undefined;
}

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

type Row = { cells: string[]; before?: number; after?: number };

const SHAPES: Record<string, Row[]> = {
  "short rows": [
    { cells: ["AAA", "BBB", "CCC"] },
    { cells: ["DDD"] },
    { cells: ["EEE", "FFF"] },
  ],
  "a declared gridAfter": [
    { cells: ["AAA", "BBB", "CCC"] },
    { cells: ["DDD"], after: 2 },
    { cells: ["EEE", "FFF"], after: 1 },
  ],
  "a declared gridBefore": [
    { cells: ["AAA", "BBB", "CCC"] },
    { cells: ["DDD"], before: 2 },
    { cells: ["EEE", "FFF"], before: 1 },
  ],
};

const documentWith = async (rows: readonly Row[]): Promise<Buffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  const table = rows
    .map(({ cells, before, after }) => {
      const properties = `${before ? `<w:gridBefore w:val="${String(before)}"/>` : ""}${after ? `<w:gridAfter w:val="${String(after)}"/>` : ""}`;
      return `<w:tr>${properties ? `<w:trPr>${properties}</w:trPr>` : ""}${cells
        .map(
          (text) =>
            `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`,
        )
        .join("")}</w:tr>`;
    })
    .join("");
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="000000"/><w:left w:val="single" w:sz="4" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:color="000000"/><w:right w:val="single" w:sz="4" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>${table}</w:tbl><w:p><w:r><w:t>TAIL</w:t></w:r></w:p></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate());
};

const load = async (page: Page, rows: readonly Row[]) => {
  await page.goto("/");
  await page.locator("#file-input").setInputFiles({
    name: "omitted.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await documentWith(rows),
  });
  await expect(page.locator(".layout-page").first()).toContainText("TAIL");
};

const cellWith = (page: Page, text: string): Locator =>
  page.locator(".layout-table-cell").filter({ hasText: text }).first();

/** Where each painted cell's text stands: left edge and top edge, in pixels. */
const paintedCells = async (page: Page) =>
  page.locator(".layout-table-cell").evaluateAll((cells) =>
    cells.map((cell) => {
      const box = cell.getBoundingClientRect();
      return {
        text: (cell.textContent ?? "").trim(),
        left: Math.round(box.left),
        top: Math.round(box.top),
      };
    }),
  );

const openMenuOn = async (page: Page, text: string): Promise<Locator> => {
  const cell = cellWith(page, text);
  await cell.scrollIntoViewIfNeeded();
  const box = await cell.boundingBox();
  if (!box) throw new Error(`no painted cell for ${text}`);
  const x = box.x + box.width / 2;
  const y = box.y + box.height / 2;
  await page.mouse.click(x, y);
  await page.mouse.click(x, y, { button: "right" });
  const menu = page.locator('[role="menu"][aria-label="Text editing menu"]');
  await menu.waitFor({ state: "visible", timeout: 5000 });
  return menu;
};

const choose = async (page: Page, text: string, item: string) => {
  const menu = await openMenuOn(page, text);
  await menu.getByRole("menuitem", { name: item }).click();
  await page.waitForTimeout(250);
};

const savedRows = async (page: Page) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioPlayground?.getEditorRef()?.save();
    return buffer ? Array.from(new Uint8Array(buffer)) : null;
  });
  expect(saved).not.toBeNull();
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  return [...xml.matchAll(/<w:tr(?:\/>|[ >].*?<\/w:tr>)/gsu)].map(([row]) => {
    let column = Number(/<w:gridBefore w:val="(\d+)"/u.exec(row)?.[1] ?? 0);
    const cells: Record<string, number> = {};
    for (const [cell] of row.matchAll(/<w:tc>.*?<\/w:tc>/gsu)) {
      const text = [...cell.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/gu)]
        .map(([, content]) => content)
        .join("");
      if (text) cells[text] = column;
      column += Number(/<w:gridSpan w:val="(\d+)"/u.exec(cell)?.[1] ?? 1);
    }
    return cells;
  });
};

/** The number of cells a saved row holds, counting from the end with a negative index. */
const savedCellCount = async (page: Page, index: number) => {
  const saved = await page.evaluate(async () => {
    const buffer = await globalThis.__folioPlayground?.getEditorRef()?.save();
    return buffer ? Array.from(new Uint8Array(buffer)) : null;
  });
  const zip = await ZipArchive.read(Uint8Array.from(saved ?? []).buffer);
  const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
  const rows = [...xml.matchAll(/<w:tr(?:\/>|[ >].*?<\/w:tr>)/gsu)];
  return (rows.at(index)?.[0].match(/<w:tc[ >]/gu) ?? []).length;
};

const columnsOf = (rows: readonly Row[]): Record<string, number> =>
  Object.fromEntries(
    rows.flatMap(({ cells, before }) =>
      cells.map((text, index) => [text, index + (before ?? 0)]),
    ),
  );

const flat = (rows: readonly Record<string, number>[]) =>
  Object.assign({}, ...rows) as Record<string, number>;

for (const [name, rows] of Object.entries(SHAPES)) {
  test.describe(`a table with ${name}`, () => {
    const source = columnsOf(rows);

    test("paints only its real cells, each in its own column", async ({
      page,
    }) => {
      await load(page, rows);
      const painted = await paintedCells(page);
      expect(painted.map(({ text }) => text)).toEqual(
        rows.flatMap(({ cells }) => cells),
      );
      // Cells of one grid column share a left edge, whichever row they are in.
      const leftOf = (text: string) =>
        painted.find((cell) => cell.text === text)?.left;
      for (const [text, column] of Object.entries(source)) {
        const header = ["AAA", "BBB", "CCC"][column];
        if (header) expect(leftOf(text)).toBe(leftOf(header));
      }
      expect(await savedRows(page)).toHaveLength(rows.length);
    });

    test("adds and deletes a row from the menu", async ({ page }) => {
      await load(page, rows);
      await choose(page, "DDD", "Insert row below");
      await expect(page.locator(".layout-page").first()).toContainText("DDD");
      const after = await savedRows(page);
      expect(after).toHaveLength(rows.length + 1);
      expect(flat(after)).toEqual(source);

      await choose(page, "AAA", "Delete row");
      const removed = await savedRows(page);
      expect(removed).toHaveLength(rows.length);
      const { AAA: _a, BBB: _b, CCC: _c, ...rest } = source;
      expect(flat(removed)).toEqual(rest);
    });

    test("adds and deletes a column from the menu, moving no text", async ({
      page,
    }) => {
      await load(page, rows);
      await choose(page, "AAA", "Insert column right");
      const added = flat(await savedRows(page));
      expect(added).toEqual(
        Object.fromEntries(
          Object.entries(source).map(([text, column]) => [
            text,
            column >= 1 ? column + 1 : column,
          ]),
        ),
      );
      // On screen too: a cell keeps the left edge of its column's first cell.
      const painted = await paintedCells(page);
      const leftOf = (text: string) =>
        painted.find((cell) => cell.text === text)?.left;
      for (const [text, column] of Object.entries(source)) {
        const header = ["AAA", "BBB", "CCC"][column];
        if (header) expect(leftOf(text)).toBe(leftOf(header));
      }

      await choose(page, "AAA", "Delete column");
      expect(flat(await savedRows(page))).toEqual(
        Object.fromEntries(
          Object.entries(added)
            .filter(([, column]) => column !== 0)
            .map(([text, column]) => [text, column - 1]),
        ),
      );
      await expect(page.locator(".layout-page").first()).toContainText("TAIL");
    });
  });
}

test("typing in a table with short rows leaves its rows as they were", async ({
  page,
}) => {
  const rows = SHAPES["short rows"] ?? [];
  await load(page, rows);
  const cell = cellWith(page, "DDD");
  const box = await cell.boundingBox();
  if (!box) throw new Error("no painted cell");
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await page.keyboard.press("End");
  await page.keyboard.type("X");
  await expect(page.locator(".layout-page").first()).toContainText("DDDX");
  const saved = await savedRows(page);
  expect(saved.map((row) => Object.keys(row).length)).toEqual([3, 1, 2]);
  expect(flat(saved)).toEqual({
    AAA: 0,
    BBB: 1,
    CCC: 2,
    DDDX: 0,
    EEE: 0,
    FFF: 1,
  });
  await expect
    .poll(async () => (await paintedCells(page)).length)
    .toBe(rows.flatMap(({ cells }) => cells).length);
});

test.describe("a table row with no cells", () => {
  for (const [name, properties] of [
    ["nothing", ""],
    ["a declared gridAfter", '<w:gridAfter w:val="3"/>'],
  ] as const)
    for (const session of ["default", "canonical"] as const)
      test(`opens and paints its neighbours with ${name} (${session} session)`, async ({
        page,
      }) => {
        const rows = SHAPES["short rows"] ?? [];
        await page.goto(session === "canonical" ? "/?session=canonical" : "/");
        const zip = await ZipArchive.read(await documentWith(rows));
        const xml = (await zip.entry("word/document.xml")?.text()) ?? "";
        zip.set(
          "word/document.xml",
          xml.replace(
            "</w:tbl>",
            `<w:tr>${properties ? `<w:trPr>${properties}</w:trPr>` : ""}</w:tr></w:tbl>`,
          ),
        );
        await page.locator("#file-input").setInputFiles({
          name: "no-cells.docx",
          mimeType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          buffer: Buffer.from(await zip.generate()),
        });
        const first = page.locator(".layout-page").first();
        await expect(first).toContainText("TAIL");
        await expect(first).toContainText("FFF");
        const saved = await savedRows(page);
        expect(saved).toHaveLength(rows.length + 1);
        // The row saves as the file states it, with no cell it never had.
        expect(saved.at(-1)).toEqual({});
        expect(await savedCellCount(page, -1)).toBe(0);
      });
});
