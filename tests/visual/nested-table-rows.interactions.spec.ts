import { expect, type Page, test } from "@playwright/test";

import { ZipArchive } from "@stll/docx-core/zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const FILLER =
  "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor incididunt ut labore et dolore.";

const paragraph = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
const cell = (twips: number, content: string): string =>
  `<w:tc><w:tcPr><w:tcW w:w="${String(twips)}" w:type="dxa"/></w:tcPr>${content}</w:tc>`;
const table = (widths: readonly number[], rows: readonly string[]): string =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${widths
    .map((width) => `<w:gridCol w:w="${String(width)}"/>`)
    .join(
      "",
    )}</w:tblGrid>${rows.map((row) => `<w:tr>${row}</w:tr>`).join("")}</w:tbl>`;

const rowToken = (index: number): string =>
  `ROW${String(index).padStart(4, "0")}X`;

/** A one-cell table holding a table much taller than a page. */
const documentWithNestedRows = async (rowCount: number): Promise<Buffer> => {
  const rows = Array.from({ length: rowCount }, (_, index) =>
    [
      cell(4000, paragraph(`${rowToken(index + 1)} ${FILLER}`)),
      cell(4000, paragraph(FILLER)),
    ].join(""),
  );
  const body = table(
    [9000],
    [
      cell(
        9000,
        table([4000, 4000], rows) + paragraph("TEXT-AFTER-THE-INNER-TABLE"),
      ),
    ],
  );
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
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}"><w:body>${paragraph("DOCUMENT-START")}${body}${paragraph("DOCUMENT-END")}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  );
  return Buffer.from(await zip.generate({ type: "uint8array" }));
};

type PageRows = {
  /** Rows with text the page clips nowhere. */
  readonly visible: readonly number[];
  /** Rows with text anywhere in the page's DOM, clipped or not. */
  readonly built: readonly number[];
  readonly endVisible: boolean;
};

/** Reads one page by real geometry: a clipped node's rect misses its clip box. */
const readPageRows = async (page: Page, index: number): Promise<PageRows> => {
  const element = page.locator(".layout-page").nth(index);
  await element.scrollIntoViewIfNeeded();
  await expect(element).toContainText(/ROW\d{4}X|DOCUMENT-END/u);
  return element.evaluate((root) => {
    const pageBox = root.getBoundingClientRect();
    const visible = new Set<number>();
    const built = new Set<number>();
    let endVisible = false;
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      const node = walker.currentNode;
      const text = node.textContent ?? "";
      const match = /ROW(\d{4})X/u.exec(text);
      const isEnd = text.includes("TEXT-AFTER-THE-INNER-TABLE");
      if (!match && !isEnd) continue;
      const range = document.createRange();
      range.selectNodeContents(node);
      const box = range.getBoundingClientRect();
      let clipped = box.top >= pageBox.bottom || box.bottom <= pageBox.top;
      for (
        let ancestor = node.parentElement;
        ancestor && ancestor !== root && !clipped;
        ancestor = ancestor.parentElement
      ) {
        if (getComputedStyle(ancestor).overflow === "visible") continue;
        const clip = ancestor.getBoundingClientRect();
        clipped = box.top >= clip.bottom || box.bottom <= clip.top;
      }
      if (match?.[1]) {
        const row = Number(match[1]);
        built.add(row);
        if (!clipped) visible.add(row);
      } else if (!clipped) {
        endVisible = true;
      }
    }
    return { visible: [...visible], built: [...built], endVisible };
  });
};

const loadDocument = async (page: Page, query: string, rowCount: number) => {
  await page.goto(`/${query}`);
  await page.locator("#file-input").setInputFiles({
    name: "nested-rows.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: await documentWithNestedRows(rowCount),
  });
  await expect(page.locator(".layout-page").first()).toContainText(
    "DOCUMENT-START",
  );
};

for (const { painter, query } of [
  { painter: "default", query: "?session=canonical" },
  {
    painter: "display-list",
    query: "?session=canonical&pageRenderer=display-list",
  },
]) {
  test(`every row of a nested table taller than a page shows on the ${painter} painter`, async ({
    page,
  }) => {
    test.setTimeout(90_000);
    const rowCount = 120;
    await loadDocument(page, query, rowCount);

    await expect
      .poll(() => page.locator(".layout-page").count())
      .toBeGreaterThan(5);
    const pageCount = await page.locator(".layout-page").count();

    const shown = new Set<number>();
    let endVisible = false;
    for (let index = 0; index < pageCount; index += 1) {
      const rows = await readPageRows(page, index);
      for (const row of rows.visible) shown.add(row);
      endVisible ||= rows.endVisible;
      // A page holds what it shows, not the whole inner table.
      expect(rows.built.length).toBeLessThanOrEqual(rows.visible.length + 2);
    }

    expect(
      Array.from({ length: rowCount }, (_, index) => index + 1).filter(
        (row) => !shown.has(row),
      ),
    ).toEqual([]);
    expect(endVisible).toBe(true);
  });

  test(`text typed into a nested row on a later page lands in that row on the ${painter} painter`, async ({
    page,
  }) => {
    test.setTimeout(90_000);
    await loadDocument(page, query, 120);
    const target = page.locator(".layout-page").nth(3);
    await target.scrollIntoViewIfNeeded();
    await expect(target).toContainText("ROW");
    const { row, x, y } = await target.evaluate((root) => {
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      const rootBox = root.getBoundingClientRect();
      while (walker.nextNode()) {
        const node = walker.currentNode;
        const match = /ROW(\d{4})X/u.exec(node.textContent ?? "");
        if (!match?.[1]) continue;
        const range = document.createRange();
        range.setStart(node, 0);
        range.setEnd(node, 2);
        const box = range.getBoundingClientRect();
        if (box.top < rootBox.top + 300 || box.bottom > rootBox.bottom - 300) {
          continue;
        }
        return {
          row: match[1],
          x: box.left + 1,
          y: box.top + box.height / 2,
        };
      }
      throw new TypeError("No nested row in the middle of the page");
    });
    await page.mouse.click(x, y);
    await page.keyboard.type("MARK");

    // The click lands inside the token's first letters; which side of them the
    // caret takes is not the question, only that the row is the one clicked.
    const expected = new RegExp(
      `(?:MARKR|RMARK|R)(?:MARKO|O)?(?:MARKW|W)${row}X`,
      "u",
    );
    await expect
      .poll(() => target.evaluate((root) => root.textContent ?? ""))
      .toMatch(expected);
  });
}
