/**
 * A table cell that ends in a table, or holds no paragraph at all, is read
 * with the empty paragraph it implies. Paragraph id normalization writes that
 * paragraph, so it carries an id as every other paragraph does, and both
 * engines open the package.
 */
import { expect, test } from "bun:test";

import { buildRawPackage, SECTION } from "../__tests__/documentShapes";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { ensureParaIds } from "./ensureParaIds";
import { visitDocxParagraphs } from "./paragraphTraversal";
import { parseDocx } from "./parser";

const cellText = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;

const innerTable = `<w:tbl><w:tr><w:tc>${cellText("Inner")}</w:tc></w:tr></w:tbl>`;

const CELLS = [
  {
    name: "ending in a table",
    cell: `<w:tc>${cellText("Outer")}${innerTable}</w:tc>`,
  },
  {
    name: "ending in a table inside a content control",
    cell: `<w:tc>${cellText("Outer")}<w:sdt><w:sdtContent>${innerTable}</w:sdtContent></w:sdt></w:tc>`,
  },
  { name: "holding nothing", cell: "<w:tc><w:tcPr/></w:tc>" },
  { name: "written self-closing", cell: "<w:tc/>" },
] as const;

const packageWith = async (cell: string): Promise<ArrayBuffer> =>
  (
    await buildRawPackage({
      body: `<w:tbl><w:tr>${cell}</w:tr></w:tbl><w:p/>${SECTION()}`,
      styles: null,
    })
  ).slice().buffer;

for (const { name, cell } of CELLS) {
  test(`a cell ${name} gets a paragraph with an id`, async () => {
    const ensured = await ensureParaIds(await packageWith(cell));
    const document = await parseDocx(ensured.docx.slice().buffer, {
      preloadFonts: false,
    });
    const ids: (string | undefined)[] = [];
    visitDocxParagraphs(
      { documentBody: document.package.document },
      (paragraph) => {
        ids.push(paragraph.paraId);
      },
    );
    expect(ids.length).toBeGreaterThan(0);
    expect(ids.every((id) => id !== undefined)).toBe(true);
    expect(new Set(ids).size).toBe(ids.length);
    // Normalizing a normalized package changes nothing.
    expect((await ensureParaIds(ensured.docx)).alreadyComplete).toBe(true);
  });

  test(`a cell ${name} reads alike on both engines`, async () => {
    const buffer = await packageWith(cell);
    const [legacy, canonical] = await Promise.all(
      (["canonical"] as const).map((session) =>
        FolioDocxReviewer.fromBuffer(buffer.slice(0), { session }),
      ),
    );
    const texts = (reviewer: FolioDocxReviewer | undefined) =>
      reviewer?.snapshot().blocks.map(({ text }) => text);
    expect(texts(canonical)).toEqual(texts(legacy));
  });
}
