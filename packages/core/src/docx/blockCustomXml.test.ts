import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { compareDocx } from "../compare/compare";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { getAllParagraphs, getAllTables, getDocumentText } from "./documentParser";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { docxToMarkdown } from "./server/docxToMarkdown";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const paragraph = (text: string, id: string): string =>
  `<w:p w14:paraId="${id}"><w:r><w:t>${text}</w:t></w:r></w:p>`;
const table = (text: string): string =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr/>${paragraph(text, "44444444")}</w:tc></w:tr></w:tbl>`;
const wrap = (content: string): string =>
  `<w:customXml w:element="clause" w:uri="urn:folio:test">` +
  `<w:customXmlPr><w:attr w:name="role" w:val="operative"/></w:customXmlPr>` +
  `${content}</w:customXml>`;

const makeDocx = async (body: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">` +
      `<w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentXml = async (bytes: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(bytes)).file("word/document.xml")?.async("text")) ?? "";

describe("block-level custom XML", () => {
  test("parses multiple paragraphs and a table as children in source order", async () => {
    const body = wrap(
      paragraph("First clause.", "11111111") +
        paragraph("Second clause.", "22222222") +
        table("Table clause."),
    );
    const parsed = await parseDocx(await makeDocx(body), { preloadFonts: false });
    const block = parsed.package.document.content.at(0);

    expect(block?.type).toBe("blockCustomXml");
    if (block?.type !== "blockCustomXml") {
      return;
    }
    expect(block.content.map(({ type }) => type)).toEqual([
      "preservedBlock",
      "paragraph",
      "paragraph",
      "table",
    ]);
    const properties = block.content.at(0);
    if (properties?.type !== "preservedBlock") {
      return;
    }
    expect(properties.xml).toContain("<w:customXmlPr>");
    expect(block.openingXml).toContain('w:element="clause"');
    expect(block.openingXml).toContain('w:uri="urn:folio:test"');
    expect(block.closingXml).toContain("</w:customXml>");
    expect(getAllParagraphs(parsed.package.document)).toHaveLength(3);
    expect(getAllTables(parsed.package.document)).toHaveLength(1);
    expect(getDocumentText(parsed.package.document)).toContain("Second clause.");
  });

  test("nested wrappers and table cells reach every reader", async () => {
    const body =
      paragraph("Before.", "10000000") +
      wrap(
        paragraph("First clause.", "11111111") +
          wrap(paragraph("Nested clause.", "22222222")) +
          table("Table clause.") +
          paragraph("Last clause.", "33333333"),
      ) +
      paragraph("After.", "50000000");
    const bytes = await makeDocx(body);
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes);
    const expected = [
      "Before.",
      "First clause.",
      "Nested clause.",
      "Table clause.",
      "Last clause.",
      "After.",
    ];

    expect(reviewer.getContent().map(({ text }) => text)).toEqual(expected);
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual(expected);
    const text = reviewer.getContentAsText();
    const markdown = await docxToMarkdown(bytes);
    for (const clause of expected) {
      expect(text).toContain(clause);
      expect(markdown).toContain(clause);
    }
  });

  test("a custom XML wrapper inside a table cell keeps its paragraph readable", async () => {
    const body =
      `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr/>${wrap(paragraph("Cell clause.", "11111111"))}</w:tc></w:tr></w:tbl>`;
    const bytes = await makeDocx(body);
    const parsed = await parseDocx(bytes, { preloadFonts: false });
    const block = parsed.package.document.content.at(0);
    expect(block?.type).toBe("table");
    if (block?.type !== "table") {
      return;
    }
    expect(block.rows.at(0)?.cells.at(0)?.content.at(0)?.type).toBe("blockCustomXml");
    expect(
      (await FolioDocxReviewer.fromBuffer(bytes)).getContent().map(({ text }) => text),
    ).toEqual(["Cell clause."]);
    expect(await docxToMarkdown(bytes)).toContain("Cell clause.");
  });

  test("comparison sees a change confined to a wrapper", async () => {
    const base = await makeDocx(
      wrap(paragraph("Original clause.", "11111111")) + paragraph("Same.", "22222222"),
    );
    const revised = await makeDocx(
      wrap(paragraph("Revised clause.", "11111111")) + paragraph("Same.", "22222222"),
    );

    const result = await compareDocx(base, revised, {
      author: "Comparison",
      timestamp: "2026-09-27T00:00:00.000Z",
    });
    expect(result.isOk()).toBe(true);
    if (result.isErr()) {
      return;
    }
    expect(result.value.changes.length).toBeGreaterThan(0);
    expect(result.value.unsupported).toEqual([]);
    expect(result.value.verification.status).toBe("verified");
  });

  test("direct editing changes a wrapped paragraph and keeps the wrapper", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await makeDocx(wrap(paragraph("Original clause.", "11111111"))),
    );
    const block = reviewer.snapshot().blocks.find(({ text }) => text === "Original clause.");
    expect(block).toBeDefined();
    if (!block) {
      return;
    }
    const applied = reviewer.applyOperations(
      [
        {
          id: "edit-clause",
          type: "replaceInBlock",
          blockId: block.id,
          find: "Original",
          replace: "Revised",
        },
      ],
      { mode: "direct" },
    );
    expect(applied.skipped).toEqual([]);

    const saved = await reviewer.toBuffer();
    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    expect(reopened.getContent().map(({ text }) => text)).toEqual(["Revised clause."]);
    const xml = await documentXml(saved);
    expect(xml).toContain('w:element="clause"');
    expect(xml).toContain('w:uri="urn:folio:test"');
    expect(xml).toContain('<w:attr w:name="role" w:val="operative"/>');
  });

  test("save and editor round trip keep wrapper metadata and nested text", async () => {
    const body = wrap(
      paragraph("First clause.", "11111111") +
        wrap(paragraph("Nested clause.", "22222222")) +
        table("Table clause."),
    );
    const parsed = await parseDocx(await makeDocx(body), { preloadFonts: false });
    for (const document of [parsed, fromProseDoc(toProseDoc(parsed), parsed)]) {
      const saved = await repackDocx(document, { updateModifiedDate: false });
      const xml = await documentXml(saved);
      expect(xml.match(/<w:customXml\b/gu)).toHaveLength(2);
      expect(xml).toContain('w:element="clause"');
      expect(xml).toContain('w:uri="urn:folio:test"');
      expect(xml).toContain('<w:attr w:name="role" w:val="operative"/>');
      expect(xml).toContain("First clause.");
      expect(xml).toContain("Nested clause.");
      expect(xml).toContain("Table clause.");
      expect(
        (await FolioDocxReviewer.fromBuffer(saved)).getContent().map(({ text }) => text),
      ).toEqual(["First clause.", "Nested clause.", "Table clause."]);
    }
  });
});
