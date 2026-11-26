import { Result } from "better-result";
import { describe, expect, test } from "bun:test";

import { analyzeDocxBlocks } from "../blocks";
import { inspectDocxTemplate } from "../inspect";
import type { DocxBlock } from "../types";
import {
  buildMinimalDocx,
  buildSource,
  buildTemplate,
  numPr,
  p,
  r,
  readPart,
  replacePart,
  SOURCE_IDS,
  SOURCE_TEXT,
  TEMPLATE_IDS,
} from "./fixtures";

const unwrap = <T>(result: Result<T, unknown>): T => {
  if (Result.isError(result)) {
    throw result.error;
  }
  return result.value;
};

const byId = (blocks: readonly DocxBlock[], id: string): DocxBlock => {
  const block = blocks
    .flatMap((candidate) => [candidate, ...(candidate.nested ?? [])])
    .find((candidate) => candidate.id === id);
  if (block === undefined) {
    throw new Error(`no block ${id}`);
  }
  return block;
};

describe("analyzeDocxBlocks", () => {
  test("reads the source story into ordered, hash-pinned blocks", async () => {
    const manifest = unwrap(await analyzeDocxBlocks(await buildSource()));

    expect(manifest.sha256).toMatch(/^[0-9a-f]{64}$/u);
    expect(manifest.blocks.map(({ id }) => id)).toEqual([
      "p0",
      "p1",
      "p2",
      "p3",
      "p4",
      "p5",
      "p6",
      "p7",
      "p8",
      "p9",
      "p10",
      "t11",
      "p16",
      "p17",
      "p18",
    ]);
    expect(byId(manifest.blocks, SOURCE_IDS.title).text).toBe(
      SOURCE_TEXT.title,
    );
    expect(byId(manifest.blocks, SOURCE_IDS.clause).text).toBe(
      SOURCE_TEXT.clause,
    );
    expect(byId(manifest.blocks, SOURCE_IDS.letterA).text).toBe(
      SOURCE_TEXT.letterA,
    );
    expect(byId(manifest.blocks, SOURCE_IDS.letterB).text).toBe(
      SOURCE_TEXT.letterB,
    );
    expect(byId(manifest.blocks, SOURCE_IDS.mixed).text).toBe(
      SOURCE_TEXT.mixed,
    );
    expect(byId(manifest.blocks, SOURCE_IDS.pageBreak).text).toBe("");
    // The two blank paragraphs and the page break share one hash.
    expect(new Set(manifest.blocks.map(({ hash }) => hash)).size).toBe(
      manifest.blocks.length - 2,
    );
  });

  test("renders automatic labels and leaves manual numbering as text", async () => {
    const { blocks } = unwrap(await analyzeDocxBlocks(await buildSource()));

    expect(byId(blocks, SOURCE_IDS.clause)).toMatchObject({
      label: "1.",
      numbering: { numId: 1, ilvl: 0, abstractNumId: 0 },
    });
    expect(byId(blocks, SOURCE_IDS.letterA).label).toBe("a)");
    expect(byId(blocks, SOURCE_IDS.letterB).label).toBe("b)");
    expect(byId(blocks, SOURCE_IDS.manualClause)).toMatchObject({
      label: null,
      numbering: null,
      text: SOURCE_TEXT.manualClause,
    });
    expect(byId(blocks, SOURCE_IDS.articleOne)).toMatchObject({
      styleId: "Heading1",
      styleName: "heading 1",
      outlineLevel: 0,
    });
  });

  test("lists a table as one block with its cells and nested paragraphs", async () => {
    const { blocks } = unwrap(await analyzeDocxBlocks(await buildSource()));
    const table = byId(blocks, SOURCE_IDS.table);

    expect(table.kind).toBe("table");
    expect(table.text).toBe("Item\tPrice\nService\t100");
    expect(table.cells).toEqual([
      { row: 0, col: 0, blockIds: ["p12"] },
      { row: 0, col: 1, blockIds: ["p13"] },
      { row: 1, col: 0, blockIds: ["p14"] },
      { row: 1, col: 1, blockIds: ["p15"] },
    ]);
    expect(table.nested?.map(({ id, text }) => [id, text])).toEqual([
      ["p12", "Item"],
      ["p13", "Price"],
      ["p14", "Service"],
      ["p15", "100"],
    ]);
  });

  test("reports notes, comments and tracked changes", async () => {
    const manifest = unwrap(await analyzeDocxBlocks(await buildSource()));

    expect(manifest.footnotes).toEqual([
      { id: "2", text: SOURCE_TEXT.footnote },
    ]);
    expect(manifest.endnotes).toEqual([]);
    expect(manifest.comments).toBe(1);
    expect(manifest.trackedChanges).toBe(0);

    const tracked = unwrap(
      await analyzeDocxBlocks(
        await buildMinimalDocx(
          p(
            `${r("kept ")}<w:ins w:id="1" w:author="a" w:date="2020-01-01T00:00:00Z">${r("added")}</w:ins><w:del w:id="2" w:author="a" w:date="2020-01-01T00:00:00Z"><w:r><w:delText>removed</w:delText></w:r></w:del>`,
          ),
        ),
      ),
    );
    expect(tracked.trackedChanges).toBe(2);
    expect(tracked.blocks.at(0)?.text).toBe("kept added");
  });

  test("includes field results and leaves field instructions out", async () => {
    const { blocks } = unwrap(await analyzeDocxBlocks(await buildTemplate()));

    expect(byId(blocks, TEMPLATE_IDS.fillIn).text).toBe("To: Recipient name");
    expect(byId(blocks, TEMPLATE_IDS.tocFirst).text).toBe("Sample article\t2");
    expect(byId(blocks, TEMPLATE_IDS.tocLast).text).toBe("Second sample\t3");
  });

  test("a paragraph of a bullet level says so, whatever its mark", async () => {
    const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const { blocks } = unwrap(
      await analyzeDocxBlocks(
        await buildMinimalDocx(
          `${p(r("a point"), numPr(1, 0))}${p(r("an item"), numPr(1, 1))}`,
          {
            "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
            "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`,
            "word/numbering.xml": numbering,
          },
        ),
      ),
    );

    // The mark of a bullet level is a mark; a letter the list counts is not.
    expect(blocks.map(({ label, numbering: level }) => [label, level])).toEqual(
      [
        ["○", { numId: 1, ilvl: 0, abstractNumId: 0, bullet: true }],
        ["a", { numId: 1, ilvl: 1, abstractNumId: 0 }],
      ],
    );
  });

  test("a paragraph numbered without a numbering part has no label", async () => {
    const { blocks } = unwrap(
      await analyzeDocxBlocks(
        await buildMinimalDocx(p(r("orphan"), numPr(7, 0))),
      ),
    );

    expect(blocks.at(0)).toMatchObject({ label: null, numbering: null });
  });

  test("refuses bytes that are not a package", async () => {
    const result = await analyzeDocxBlocks(new Uint8Array([1, 2, 3]));

    expect(Result.isError(result)).toBe(true);
    if (Result.isError(result)) {
      expect(result.error.code).toBe("archive-unreadable");
    }
  });
});

describe("inspectDocxTemplate", () => {
  test("lists the styles with the list each binds", async () => {
    const { styles } = unwrap(await inspectDocxTemplate(await buildTemplate()));
    const style = (id: string) => styles.find(({ styleId }) => styleId === id);

    expect(style("HouseArticle")).toEqual({
      styleId: "HouseArticle",
      name: "House Article",
      type: "paragraph",
      basedOn: "Normal",
      numPr: { numId: 1, ilvl: 0 },
      outlineLvl: 0,
      alignment: null,
      isDefault: false,
    });
    expect(style("HouseTitle")?.alignment).toBe("center");
    expect(style("HouseClause")?.numPr).toEqual({ numId: 1, ilvl: 1 });
    // The style names the list only; the level that names the style back decides.
    expect(style("HouseSubClause")?.numPr).toEqual({ numId: 1, ilvl: 2 });
    expect(style("HouseLetter")?.numPr).toEqual({ numId: 1, ilvl: 3 });
    expect(style("HouseBody")?.numPr).toBeNull();
    expect(style("Normal")?.isDefault).toBe(true);
    expect(style("HouseTable")?.type).toBe("table");
  });

  test("lists numbering definitions level by level", async () => {
    const { numbering } = unwrap(
      await inspectDocxTemplate(await buildTemplate()),
    );

    expect(numbering.nums).toEqual([
      { numId: 1, abstractNumId: 0, overrides: [] },
      { numId: 2, abstractNumId: 1, overrides: [] },
    ]);
    expect(
      numbering.abstractNums
        .at(0)
        ?.levels.map((level) => [
          level.ilvl,
          level.numFmt,
          level.lvlText,
          level.start,
          level.suff,
          level.pStyle,
        ]),
    ).toEqual([
      [0, "decimal", "Art. %1", 1, "tab", "HouseArticle"],
      [1, "decimal", "%1.%2.", 1, "tab", "HouseClause"],
      [2, "decimal", "%1.%2.%3.", 1, "tab", "HouseSubClause"],
      [3, "lowerLetter", "%4)", 1, "tab", "HouseLetter"],
      [4, "lowerRoman", "(%5)", 1, "tab", null],
    ]);
    expect(
      numbering.abstractNums.at(0)?.levels.map(({ indent }) => indent),
    ).toEqual(
      [567, 1134, 1701, 2268, 2835].map((left) => ({
        left,
        hanging: 567,
        firstLine: null,
      })),
    );
  });

  test("resolves a style's alignment through the styles it is based on", async () => {
    const base = await buildTemplate();
    const styles = (await readPart(base, "word/styles.xml")) ?? "";
    const template = await replacePart(
      base,
      "word/styles.xml",
      styles.replace(
        '<w:name w:val="House Article"/><w:basedOn w:val="Normal"/>',
        '<w:name w:val="House Article"/><w:basedOn w:val="HouseTitle"/>',
      ),
    );
    const inspection = unwrap(await inspectDocxTemplate(template));
    const alignmentOf = (id: string) =>
      inspection.styles.find(({ styleId }) => styleId === id)?.alignment;

    expect(alignmentOf("HouseArticle")).toBe("center");
    expect(alignmentOf("HouseClause")).toBeNull();
  });

  test("renders the template's own labels", async () => {
    const { blocks } = unwrap(await inspectDocxTemplate(await buildTemplate()));

    expect(byId(blocks, TEMPLATE_IDS.article).label).toBe("Art. 1");
    expect(byId(blocks, TEMPLATE_IDS.clause).label).toBe("1.1.");
  });

  test("inventories fields across the body, header and footer", async () => {
    const { fields } = unwrap(await inspectDocxTemplate(await buildTemplate()));

    expect(
      fields.map(({ part, kind, blockId, endBlockId }) => [
        part,
        kind,
        blockId ?? null,
        endBlockId ?? null,
      ]),
    ).toEqual([
      // The table of contents field spans two blocks; a plan keeps both.
      ["word/document.xml", "TOC", TEMPLATE_IDS.tocFirst, TEMPLATE_IDS.tocLast],
      [
        "word/document.xml",
        "PAGEREF",
        TEMPLATE_IDS.tocFirst,
        TEMPLATE_IDS.tocFirst,
      ],
      [
        "word/document.xml",
        "PAGEREF",
        TEMPLATE_IDS.tocLast,
        TEMPLATE_IDS.tocLast,
      ],
      ["word/document.xml", "FILLIN", TEMPLATE_IDS.fillIn, TEMPLATE_IDS.fillIn],
      ["word/header1.xml", "PAGE", null, null],
      ["word/footer1.xml", "DATE", null, null],
      ["word/footer1.xml", "FILENAME", null, null],
    ]);
    expect(fields.find(({ kind }) => kind === "FILLIN")).toMatchObject({
      instruction: 'FILLIN "Recipient"',
      result: "Recipient name",
    });
    expect(fields.find(({ kind }) => kind === "TOC")?.result).toBe(
      "Sample article\t2\nSecond sample\t3",
    );
  });

  test("describes sections and bookmarks", async () => {
    const { sections, bookmarks } = unwrap(
      await inspectDocxTemplate(await buildTemplate()),
    );

    expect(sections).toEqual([
      {
        index: 0,
        start: "nextPage",
        orientation: "portrait",
        columns: 1,
        size: { width: 11_906, height: 16_838 },
        margins: { top: 1417, right: 1417, bottom: 1134, left: 1417 },
        headers: [{ type: "default", part: "word/header1.xml" }],
        footers: [{ type: "default", part: "word/footer1.xml" }],
      },
    ]);
    expect(bookmarks).toEqual([
      { name: "_Toc100000001", id: "1", part: "word/document.xml" },
    ]);
  });
});
