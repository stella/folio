import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { createStellaStyleDocumentPreset } from "../../style-sets/stellaStyle";
import type {
  BlockContent,
  Document,
  Paragraph,
  Style,
  Table,
  TableCell,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { ensureParaIds } from "../ensureParaIds";
import { getParagraphText } from "../paragraphParser";
import { parseDocx } from "../parser";
import { createDocx } from "../rezip";
import {
  createBilingualDocument,
  InvalidBilingualDocumentOptionsError,
} from "./createBilingualDocument";
import { createBilingualDocx, readBilingualDocx } from "./createBilingualDocx";

const SUFFIX = "en";

const paragraph = (
  text: string,
  styleId?: string,
  numPr?: { numId: number; ilvl: number },
): Paragraph => ({
  type: "paragraph",
  formatting: {
    ...(styleId !== undefined && { styleId }),
    ...(numPr !== undefined && { numPr }),
  },
  content: text === "" ? [] : [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
});

const cell = (content: (Paragraph | Table)[]): TableCell => ({ type: "tableCell", content });

const sourceTable = (): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [cell([paragraph("Cell A1")]), cell([paragraph("Cell B1")])],
    },
  ],
});

/**
 * A contract-shaped source: style-linked clause numbering from the Stella
 * preset (ClauseHeading / ClauseParagraph1), a directly numbered list that
 * reuses the preset's bullet instance, a nested table, an empty paragraph and
 * a section break. Serialized and re-parsed so the document carries
 * `originalBuffer`, paraIds and style-materialized `numPr`, exactly like a
 * document loaded from disk.
 */
const buildSource = async (): Promise<Document> => {
  const doc = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
  const bulletNumId = findStyle(doc, "ListParagraph")?.pPr?.numPr?.numId;
  if (bulletNumId === undefined) {
    throw new Error("preset lost its bullet numbering");
  }
  const firstSection: BlockContent[] = [
    paragraph("Preamble", "Normal"),
    paragraph("Definitions", "ClauseHeading1"),
    paragraph("Agreement means this contract.", "ClauseParagraph1"),
    paragraph(""),
    paragraph("first bullet", "Normal", { numId: bulletNumId, ilvl: 0 }),
    sourceTable(),
  ];
  const sectionBreak: Paragraph = {
    type: "paragraph",
    content: [],
    sectionProperties: { pageWidth: 11_906, pageHeight: 16_838 },
  };
  const secondSection: BlockContent[] = [
    paragraph("Term", "ClauseHeading1"),
    paragraph("This Agreement lasts one year.", "ClauseParagraph1"),
  ];
  doc.package.document.content = [...firstSection, sectionBreak, ...secondSection];
  const stamped = await ensureParaIds(await createDocx(doc));
  return parseDocx(stamped.docx, { preloadFonts: false });
};

const bilingualDocumentOptions = async (
  source: Document,
): Promise<Parameters<typeof createBilingualDocument>[1]> => ({
  targetStyleSuffix: SUFFIX,
  editableParagraphIds: new Set(
    (await FolioDocxReviewer.fromBuffer(source.originalBuffer ?? (await createDocx(source))))
      .snapshot()
      .blocks.map(({ id }) => id),
  ),
});

const findStyle = (doc: Document, styleId: string): Style | undefined =>
  doc.package.styles?.styles.find((style) => style.styleId === styleId);

const bodyTables = (doc: Document): Table[] =>
  doc.package.document.content.filter((block): block is Table => block.type === "table");

/**
 * Paragraphs in a bilingual cell, excluding the empty filler paragraph a cell
 * must end with when it holds a nested table (OOXML forbids a bare `w:tbl`
 * as the last cell child).
 */
const cellParagraphs = (tableCell: TableCell): Paragraph[] => {
  const out: Paragraph[] = [];
  const visit = (item: Paragraph | Table): void => {
    if (item.type === "paragraph") {
      if (item.content.length > 0) {
        out.push(item);
      }
      return;
    }
    for (const row of item.rows) {
      for (const inner of row.cells) {
        inner.content.forEach(visit);
      }
    }
  };
  tableCell.content.forEach(visit);
  return out;
};

type ColumnParagraphs = { left: Paragraph[]; right: Paragraph[] };

const columnParagraphs = (doc: Document): ColumnParagraphs => {
  const left: Paragraph[] = [];
  const right: Paragraph[] = [];
  for (const table of bodyTables(doc)) {
    for (const row of table.rows) {
      const [leftCell, rightCell] = row.cells;
      if (!leftCell) {
        throw new Error("bilingual row without cells");
      }
      if (!rightCell) {
        // A spanning row carries a source table once; not a column pair.
        continue;
      }
      left.push(...cellParagraphs(leftCell));
      right.push(...cellParagraphs(rightCell));
    }
  }
  return { left, right };
};

const effectiveNumId = (doc: Document, p: Paragraph): number | undefined => {
  const direct = p.formatting?.numPr?.numId;
  if (direct !== undefined) {
    return direct;
  }
  const style = p.formatting?.styleId ? findStyle(doc, p.formatting.styleId) : undefined;
  return style?.pPr?.numPr?.numId;
};

const abstractNumIdOf = (doc: Document, numId: number): number | undefined =>
  doc.package.numbering?.nums.find((num) => num.numId === numId)?.abstractNumId;

describe("createBilingualDocument", () => {
  test("lays every non-empty block out as a two-column row, one table per section", async () => {
    const source = await buildSource();
    const { document, rows } = createBilingualDocument(
      source,
      await bilingualDocumentOptions(source),
    );

    const content = document.package.document.content;
    expect(content.map((block) => block.type)).toEqual(["table", "paragraph", "table"]);
    const [first, second] = bodyTables(document);
    // Preamble, heading, clause, bullet, table (the empty paragraph is dropped).
    expect(first?.rows).toHaveLength(5);
    expect(second?.rows).toHaveLength(2);
    expect(rows.map((row) => row.kind)).toEqual([
      "paragraph",
      "heading",
      "listItem",
      "listItem",
      "table",
      "heading",
      "listItem",
    ]);
    expect(content[1]?.type === "paragraph" && content[1].sectionProperties).toBeTruthy();
  });

  test("keeps the left column content and identities from the source blocks", async () => {
    const source = await buildSource();
    const { document } = createBilingualDocument(source, await bilingualDocumentOptions(source));

    const sourceParagraphs = source.package.document.content.flatMap((block): Paragraph[] =>
      block.type === "paragraph" && !block.sectionProperties && block.content.length > 0
        ? [block]
        : [],
    );
    expect(
      columnParagraphs(document).left.map(({ formatting: _formatting, ...content }) => content),
    ).toEqual(sourceParagraphs.map(({ formatting: _formatting, ...content }) => content));
  });

  test("projects direct and inherited full-page geometry into each column", async () => {
    const source = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    source.package.styles?.styles.push({
      styleId: "FormLine",
      type: "paragraph",
      name: "Form Line",
      pPr: {
        indentLeft: 5040,
        indentRight: -360,
        tabs: [
          { position: 4320, alignment: "left" },
          { position: 9648, alignment: "left" },
        ],
      },
    });
    source.package.document.content = [
      {
        ...paragraph("Signature field", "FormLine"),
        formatting: {
          styleId: "FormLine",
          indentFirstLine: 540,
        },
      },
    ];
    const stamped = await ensureParaIds(await createDocx(source));
    const parsed = await parseDocx(stamped.docx, { preloadFonts: false });
    const { document } = createBilingualDocument(parsed, await bilingualDocumentOptions(parsed));
    const { left, right } = columnParagraphs(document);

    for (const projected of [left.at(0), right.at(0)]) {
      expect(projected?.formatting).toMatchObject({
        indentLeft: 2520,
        indentRight: 0,
        indentFirstLine: 270,
        tabs: [
          { position: 2160, alignment: "left" },
          { position: 4153, alignment: "left" },
        ],
      });
    }
    expect(getParagraphText(left.at(0)!)).toBe("Signature field");
    expect(getParagraphText(right.at(0)!)).toBe("Signature field");
  });

  test("keeps a source table once, in a row spanning both columns", async () => {
    const source = await buildSource();
    const { document, rows } = createBilingualDocument(
      source,
      await bilingualDocumentOptions(source),
    );

    const allRows = bodyTables(document).flatMap((table) => table.rows);
    const spanningRows = allRows.filter((row) => row.cells.length === 1);
    expect(spanningRows).toHaveLength(1);
    expect(allRows.every((row) => row.cells.length === 1 || row.cells.length === 2)).toBe(true);

    const spanningCell = spanningRows[0]!.cells[0]!;
    expect(spanningCell.formatting?.gridSpan).toBe(2);
    const nested = spanningCell.content.filter((item): item is Table => item.type === "table");
    expect(nested).toHaveLength(1);
    const originalTable = source.package.document.content.find(
      (block): block is Table => block.type === "table",
    );
    expect(nested[0]).toEqual(originalTable!);

    // The manifest points at the table's own paragraphs (translated in place).
    const tableRow = rows.find((row) => row.kind === "table");
    const tableParaIds = originalTable!.rows.flatMap((row) =>
      row.cells.flatMap((c) => cellParagraphs(c).map((p) => p.paraId)),
    );
    expect(tableRow?.kind === "table" && tableRow.paragraphs.map((p) => p.paraId)).toEqual(
      tableParaIds,
    );
  });

  test("right column never shares a numbering instance or abstract with the left", async () => {
    const source = await buildSource();
    const { document } = createBilingualDocument(source, await bilingualDocumentOptions(source));
    const { left, right } = columnParagraphs(document);
    expect(right).toHaveLength(left.length);

    const leftNumIds = new Set(
      left.map((p) => effectiveNumId(document, p)).filter((id) => id !== undefined),
    );
    const leftAbstracts = new Set(
      [...leftNumIds].map((id) => abstractNumIdOf(document, id)).filter((id) => id !== undefined),
    );
    expect(leftNumIds.size).toBeGreaterThan(1);

    let numberedPairs = 0;
    for (const [index, rightParagraph] of right.entries()) {
      const leftParagraph = left[index];
      if (!leftParagraph) {
        throw new Error("column length mismatch");
      }
      const leftNumId = effectiveNumId(document, leftParagraph);
      const rightNumId = effectiveNumId(document, rightParagraph);
      if (leftNumId === undefined) {
        expect(rightNumId).toBeUndefined();
        continue;
      }
      numberedPairs += 1;
      expect(rightNumId).toBeDefined();
      expect(leftNumIds.has(rightNumId!)).toBe(false);
      expect(leftAbstracts.has(abstractNumIdOf(document, rightNumId!)!)).toBe(false);
      // Same level, same marker template: the clone renders like the source.
      expect(rightParagraph.listRendering?.marker).toBe(leftParagraph.listRendering?.marker);
      expect(rightParagraph.listRendering?.level).toBe(leftParagraph.listRendering?.level);
    }
    expect(numberedPairs).toBe(5);
  });

  test("clones numbered paragraph styles with the suffix and rewrites only their numPr", async () => {
    const source = await buildSource();
    const { document } = createBilingualDocument(source, await bilingualDocumentOptions(source));

    const original = findStyle(source, "ClauseParagraph1");
    const clone = findStyle(document, `ClauseParagraph1-${SUFFIX}`);
    expect(original).toBeDefined();
    expect(clone).toBeDefined();
    expect(clone?.pPr?.numPr?.numId).not.toBe(original?.pPr?.numPr?.numId);
    expect(clone?.pPr?.numPr?.ilvl).toBe(original?.pPr?.numPr?.ilvl);
    expect({
      ...clone,
      styleId: "",
      name: "",
      pPr: undefined,
      next: undefined,
      default: undefined,
    }).toEqual({
      ...original,
      styleId: "",
      name: "",
      pPr: undefined,
      next: undefined,
      default: undefined,
    });
    // Unnumbered styles are not cloned.
    expect(findStyle(document, `Normal-${SUFFIX}`)).toBeUndefined();
    // The source styles are untouched.
    expect(document.package.styles?.styles.slice(0, source.package.styles?.styles.length)).toEqual(
      source.package.styles?.styles ?? [],
    );
  });

  test("mints stable, unique right-column paraIds and reports them as row handles", async () => {
    const source = await buildSource();
    const options = await bilingualDocumentOptions(source);
    const first = createBilingualDocument(source, options);
    const second = createBilingualDocument(source, options);

    const handles = first.rows.map((row) => row.rowId);
    expect(new Set(handles).size).toBe(handles.length);
    expect(second.rows.map((row) => row.rowId)).toEqual(handles);

    const { left, right } = columnParagraphs(first.document);
    const leftIds = new Set(left.map((p) => p.paraId));
    for (const p of right) {
      expect(p.paraId).toBeDefined();
      expect(leftIds.has(p.paraId)).toBe(false);
    }
    const tableRow = first.rows.find((row) => row.kind === "table");
    expect(tableRow?.kind === "table" && tableRow.paragraphs.map((p) => p.sourceText)).toEqual([
      "Cell A1",
      "Cell B1",
    ]);
  });

  test("rejects a style suffix that cannot form a style id", async () => {
    const source = await buildSource();
    const options = await bilingualDocumentOptions(source);
    expect(() =>
      createBilingualDocument(source, { ...options, targetStyleSuffix: "en US" }),
    ).toThrow(InvalidBilingualDocumentOptionsError);
  });

  test("does not mutate the source document", async () => {
    const source = await buildSource();
    const snapshot = JSON.stringify({ ...source, originalBuffer: undefined });
    createBilingualDocument(source, await bilingualDocumentOptions(source));
    expect(JSON.stringify({ ...source, originalBuffer: undefined })).toBe(snapshot);
  });
});

describe("createBilingualDocx", () => {
  test("round-trips cloned styles and numbering through the repacked package", async () => {
    const source = await buildSource();
    const { buffer, rows } = await createBilingualDocx(source.originalBuffer!, {
      targetStyleSuffix: SUFFIX,
    });
    const reparsed = await parseDocx(buffer, { preloadFonts: false });

    expect(findStyle(reparsed, `ClauseParagraph1-${SUFFIX}`)).toBeDefined();
    expect(findStyle(reparsed, `ClauseHeading1-${SUFFIX}`)).toBeDefined();

    const sourceNumIds = new Set(source.package.numbering?.nums.map((num) => num.numId));
    const addedNums = reparsed.package.numbering?.nums.filter(
      (num) => !sourceNumIds.has(num.numId),
    );
    expect(addedNums?.length).toBeGreaterThan(0);

    const { left, right } = columnParagraphs(reparsed);
    expect(right).toHaveLength(left.length);
    for (const [index, rightParagraph] of right.entries()) {
      const leftParagraph = left[index]!;
      const leftNumId = effectiveNumId(reparsed, leftParagraph);
      const rightNumId = effectiveNumId(reparsed, rightParagraph);
      if (leftNumId === undefined) {
        continue;
      }
      expect(rightNumId).toBeDefined();
      expect(rightNumId).not.toBe(leftNumId);
      expect(abstractNumIdOf(reparsed, rightNumId!)).not.toBe(abstractNumIdOf(reparsed, leftNumId));
      expect(rightParagraph.listRendering?.marker).toBe(leftParagraph.listRendering?.marker);
    }

    // Row handles survive the save: every reported target paraId is a real
    // paragraph in the reparsed right column.
    const rightIds = new Set(right.map((p) => p.paraId));
    for (const row of rows) {
      if (row.kind === "table") {
        continue;
      }
      expect(rightIds.has(row.targetParaId)).toBe(true);
    }
  });

  test("reads the same row manifest back from the saved document", async () => {
    const source = await buildSource();
    const { buffer, rows } = await createBilingualDocx(source.originalBuffer!, {
      targetStyleSuffix: SUFFIX,
    });

    const read = await readBilingualDocx(buffer);
    expect(read).toEqual(rows);
    // Every left paragraph is addressable: the stamp pass gave it a paraId.
    for (const row of read) {
      if (row.kind !== "table") {
        expect(row.sourceParaId).toBeDefined();
      }
    }
    // A plain document has no bilingual table to read.
    expect(await readBilingualDocx(source.originalBuffer!)).toEqual([]);
  });

  test("keeps structural-only paragraphs out of the editable row manifest", async () => {
    const source = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const columnBreak: Paragraph = {
      type: "paragraph",
      content: [{ type: "run", content: [{ type: "break", breakType: "column" }] }],
    };
    source.package.document.content = [
      paragraph("Before break", "Normal"),
      columnBreak,
      paragraph("After break", "Normal"),
    ];

    const { buffer } = await createBilingualDocx(await createDocx(source), {
      targetStyleSuffix: SUFFIX,
    });
    const rows = await readBilingualDocx(buffer);
    const editableIds = new Set(
      (await FolioDocxReviewer.fromBuffer(buffer)).snapshot().blocks.map(({ id }) => id),
    );

    expect(rows).toHaveLength(2);
    for (const row of rows) {
      if (row.kind === "table") {
        continue;
      }
      expect(row.sourceParaId).toBeDefined();
      if (row.sourceParaId !== undefined) {
        expect(editableIds.has(row.sourceParaId)).toBe(true);
      }
      expect(editableIds.has(row.targetParaId)).toBe(true);
    }

    const reparsed = await parseDocx(buffer, { preloadFonts: false });
    expect(reparsed.package.document.content).toHaveLength(3);
    expect(reparsed.package.document.content.at(1)).toMatchObject(columnBreak);
  });

  test("keeps field-only paragraphs out of the editable row manifest", async () => {
    const source = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const fieldOnly: Paragraph = {
      type: "paragraph",
      content: [
        {
          type: "complexField",
          instruction: 'TOC \\o "1-3"',
          fieldType: "TOC",
          fieldCode: [],
          fieldResult: [
            {
              type: "run",
              content: [{ type: "text", text: "Update this field in Word" }],
            },
          ],
        },
      ],
    };
    source.package.document.content = [
      paragraph("Before field", "Normal"),
      fieldOnly,
      paragraph("After field", "Normal"),
    ];

    const { buffer, rows } = await createBilingualDocx(await createDocx(source), {
      targetStyleSuffix: SUFFIX,
    });

    expect(rows).toHaveLength(2);
    const reparsed = await parseDocx(buffer, { preloadFonts: false });
    expect(reparsed.package.document.content).toHaveLength(3);
    expect(reparsed.package.document.content.at(1)).toMatchObject({
      type: "paragraph",
      content: [
        {
          type: "complexField",
          instruction: 'TOC \\o "1-3"',
          fieldResult: [
            {
              type: "run",
              content: [{ type: "text", text: "Update this field in Word" }],
            },
          ],
        },
      ],
    });
  });

  test("keeps a structural-only source table out of the editable row manifest", async () => {
    const source = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    const structuralTable: Table = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            cell([
              {
                type: "paragraph",
                content: [{ type: "run", content: [{ type: "break", breakType: "column" }] }],
              },
            ]),
          ],
        },
      ],
    };
    source.package.document.content = [
      paragraph("Before table"),
      structuralTable,
      paragraph("After table"),
    ];

    const { buffer, rows } = await createBilingualDocx(await createDocx(source), {
      targetStyleSuffix: SUFFIX,
    });

    expect(rows).toHaveLength(2);
    const reparsed = await parseDocx(buffer, { preloadFonts: false });
    expect(reparsed.package.document.content).toHaveLength(3);
    expect(reparsed.package.document.content.at(1)).toMatchObject(structuralTable);
  });

  test("refuses a manifest whose handles are absent from the canonical edit snapshot", async () => {
    const source = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    source.package.document.content = [paragraph("Hidden translation row")];
    const valid = await createBilingualDocx(await createDocx(source), {
      targetStyleSuffix: SUFFIX,
    });
    const malformed = await parseDocx(valid.buffer, { preloadFonts: false });
    const table = bodyTables(malformed).at(0);
    const row = table?.rows.at(0);
    if (!row) {
      throw new Error("bilingual fixture lost its translation row");
    }
    row.formatting = { ...row.formatting, hidden: true };

    await expect(readBilingualDocx(await createDocx(malformed))).rejects.toThrow(
      "Bilingual manifest contains handles absent from the Folio AI-edit snapshot.",
    );
  });

  test("materializes a styles part when the source package has none", async () => {
    const source = await buildSource();
    // Strip word/styles.xml and its relationship: legal OOXML, and the path
    // where minted styles used to be dropped silently on save.
    const zip = await JSZip.loadAsync(source.originalBuffer!);
    zip.remove("word/styles.xml");
    const relsPath = "word/_rels/document.xml.rels";
    const rels = await zip.file(relsPath)!.async("text");
    zip.file(relsPath, rels.replace(/<Relationship [^>]*relationships\/styles"[^>]*\/>/u, ""));
    const stripped = await parseDocx(await zip.generateAsync({ type: "arraybuffer" }), {
      preloadFonts: false,
    });
    expect(stripped.package.styles).toBeUndefined();

    // Give the model a numbered style the body uses, so the copy must mint a
    // clone that only a written styles part can carry.
    const numId = stripped.package.numbering?.nums.at(0)?.numId;
    expect(numId).toBeDefined();
    const styled: Document = {
      ...stripped,
      package: {
        ...stripped.package,
        styles: {
          styles: [
            {
              styleId: "Clause",
              type: "paragraph",
              name: "Clause",
              pPr: { numPr: { numId: numId!, ilvl: 0 } },
            },
          ],
        },
        document: {
          ...stripped.package.document,
          content: stripped.package.document.content.map((block): BlockContent => {
            if (
              block.type !== "paragraph" ||
              block.content.length === 0 ||
              block.sectionProperties
            ) {
              return block;
            }
            const restyled: Paragraph = Object.assign({}, block, {
              formatting: { styleId: "Clause" },
            });
            return restyled;
          }),
        },
      },
    };
    const { document } = createBilingualDocument(styled, await bilingualDocumentOptions(styled));
    expect(findStyle(document, `Clause-${SUFFIX}`)).toBeDefined();

    const reparsed = await parseDocx(await createDocx(document), { preloadFonts: false });
    expect(findStyle(reparsed, `Clause-${SUFFIX}`)).toBeDefined();
    const relsAfter = await (
      await JSZip.loadAsync(await createDocx(document))
    )
      .file(relsPath)!
      .async("text");
    expect(relsAfter).toContain("relationships/styles");
  });

  test("leaves a document without numbering free of clones", async () => {
    const doc = createEmptyDocument({ preset: createStellaStyleDocumentPreset() });
    doc.package.document.content = [
      paragraph("Plain one", "Normal"),
      paragraph("Plain two", "Normal"),
    ];
    const bytes = await createDocx(doc);
    const { buffer, rows, warnings } = await createBilingualDocx(bytes, {
      targetStyleSuffix: SUFFIX,
    });
    const reparsed = await parseDocx(buffer, { preloadFonts: false });

    expect(warnings).toEqual([]);
    expect(rows).toHaveLength(2);
    expect(
      reparsed.package.styles?.styles.some((style) => style.styleId.endsWith(`-${SUFFIX}`)),
    ).toBe(false);
    expect(reparsed.package.numbering?.nums.length).toBe(doc.package.numbering?.nums.length);
  });
});
