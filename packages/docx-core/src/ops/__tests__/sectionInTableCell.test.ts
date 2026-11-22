/**
 * A section ends at a paragraph outside every table. No operation puts one on
 * a paragraph in a cell, whether it makes, moves, copies or restores it; one
 * a document arrived with stays where it stands. The answer depends on the
 * document's content and the operation alone.
 */

import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
  TableRow,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { hasSectionInTableCell } from "../cellSections";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  SECTION_BOUNDARY_POLICIES,
} from "../types";

const SECTION = { pageWidth: 11_906 };

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const ending = (source: Paragraph): Paragraph => ({
  ...source,
  sectionProperties: SECTION,
});

const rowOf = (content: BlockContent[]): TableRow => ({
  type: "tableRow",
  cells: [{ type: "tableCell", content }],
});

const tableOf = (...rows: TableRow[]): Table => ({
  type: "table",
  columnWidths: [2400],
  rows,
});

const inner = paragraph("00000003", "inner");
const beside = paragraph("00000004", "beside");

/** Where the edited paragraph stands: each is inside a table cell. */
const CELLS: Record<string, BlockContent[]> = {
  "a cell": [inner, beside],
  "a control in a cell": [
    {
      type: "blockSdt",
      properties: { sdtType: "richText", id: 41 },
      content: [inner, beside],
    },
    paragraph("00000005", "end"),
  ],
  "a table in a cell": [
    tableOf(rowOf([inner, beside])),
    paragraph("00000005", "end"),
  ],
};

const documentOf = (blocks: BlockContent[]): Document => {
  const finalSectionProperties = { pageWidth: 12_240 };
  return {
    package: {
      document: {
        content: blocks,
        sections: [{ properties: finalSectionProperties, content: blocks }],
        finalSectionProperties,
      },
    },
  };
};

const cellDocument = (cell: BlockContent[]): Document =>
  structuredClone(
    documentOf([
      paragraph("00000001", "before"),
      tableOf(rowOf(cell)),
      paragraph("00000009", "after"),
    ]),
  );

const bodyDocument = (): Document =>
  structuredClone(documentOf([inner, beside, paragraph("00000009", "after")]));

const story = OP_STORIES.MAIN;

const replaced: DocumentOp = {
  type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
  story,
  expected: [inner],
  blocks: [ending(inner)],
  sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
};
const endpoint: DocumentOp = {
  type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
  endpoint: { type: "paragraph", blockId: "00000003" },
  properties: { type: "present", value: SECTION },
};
const split: DocumentOp = {
  type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
  at: { story, blockId: "00000003", offset: 2 },
  newBlockId: "0000000A",
  newHalf: "first",
  firstSectionProperties: SECTION,
};

const refused = (document: Document, op: DocumentOp): void => {
  const before = structuredClone(document);
  const result = applyDocumentOp(document, op);
  expect(result.isErr() ? result.error.reason : "applied").toBe(
    DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
  );
  assertExactModel(document, before);
};

/** The inverse as a journal carries it. */
const journaled = (ops: readonly DocumentOp[]): DocumentOp[] =>
  JSON.parse(JSON.stringify(ops));

describe("an operation that would end a section in a table cell", () => {
  for (const [place, cell] of Object.entries(CELLS)) {
    test(`replacing a paragraph of ${place} is refused`, () => {
      refused(cellDocument(cell), replaced);
    });

    test(`setting the section endpoint of ${place} is refused`, () => {
      refused(cellDocument(cell), endpoint);
    });

    test(`splitting a paragraph of ${place} is refused`, () => {
      refused(cellDocument(cell), split);
    });
  }

  test.each([
    ["replacing a paragraph", replaced],
    ["setting the section endpoint", endpoint],
    ["splitting a paragraph", split],
  ] as const)("outside every table, %s applies", (_title, op) => {
    const document = bodyDocument();
    const before = structuredClone(document);
    const applied = applyDocumentOp(document, op).unwrap();
    expect(
      storyParagraphs(applied.document.package.document).filter(
        ({ paragraph: item }) => item.sectionProperties !== undefined,
      ),
    ).toHaveLength(1);
    const undone = applyDocumentOps(
      applied.document,
      journaled(applied.inverse),
    ).unwrap();
    assertExactModel(undone.document, before);
  });

  const rows = [rowOf([inner, beside])];
  const endingRow = rowOf([ending(inner), beside]);
  const table = tableOf(...rows);

  test("setting the rows of a table is refused", () => {
    refused(cellDocument([inner, beside]), {
      type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
      story,
      blockId: "00000003",
      expected: rows,
      rows: [endingRow],
    });
  });

  test("setting a table is refused", () => {
    refused(cellDocument([inner, beside]), {
      type: DOCUMENT_OP_TYPES.SET_TABLE,
      story,
      blockId: "00000003",
      expected: table,
      table: tableOf(endingRow),
    });
  });

  test("inserting a row is refused", () => {
    refused(cellDocument([inner, beside]), {
      type: DOCUMENT_OP_TYPES.INSERT_ROW,
      story,
      blockId: "00000003",
      at: 1,
      row: rowOf([ending(paragraph("0000000A", "row"))]),
    });
  });

  test("inserting a table is refused", () => {
    refused(bodyDocument(), {
      type: DOCUMENT_OP_TYPES.INSERT_TABLE,
      story,
      at: { type: "after", blockId: "00000003" },
      table: tableOf(rowOf([ending(paragraph("0000000A", "table"))])),
    });
  });

  test("a new header holding one is refused", () => {
    refused(bodyDocument(), {
      type: DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER,
      sectionIndex: 0,
      story: { kind: "header", rId: "rIdHeader" },
      referenceType: "default",
      content: [
        tableOf(rowOf([ending(paragraph("0000000A", "header"))])),
        paragraph("0000000B", "end"),
      ],
    });
  });

  test("drawing a text box that holds one is refused", () => {
    refused(bodyDocument(), {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: { story, blockId: "00000003", offset: 0 },
      slice: {
        openStart: 0,
        openEnd: 0,
        content: [
          {
            type: "run",
            content: [
              {
                type: "shape",
                shape: {
                  type: "shape",
                  shapeType: "textBox",
                  id: "1",
                  size: { width: 2_000_000, height: 800_000 },
                  textBody: {
                    content: [
                      tableOf(rowOf([ending(paragraph("0000000A", "box"))])),
                      paragraph("0000000B", "end"),
                    ],
                  },
                },
              },
            ],
          },
        ],
      },
    });
  });

  test("restoring a body that holds one is refused", () => {
    const document = bodyDocument();
    const { content } = document.package.document;
    refused(document, {
      type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
      expected: { body: { content } },
      parts: {
        body: {
          content: [
            ...content,
            tableOf(rowOf([ending(paragraph("0000000A", "table"))])),
            paragraph("0000000B", "end"),
          ],
        },
      },
    });
  });
});

/** A row of two cells, the first holding the section, and rows around it. */
const ROWS: readonly TableRow[] = [
  rowOf([paragraph("00000010", "first")]),
  {
    type: "tableRow",
    cells: [
      { type: "tableCell", content: [ending(inner), beside] },
      { type: "tableCell", content: [paragraph("00000011", "other")] },
    ],
  },
  rowOf([paragraph("00000012", "last")]),
];

const arrived = (): Document =>
  structuredClone(
    documentOf([
      paragraph("00000001", "before"),
      { type: "table", columnWidths: [2400, 2400], rows: [...ROWS] },
      paragraph("00000009", "after"),
    ]),
  );

const rowsSet = (rows: readonly TableRow[]): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
  story,
  blockId: "00000003",
  expected: ROWS,
  rows,
});

const rowWith = (...cells: BlockContent[][]): TableRow => ({
  type: "tableRow",
  cells: cells.map((content) => ({ type: "tableCell", content })),
});

const [firstRow, sectionRow, lastRow] = ROWS;
if (firstRow === undefined || sectionRow === undefined || lastRow === undefined)
  panic("The fixture has its three rows.");

const typed = (blockId: string): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_TEXT,
  at: { story, blockId, offset: 0 },
  text: "x",
  runProps: {},
});

const removal: DocumentOp = {
  type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
  story,
  expected: [ending(inner)],
  blocks: [inner],
  sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
};

const LEAVES: Record<string, DocumentOp> = {
  "typing in its paragraph": typed("00000003"),
  "typing beside it": typed("00000004"),
  "typing outside the table": typed("00000009"),
  "changing the section's own properties": {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story,
    expected: [ending(inner)],
    blocks: [{ ...inner, sectionProperties: { pageWidth: 16_838 } }],
    sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
  },
  "splitting its paragraph before its mark": {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: { story, blockId: "00000003", offset: 2 },
    newBlockId: "0000000A",
  },
  "inserting a paragraph before the table": {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story,
    at: { type: "before", blockId: "00000001" },
    blocks: [paragraph("0000000A", "new")],
  },
  "inserting a row above": {
    type: DOCUMENT_OP_TYPES.INSERT_ROW,
    story,
    blockId: "00000003",
    at: 0,
    row: rowOf([paragraph("0000000A", "new")]),
  },
  "deleting the row above": {
    type: DOCUMENT_OP_TYPES.DELETE_ROW,
    story,
    blockId: "00000010",
  },
  "splitting the paragraph beside it": {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: { story, blockId: "00000004", offset: 2 },
    newBlockId: "0000000A",
  },
  "merging another cell into its cell": rowsSet([
    firstRow,
    rowWith([ending(inner), beside, paragraph("00000011", "other")]),
    lastRow,
  ]),
};

const GAINS: Record<string, DocumentOp> = {
  "moving it to another cell": rowsSet([
    firstRow,
    rowWith([beside], [ending(inner), paragraph("00000011", "other")]),
    lastRow,
  ]),
  "merging its cell into another": rowsSet([
    firstRow,
    rowWith([paragraph("00000011", "other"), ending(inner), beside]),
    lastRow,
  ]),
  "moving its row": rowsSet([sectionRow, firstRow, lastRow]),
  "moving it past the paragraph beside it": rowsSet([
    firstRow,
    rowWith([beside, ending(inner)], [paragraph("00000011", "other")]),
    lastRow,
  ]),
  "moving it to the paragraph beside it": rowsSet([
    firstRow,
    rowWith([inner, ending(beside)], [paragraph("00000011", "other")]),
    lastRow,
  ]),
  "moving it to a new paragraph": {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: { story, blockId: "00000003", offset: 5 },
    newBlockId: "0000000A",
  },
  "copying it into its cell": rowsSet([
    firstRow,
    rowWith(
      [ending(inner), beside, ending(paragraph("0000000A", "inner"))],
      [paragraph("00000011", "other")],
    ),
    lastRow,
  ]),
  "copying it into another cell": {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story,
    expected: [paragraph("00000011", "other")],
    blocks: [ending(paragraph("00000011", "other"))],
    sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
  },
};

/** The same content as a fresh value, as JSON reads it back, and after an edit that was undone. */
const sameContent = (document: Document): Document[] => {
  const edited = applyDocumentOp(document, typed("00000009")).unwrap();
  return [
    document,
    structuredClone(document),
    JSON.parse(JSON.stringify(document)),
    applyDocumentOps(edited.document, edited.inverse).unwrap().document,
  ];
};

describe("a document that arrived with a section in a table cell", () => {
  for (const [title, op] of Object.entries(LEAVES))
    test(`${title} applies and is undone exactly`, () => {
      for (const document of sameContent(arrived())) {
        const before = structuredClone(document);
        const applied = applyDocumentOp(document, op).unwrap();
        expect(hasSectionInTableCell(applied.document)).toBe(true);
        const undone = applyDocumentOps(
          applied.document,
          journaled(applied.inverse),
        ).unwrap();
        assertExactModel(undone.document, before);
      }
    });

  for (const [title, op] of Object.entries(GAINS))
    test(`${title} is refused`, () => {
      for (const document of sameContent(arrived())) refused(document, op);
    });

  test("deleting its table is refused", () => {
    for (const document of sameContent(arrived()))
      refused(document, {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story,
        blockId: "00000003",
      });
  });

  test("an edit that removes it applies, and its undo is refused", () => {
    for (const document of sameContent(arrived())) {
      const removed = applyDocumentOp(document, removal).unwrap();
      expect(hasSectionInTableCell(removed.document)).toBe(false);
      const [undo] = journaled(removed.inverse);
      if (undo === undefined) panic("The removal has an inverse.");
      for (const without of sameContent(removed.document))
        refused(without, undo);
    }
  });

  test("a removal that also refers to a kept note applies", () => {
    const document = arrived();
    document.package.footnotes = [
      { type: "footnote", id: 2, content: [paragraph("000000F0", "note")] },
    ];
    const noted: Paragraph = {
      ...inner,
      content: [{ type: "run", content: [{ type: "footnoteRef", id: 2 }] }],
    };
    const applied = applyDocumentOp(document, {
      ...removal,
      blocks: [noted],
    }).unwrap();
    expect(hasSectionInTableCell(applied.document)).toBe(false);
    expect(applied.document.package.footnotes).toHaveLength(1);
  });
});
