/**
 * A section ends at a paragraph outside every table. No operation puts one on
 * a paragraph in a cell; a document that entered with one keeps its edits and
 * their exact inverses.
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
import { hasSectionInTableCell } from "../contract";
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

describe("a document that entered with a section in a table cell", () => {
  const entered = (): Document => cellDocument([ending(inner), beside]);

  test("keeps it through an edit beside it, undone exactly", () => {
    const document = entered();
    const before = structuredClone(document);
    const applied = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story, blockId: "00000004", offset: 0 },
      text: "x",
      runProps: {},
    }).unwrap();
    expect(hasSectionInTableCell(applied.document)).toBe(true);
    const undone = applyDocumentOps(
      applied.document,
      journaled(applied.inverse),
    ).unwrap();
    assertExactModel(undone.document, before);
  });

  test("loses it in an edit, and has it back when that is undone", () => {
    const document = entered();
    const before = structuredClone(document);
    const removed = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story,
      expected: [ending(inner)],
      blocks: [inner],
      sectionBoundaries: SECTION_BOUNDARY_POLICIES.REPLACE,
    }).unwrap();
    expect(hasSectionInTableCell(removed.document)).toBe(false);
    const undone = applyDocumentOps(
      removed.document,
      journaled(removed.inverse),
    ).unwrap();
    assertExactModel(undone.document, before);
    const [first] = removed.inverse;
    if (first === undefined) panic("The removal has an inverse.");
    // The same content entered without one: the same operation is refused.
    refused(structuredClone(removed.document), first);
  });
});
