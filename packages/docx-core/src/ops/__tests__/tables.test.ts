import { describe, expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph, Table } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES, type DocumentOp } from "../types";

const paragraph = (paraId: string, text = "kept"): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const table = (paraId = "00000010"): Table => ({
  type: "table",
  rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [paragraph(paraId)] }] }],
});
const documentOf = (content: BlockContent[]): Document => ({ package: { document: { content } } });
const insertion = (value: Table): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_TABLE,
  story: OP_STORIES.MAIN,
  at: { type: "before", blockId: "00000001" },
  table: value,
});
const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
  return result.value;
};
const refused = (document: Document, op: DocumentOp, reason: string) => {
  const before = structuredClone(document);
  const result = applyDocumentOp(document, op);
  if (result.isOk()) throw new Error(`Expected ${reason} refusal.`);
  expect(result.error.reason).toBe(reason);
  expect(document).toStrictEqual(before);
};

describe("whole table operations", () => {
  test("insertion and deletion retain outside paragraph identities and exact inverses", () => {
    const anchor = paragraph("00000001");
    const document = documentOf([anchor]);
    const inserted = table();
    const result = applied(document, insertion(inserted));
    expect(result.document.package.document.content).toStrictEqual([inserted, anchor]);
    expect(result.document.package.document.content.at(-1)).toBe(anchor);
    expect(result.touched).toEqual({ modified: [], inserted: ["00000010"], removed: [] });
    const deleted = applied(result.document, {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      expected: inserted,
    });
    expect(deleted.document).toStrictEqual(document);
    expect(deleted.touched).toEqual({ modified: [], inserted: [], removed: ["00000010"] });
  });

  test("deletion selects the innermost table and preserves its outer siblings", () => {
    const inner = table();
    const sibling = table("00000020");
    const carrier = paragraph("00000030");
    const outer: Table = {
      type: "table",
      rows: [
        { type: "tableRow", cells: [{ type: "tableCell", content: [inner, sibling, carrier] }] },
      ],
    };
    const document = documentOf([outer, paragraph("00000001")]);
    const result = applied(document, {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      expected: inner,
    });
    const nextOuter = result.document.package.document.content.at(0);
    if (nextOuter?.type !== "table") throw new Error("Outer table remains present.");
    const blocks = nextOuter.rows.at(0)?.cells.at(0)?.content;
    expect(blocks).toEqual([sibling, carrier]);
    expect(blocks?.at(0)).toBe(sibling);
    expect(blocks?.at(1)).toBe(carrier);
  });

  test("table and container inverse preconditions reject stale snapshots", () => {
    const original = table();
    const document = documentOf([original, paragraph("00000001")]);
    refused(
      document,
      {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "00000010",
        expected: { ...original, formatting: { justification: "center" } },
      },
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
    );
    const inserted = applied(documentOf([paragraph("00000001")]), insertion(original));
    const inverse = inserted.inverse.at(0);
    if (inverse === undefined) throw new Error("Insertion has an inverse.");
    const changed = documentOf([original, paragraph("00000001", "changed")]);
    refused(changed, inverse, DOCUMENT_OP_REFUSAL_REASONS.STALE);
  });

  test.each(["body", "cell"] as const)(
    "insertion after a final %s paragraph is refused",
    (kind) => {
      const final = paragraph("00000001");
      const document = documentOf(
        kind === "body"
          ? [final]
          : [
              {
                type: "table",
                rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [final] }] }],
              },
              paragraph("00000002"),
            ],
      );
      refused(
        document,
        {
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          story: OP_STORIES.MAIN,
          at: { type: "after", blockId: "00000001" },
          table: table(),
        },
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      );
    },
  );

  test("deletion without a surviving paragraph in its block list is untrackable", () => {
    const original = table();
    refused(
      documentOf([original]),
      {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "00000010",
      },
      DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    );
  });

  test("missing and non-table targets are refused without mutation", () => {
    const document = documentOf([paragraph("00000001")]);
    refused(
      document,
      {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "00000002",
      },
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
    );
    refused(
      document,
      {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
      },
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
  });

  test("insertion rejects identity collisions across the package and within the table", () => {
    const document = documentOf([paragraph("00000001")]);
    for (const value of [
      table("00000001"),
      { ...table(), rows: [...table().rows, ...table().rows] },
    ]) {
      refused(document, insertion(value), DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION);
    }
  });

  test("insertion rejects revision identity collisions in other package paragraphs", () => {
    const info = { id: 50, author: "Reviewer" };
    const marked = (paraId: string): Paragraph => ({
      ...paragraph(paraId),
      content: [
        {
          type: "insertion",
          info,
          content: [{ type: "run", content: [{ type: "text", text: "tracked" }] }],
        },
      ],
    });
    const document = documentOf([marked("00000001")]);
    const value: Table = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [marked("00000010")],
            },
          ],
        },
      ],
    };
    refused(document, insertion(value), DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION);
  });

  test("deletion preserves section-bearing cell paragraphs", () => {
    const value: Table = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [{ ...paragraph("00000010"), sectionProperties: {} }],
            },
          ],
        },
      ],
    };
    const document = documentOf([value, paragraph("00000001")]);
    refused(
      document,
      {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "00000010",
      },
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
    );
  });

  test("malformed table rows and cells are refused", () => {
    const document = documentOf([paragraph("00000001")]);
    const malformed: Table[] = [
      {
        type: "table",
        rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [table("00000020")] }] }],
      },
      { type: "table", rows: [] },
      { type: "table", rows: [{ type: "tableRow", cells: [] }] },
      { type: "table", rows: [{ type: "tableRow", cells: [{ type: "tableCell", content: [] }] }] },
    ];
    for (const value of malformed)
      refused(document, insertion(value), DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
  });

  test("insertion rejects missing paragraph ids, section boundaries and XML-invalid text", () => {
    const document = documentOf([paragraph("00000001")]);
    const withParagraph = (value: Paragraph): Table => ({
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [{ type: "tableCell", content: [value] }],
        },
      ],
    });
    for (const value of [undefined, "00000000", "80000000", "short"]) {
      refused(
        document,
        insertion(withParagraph({ ...paragraph("00000010"), paraId: value })),
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      );
    }
    refused(
      document,
      insertion(withParagraph({ ...paragraph("00000010"), sectionProperties: {} })),
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
    );
    for (const text of ["bad\u0000text", "bad\u000Btext", "bad\uD800text", "bad\uFFFEtext"]) {
      refused(
        document,
        insertion(withParagraph(paragraph("00000010", text))),
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      );
    }
  });
});
