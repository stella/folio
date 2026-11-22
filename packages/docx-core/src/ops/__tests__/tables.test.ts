import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES } from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

const paragraph = (paraId: string, text = "kept"): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const table = (paraId = "00000010"): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [paragraph(paraId)] }],
    },
  ],
});
const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});
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
  test("tables placed beside tables preserve topology, review outcomes and wire history", () => {
    assertProperty(
      fc.property(
        fc.constantFrom("beforeTable", "afterTable"),
        fc.constantFrom("start", "end", "touching"),
        fc.boolean(),
        fc.integer({ min: 1, max: 3 }),
        (type, position, nested, rowCount) => {
          const target = table("00000010");
          const carrier = paragraph("00000001");
          const adjacent = table("00000020");
          let siblings: BlockContent[];
          if (position === "start") siblings = [target, carrier];
          else if (position === "end") siblings = [carrier, target];
          else siblings = [carrier, adjacent, target, table("00000030")];
          // A cell ends in a paragraph; a document body may end in a table.
          if (nested && position !== "start")
            siblings.push(paragraph("00000040"));
          const baseline = documentOf(
            nested
              ? [
                  paragraph("00000002"),
                  {
                    type: "table",
                    rows: [
                      {
                        type: "tableRow",
                        cells: [{ type: "tableCell", content: siblings }],
                      },
                    ],
                  },
                ]
              : siblings,
          );
          const inserted: Table = {
            type: "table",
            formatting: { layout: "fixed", justification: "center" },
            columnWidths: [1200],
            rows: Array.from({ length: rowCount }, (_, index) => ({
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [paragraph(`0000010${index}`, `Inserted ${index}`)],
                },
              ],
            })),
          };
          const expected = structuredClone(baseline);
          const outer = expected.package.document.content.at(1);
          const expectedSiblings =
            nested && outer?.type === "table"
              ? outer.rows.at(0)?.cells.at(0)?.content
              : expected.package.document.content;
          if (!expectedSiblings)
            throw new Error("The target block list exists.");
          const targetIndex = siblings.indexOf(target);
          expectedSiblings.splice(
            targetIndex + (type === "afterTable" ? 1 : 0),
            0,
            inserted,
          );
          const operation = {
            type: DOCUMENT_OP_TYPES.INSERT_TABLE,
            story: OP_STORIES.MAIN,
            at: { type, blockId: "00000010" },
            table: inserted,
          } as const satisfies DocumentOp;
          const direct = applied(
            baseline,
            JSON.parse(JSON.stringify(operation)),
          );
          expect(direct.document).toStrictEqual(expected);
          const tracked = applied(baseline, {
            ...operation,
            revision: { id: 300, author: "Reviewer" },
            newIds: {
              revision: Array.from({ length: 32 }, (_, index) => 301 + index),
            },
          });
          for (const decision of ["accept", "reject"] as const) {
            const resolved = applied(tracked.document, {
              type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
              story: OP_STORIES.MAIN,
              revisionIds: tracked.revisions ?? [],
              decision,
            });
            expect(resolved.document).toStrictEqual(
              decision === "accept" ? expected : baseline,
            );
            const undone = applyDocumentOps(
              resolved.document,
              JSON.parse(JSON.stringify(resolved.inverse)),
            ).unwrap();
            expect(undone.document).toStrictEqual(tracked.document);
            const redone = applyDocumentOps(
              undone.document,
              JSON.parse(JSON.stringify(undone.inverse)),
            ).unwrap();
            expect(redone.document).toStrictEqual(resolved.document);
          }
        },
      ),
      {
        id: "tables placed beside tables preserve topology, review outcomes and wire history",
        numRuns: 60,
      },
    );
  });

  test("insertion and deletion retain outside paragraph identities and exact inverses", () => {
    const anchor = paragraph("00000001");
    const document = documentOf([anchor]);
    const inserted = table();
    const result = applied(document, insertion(inserted));
    expect(result.document.package.document.content).toStrictEqual([
      inserted,
      anchor,
    ]);
    expect(result.document.package.document.content.at(-1)).toBe(anchor);
    expect(result.touched).toEqual({
      modified: [],
      inserted: ["00000010"],
      removed: [],
    });
    const deleted = applied(result.document, {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      expected: inserted,
    });
    expect(deleted.document).toStrictEqual(document);
    expect(deleted.touched).toEqual({
      modified: [],
      inserted: [],
      removed: ["00000010"],
    });
  });

  test("deletion selects the innermost table and preserves its outer siblings", () => {
    const inner = table();
    const sibling = table("00000020");
    const carrier = paragraph("00000030");
    const outer: Table = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [{ type: "tableCell", content: [inner, sibling, carrier] }],
        },
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
    if (nextOuter?.type !== "table")
      throw new Error("Outer table remains present.");
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
    const inserted = applied(
      documentOf([paragraph("00000001")]),
      insertion(original),
    );
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
                rows: [
                  {
                    type: "tableRow",
                    cells: [{ type: "tableCell", content: [final] }],
                  },
                ],
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
      refused(
        document,
        insertion(value),
        DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      );
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
          content: [
            { type: "run", content: [{ type: "text", text: "tracked" }] },
          ],
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
    refused(
      document,
      insertion(value),
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
    );
  });

  test("malformed table rows and cells are refused", () => {
    const document = documentOf([paragraph("00000001")]);
    const malformed: Table[] = [
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [{ type: "tableCell", content: [table("00000020")] }],
          },
        ],
      },
      { type: "table", rows: [] },
      { type: "table", rows: [{ type: "tableRow", cells: [] }] },
      {
        type: "table",
        rows: [
          { type: "tableRow", cells: [{ type: "tableCell", content: [] }] },
        ],
      },
    ];
    for (const value of malformed)
      refused(
        document,
        insertion(value),
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      );
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
      insertion(
        withParagraph({ ...paragraph("00000010"), sectionProperties: {} }),
      ),
      DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
    );
    for (const text of [
      "bad\u0000text",
      "bad\u000Btext",
      "bad\uD800text",
      "bad\uFFFEtext",
    ]) {
      refused(
        document,
        insertion(withParagraph(paragraph("00000010", text))),
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
      );
    }
  });
});
