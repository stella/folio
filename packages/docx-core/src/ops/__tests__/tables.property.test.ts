/** Direct whole-table laws: exact inverse, determinism, locality and contract closure. */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import type { BlockContent, Document, Paragraph, Table } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import { paragraphIdsIn } from "../ids";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES, type DocumentOp } from "../types";

const NUM_RUNS = 2_000;
const containers = ["body", "cell", "customXml", "sdt", "wrapperCell"] as const;
const shapes = fc.record({
  container: fc.constantFrom(...containers),
  rows: fc.integer({ min: 1, max: 4 }),
  cells: fc.integer({ min: 1, max: 3 }),
  paragraphs: fc.integer({ min: 1, max: 3 }),
  anchors: fc.integer({ min: 1, max: 5 }),
  seed: fc.integer({ min: 0, max: 0xffff_ffff }),
  side: fc.constantFrom("before", "after"),
  inlineReview: fc.constantFrom("none", "insertion", "deletion", "property"),
  rowReview: fc.constantFrom("none", "insertion", "deletion"),
  markPlacement: fc.constantFrom("none", "nonfinal", "final"),
});
type Shape = typeof shapes extends fc.Arbitrary<infer Value> ? Value : never;
const id = (value: number) => value.toString(16).toUpperCase().padStart(8, "0");
const makeParagraph = (value: number, seed: number): Paragraph => ({
  type: "paragraph",
  paraId: id(value),
  formatting: { alignment: seed % 2 === 0 ? "start" : "end" },
  content: [
    {
      type: "run",
      formatting: { bold: seed % 2 === 0 },
      content: [{ type: "text", text: `text-${seed}` }],
    },
  ],
});
const reviewInfo = (value: number) => ({
  id: value,
  author: "Reviewer",
  date: "2026-05-06T07:08:09Z",
});

type ReviewedParagraphOptions = {
  shape: Shape;
  value: number;
  index: number;
  count: number;
};
const reviewedParagraph = ({ shape, value, index, count }: ReviewedParagraphOptions): Paragraph => {
  const source = makeParagraph(value, shape.seed + index);
  let reviewed = source;
  switch (shape.inlineReview) {
    case "none":
      break;
    case "property":
      reviewed = {
        ...source,
        propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: reviewInfo(10_000 + value * 10),
            previousFormatting: { alignment: "center" },
          },
        ],
      };
      break;
    case "insertion":
    case "deletion":
      reviewed = {
        ...source,
        content: [
          {
            type: shape.inlineReview,
            info: reviewInfo(10_000 + value * 10),
            content: [
              { type: "run", content: [{ type: "text", text: `review-${shape.seed}-${index}` }] },
            ],
          },
        ],
      };
      break;
    default: {
      const unreachable: never = shape.inlineReview;
      return unreachable;
    }
  }
  const marked =
    shape.markPlacement === "final"
      ? index === count - 1
      : shape.markPlacement === "nonfinal" && index === 0;
  if (!marked) return reviewed;
  return {
    ...reviewed,
    pPrMark: { kind: shape.seed % 2 === 0 ? "ins" : "del", info: reviewInfo(10_001 + value * 10) },
  };
};

const tableFor = (shape: Shape, review: "plain" | "existing"): Table => ({
  type: "table",
  formatting: { justification: "center" },
  rows: Array.from({ length: shape.rows }, (_rowValue, row) => ({
    type: "tableRow",
    formatting: { cantSplit: row % 2 === 0 },
    ...(review === "existing" && shape.rowReview !== "none"
      ? {
          structuralChange: {
            type:
              shape.rowReview === "insertion"
                ? ("tableRowInsertion" as const)
                : ("tableRowDeletion" as const),
            info: reviewInfo(30_000 + row),
          },
        }
      : {}),
    cells: Array.from({ length: 1 + ((shape.cells + row) % 3) }, (_cellValue, cell) => {
      const count =
        review === "existing" && shape.markPlacement === "nonfinal"
          ? 2 + ((shape.paragraphs + cell) % 2)
          : 1 + ((shape.paragraphs + cell) % 3);
      return {
        type: "tableCell",
        content: Array.from({ length: count }, (_paragraphValue, index) => {
          const value = 0x100 + row * 27 + cell * 9 + index;
          return review === "existing"
            ? reviewedParagraph({ shape, value, index, count })
            : makeParagraph(value, shape.seed + index);
        }),
      };
    }),
  })),
});

const wrap = (shape: Shape, content: BlockContent[]): BlockContent[] => {
  const customXml: BlockContent = {
    type: "blockCustomXml",
    openingXml: '<w:customXml w:element="clause">',
    closingXml: "</w:customXml>",
    content,
  };
  switch (shape.container) {
    case "body":
      return content;
    case "customXml":
      return [customXml, makeParagraph(0x90, shape.seed)];
    case "sdt":
      return [
        { type: "blockSdt", properties: { sdtType: "group", id: 9 }, content },
        makeParagraph(0x90, shape.seed),
      ];
    case "cell":
    case "wrapperCell":
      return [
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content:
                    shape.container === "cell"
                      ? content
                      : [customXml, makeParagraph(0x91, shape.seed)],
                },
              ],
            },
          ],
        },
        makeParagraph(0x90, shape.seed),
      ];
    default: {
      const unreachable: never = shape.container;
      return unreachable;
    }
  }
};
const documentOf = (shape: Shape, content: BlockContent[]): Document => ({
  package: { document: { content: wrap(shape, content) } },
});
const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  return result.value;
};
const families = [
  "insertTable",
  "deleteTable",
  "deleteFinalRow",
  "emptyRows",
  "acceptAllRowDeletions",
  "rejectAllRowInsertions",
] as const;

const fixtureFor = (shape: Shape, family: (typeof families)[number]) => {
  const anchors = Array.from({ length: shape.anchors + 1 }, (_paragraphValue, index) =>
    makeParagraph(index + 1, shape.seed + index),
  );
  const table = tableFor(
    shape,
    family === "acceptAllRowDeletions" || family === "rejectAllRowInsertions"
      ? "plain"
      : "existing",
  );
  const blockId = paragraphIdsIn(table).at(0);
  if (blockId === undefined) throw new Error("Generated table has paragraphs.");
  if (family === "insertTable") {
    const index = shape.seed % anchors.length;
    const anchor = anchors.at(index);
    if (anchor?.paraId === undefined) throw new Error("Generated anchor has an id.");
    // After-insertion anchors exclude the last paragraph: terminal carrier creation is deferred.
    const afterAnchor = anchors.at(shape.seed % (anchors.length - 1));
    if (afterAnchor?.paraId === undefined) throw new Error("Generated after-anchor has an id.");
    return {
      document: documentOf(shape, anchors),
      op: {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story: OP_STORIES.MAIN,
        at: {
          type: shape.side,
          blockId: shape.side === "before" ? anchor.paraId : afterAnchor.paraId,
        },
        table,
      } satisfies DocumentOp,
    };
  }
  const content: BlockContent[] = [...anchors];
  content.splice(shape.seed % anchors.length, 0, table);
  const document = documentOf(shape, content);
  switch (family) {
    case "deleteTable":
      return {
        document,
        op: {
          type: DOCUMENT_OP_TYPES.DELETE_TABLE,
          story: OP_STORIES.MAIN,
          blockId,
          expected: table,
        } satisfies DocumentOp,
      };
    case "deleteFinalRow": {
      const row = table.rows.at(0);
      if (row === undefined) throw new Error("Generated table has a row.");
      const single = { ...table, rows: [row] };
      content.splice(content.indexOf(table), 1, single);
      return {
        document: documentOf(shape, content),
        op: {
          type: DOCUMENT_OP_TYPES.DELETE_ROW,
          story: OP_STORIES.MAIN,
          blockId,
          expected: row,
        } satisfies DocumentOp,
      };
    }
    case "emptyRows":
      return {
        document,
        op: {
          type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
          story: OP_STORIES.MAIN,
          blockId,
          expected: table.rows,
          rows: [],
        } satisfies DocumentOp,
      };
    case "acceptAllRowDeletions":
    case "rejectAllRowInsertions": {
      const marked: Table = {
        ...table,
        rows: table.rows.map((row, index) => ({
          ...row,
          structuralChange: {
            type: family === "acceptAllRowDeletions" ? "tableRowDeletion" : "tableRowInsertion",
            info: { id: 1000 + index, author: "Reviewer", date: "2026-05-06T07:08:09Z" },
          },
        })),
      };
      content.splice(content.indexOf(table), 1, marked);
      return {
        document: documentOf(shape, content),
        op: {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: marked.rows.map((_paragraphValue, index) => 1000 + index),
          decision: family === "acceptAllRowDeletions" ? "accept" : "reject",
        } satisfies DocumentOp,
      };
    }
    default: {
      const unreachable: never = family;
      return unreachable;
    }
  }
};

describe("whole table operation properties", () => {
  for (const family of families) {
    test(
      `${family}: L4 inverse, L5 determinism, L6 locality and L7 contract`,
      () => {
        assertProperty(
          fc.property(
            shapes.map((shape) => ({
              ...shape,
              markPlacement:
                shape.markPlacement === "final" ? ("nonfinal" as const) : shape.markPlacement,
            })),
            (shape) => {
              const { document, op } = fixtureFor(shape, family);
              expect(contractViolation(document)).toBeUndefined();
              const snapshot = structuredClone(document);
              const result = applied(document, op);
              expect(document).toStrictEqual(snapshot);
              const undone = applyDocumentOps(result.document, result.inverse);
              if (undone.isErr()) throw undone.error;
              expect(undone.value.document).toStrictEqual(document);
              expect(contractViolation(undone.value.document)).toBeUndefined();
              const repeated = applied(structuredClone(document), structuredClone(op));
              expect(repeated.document).toStrictEqual(result.document);
              expect(repeated.inverse).toStrictEqual(result.inverse);
              expect(repeated.touched).toStrictEqual(result.touched);
              const touched = new Set([
                ...result.touched.modified,
                ...result.touched.inserted,
                ...result.touched.removed,
              ]);
              const after = new Map(
                storyParagraphs(result.document.package.document).map(({ paragraph }) => [
                  paragraph.paraId,
                  paragraph,
                ]),
              );
              for (const { paragraph } of storyParagraphs(document.package.document)) {
                if (touched.has(paragraph.paraId ?? "")) continue;
                expect(after.get(paragraph.paraId)).toBe(paragraph);
              }
              const beforeIds = paragraphIdsIn(document.package.document);
              const afterIds = paragraphIdsIn(result.document.package.document);
              const removed = beforeIds.filter((value) => !afterIds.includes(value));
              const inserted = afterIds.filter((value) => !beforeIds.includes(value));
              expect(new Set(result.touched.removed)).toEqual(new Set(removed));
              expect(new Set(result.touched.inserted)).toEqual(new Set(inserted));
              if (family !== "insertTable") {
                expect(removed.length).toBeGreaterThan(0);
                expect(afterIds).toEqual(beforeIds.filter((value) => !removed.includes(value)));
              }
            },
          ),
          { numRuns: NUM_RUNS },
        );
      },
      propertyTestTimeout(240_000),
    );
  }

  test(
    "final cell marks refuse whole-table insertion and removal without mutation",
    () => {
      assertProperty(
        fc.property(shapes, (shape) => {
          const markedShape = { ...shape, markPlacement: "final" } as const;
          for (const family of [
            "insertTable",
            "deleteTable",
            "deleteFinalRow",
            "emptyRows",
          ] as const) {
            const { document, op } = fixtureFor(markedShape, family);
            expect(contractViolation(document)).toBeUndefined();
            const snapshot = structuredClone(document);
            const result = applyDocumentOp(document, op);
            if (result.isOk()) throw new Error("Final cell paragraph marks must be refused.");
            expect(result.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK);
            expect(document).toStrictEqual(snapshot);
          }
        }),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(240_000),
  );

  test(
    "unsupported table records refuse insertion and removal without mutation",
    () => {
      assertProperty(
        fc.property(
          shapes,
          fc.constantFrom("invalidText", "missingFinalParagraph", "emptyRow"),
          fc.constantFrom("\u0000", "\u000B", "\uD800", "\uDC00", "\uFFFE", "\uFFFF"),
          (shape, unsupported, illegalText) => {
            // Keep a good first cell so every removal resolves the same outer table;
            // nested malformed rows also remain in the sole row deleted by deleteRow.
            const source = tableFor({ ...shape, rows: 1 }, "plain");
            const row = source.rows.at(0);
            if (row === undefined) throw new Error("Generated table has its sole row.");
            const invalidParagraph: Paragraph = {
              ...makeParagraph(0x70000, shape.seed),
              content: [
                { type: "run", content: [{ type: "text", text: `invalid-${illegalText}` }] },
              ],
            };
            let unsupportedBlocks: BlockContent[];
            switch (unsupported) {
              case "invalidText":
                unsupportedBlocks = [invalidParagraph];
                break;
              case "missingFinalParagraph":
                unsupportedBlocks = [
                  makeParagraph(0x70000, shape.seed),
                  { type: "preservedBlock", xml: "<w:altChunk/>" },
                ];
                break;
              case "emptyRow":
                unsupportedBlocks = [
                  { type: "table", rows: [{ type: "tableRow", cells: [] }] },
                  makeParagraph(0x70000, shape.seed),
                ];
                break;
              default: {
                const unreachable: never = unsupported;
                return unreachable;
              }
            }
            const malformed: Table = {
              ...source,
              rows: [
                {
                  ...row,
                  cells: [...row.cells, { type: "tableCell", content: unsupportedBlocks }],
                },
              ],
            };
            const malformedRow = malformed.rows.at(0);
            if (malformedRow === undefined)
              throw new Error("Malformed table retains a target row.");
            const before = makeParagraph(1, shape.seed);
            const after = makeParagraph(2, shape.seed);
            const insertionDocument = documentOf(shape, [before, after]);
            const removalDocument = documentOf(shape, [before, malformed, after]);
            const insertion: DocumentOp = {
              type: DOCUMENT_OP_TYPES.INSERT_TABLE,
              story: OP_STORIES.MAIN,
              at: { type: "before", blockId: id(1) },
              table: malformed,
            };
            const removals = [
              {
                type: DOCUMENT_OP_TYPES.DELETE_TABLE,
                story: OP_STORIES.MAIN,
                blockId: id(0x100),
                expected: malformed,
              },
              {
                type: DOCUMENT_OP_TYPES.DELETE_ROW,
                story: OP_STORIES.MAIN,
                blockId: id(0x100),
                expected: malformedRow,
              },
              {
                type: DOCUMENT_OP_TYPES.SET_TABLE_ROWS,
                story: OP_STORIES.MAIN,
                blockId: id(0x100),
                expected: malformed.rows,
                rows: [],
              },
            ] as const satisfies readonly DocumentOp[];
            const cases = [
              { document: insertionDocument, op: insertion },
              ...removals.map((op) => ({ document: removalDocument, op })),
            ];
            for (const { document, op } of cases) {
              expect(contractViolation(document)).toBeUndefined();
              const snapshot = structuredClone(document);
              const result = applyDocumentOp(document, op);
              if (result.isOk()) throw new Error("Unsupported table records must be refused.");
              expect(result.error.reason).toBe(
                unsupported === "invalidText"
                  ? DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT
                  : DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
              );
              expect(document).toStrictEqual(snapshot);
            }
          },
        ),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(240_000),
  );
});
