import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { validateDocumentModel } from "../../validate/docx";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { deleteBetween } from "../inline";
import { leafPlaces } from "../leaves";
import { rangeBoundaryDefinitions } from "../rangeBoundaries";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const paragraph = (paraId: string, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});

const documentOf = (...content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const position = (blockId: string, offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

type RangeDefinitions = ReturnType<typeof rangeBoundaryDefinitions>;
type BalancedFamily = Exclude<
  RangeDefinitions[keyof RangeDefinitions]["family"],
  "bookmark"
>;

/** Every family whose unmatched endpoints make the document invalid. */
const FAMILIES = {
  comment: {
    start: { type: "commentRangeStart", id: 1 },
    end: { type: "commentRangeEnd", id: 1 },
  },
  moveFrom: {
    start: { type: "moveFromRangeStart", id: 1, name: "_Move1", author: "A" },
    end: { type: "moveFromRangeEnd", id: 1 },
  },
  moveTo: {
    start: { type: "moveToRangeStart", id: 1, name: "_Move1", author: "A" },
    end: { type: "moveToRangeEnd", id: 1 },
  },
} as const satisfies Record<
  BalancedFamily,
  { start: ParagraphContent; end: ParagraphContent }
>;

type DeleteInOptions = {
  document: Document;
  blockId: string;
  from: number;
  to: number;
};
const deleteIn = ({ document, blockId, from, to }: DeleteInOptions) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: position(blockId, from),
    to: position(blockId, to),
  });
  if (applied.isErr()) throw new Error(applied.error.message);
  return applied.value;
};

type NestedBoundaryOptions = {
  node: ParagraphContent;
  depth: number;
  id: number;
};
const nestedBoundary = ({
  node,
  depth,
  id,
}: NestedBoundaryOptions): ParagraphContent => {
  let nested = node;
  for (let level = 0; level < depth; level++)
    nested = {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: id + level, tag: "range-owner" },
      content: [
        {
          type: "hyperlink",
          rId: "rId9",
          href: "https://example.org",
          children: [
            {
              type: "inlineWrapper",
              kind: "bidi",
              control: "embedding",
              content: [nested],
            },
          ],
        },
      ],
    };
  return nested;
};

const contentOf = (document: Document, index: number): ParagraphContent[] => {
  const block = document.package.document.content.at(index);
  if (block?.type !== "paragraph") throw new Error("expected a paragraph");
  return block.content;
};

const undone = (applied: ReturnType<typeof deleteIn>): Document => {
  const restored = applyDocumentOps(
    applied.document,
    JSON.parse(JSON.stringify(applied.inverse)),
  );
  if (restored.isErr()) throw new Error(restored.error.message);
  return restored.value.document;
};

describe("a direct deletion keeps comment and move ranges well formed", () => {
  for (const [family, { start, end }] of Object.entries(FAMILIES)) {
    test(`all ordered ${family} pair cuts retain precisely the partially covered pair`, () => {
      const document = documentOf(
        paragraph("00000001", [
          run("a"),
          nestedBoundary({ node: start, depth: 1, id: 10 }),
          run("b"),
          nestedBoundary({ node: end, depth: 1, id: 20 }),
          run("c"),
          nestedBoundary({ node: structuredClone(start), depth: 1, id: 30 }),
          run("d"),
          nestedBoundary({ node: end, depth: 1, id: 40 }),
          run("e"),
        ]),
      );
      const source = document.package.document.content.at(0);
      if (source?.type !== "paragraph")
        throw new Error("The fixture has its paragraph.");
      const leaves = leafPlaces(source.content);
      const last = leaves.at(-1);
      if (!last) throw new Error("The fixture has leaves.");
      const gaps = [...leaves.map(({ before }) => before), last.after];
      const exercised = new Set<string>();
      for (const [first, from] of gaps.entries())
        for (const [final, to] of gaps.entries()) {
          if (from.offset >= to.offset || first >= final) continue;
          const covered = [1, 3, 5, 7].map(
            (index) => first <= index && index < final,
          );
          exercised.add(covered.map(Number).join(""));
          const expected = [
            [covered.at(0), covered.at(1)],
            [covered.at(2), covered.at(3)],
          ].flatMap(([a, b]) => (a && b ? [] : [start.type, end.type]));
          const direct = applyDocumentOp(document, {
            type: DOCUMENT_OP_TYPES.DELETE_RANGE,
            from: {
              ...position("00000001", from.offset),
              zeroWidthBefore: from.zeroWidthBefore,
            },
            to: {
              ...position("00000001", to.offset),
              zeroWidthBefore: to.zeroWidthBefore,
            },
          });
          if (direct.isErr()) throw direct.error;
          const block = direct.value.document.package.document.content.at(0);
          if (block?.type !== "paragraph")
            throw new Error("The deletion keeps its paragraph.");
          expect(
            leafPlaces(block.content).flatMap(({ node }) =>
              node.type === start.type || node.type === end.type
                ? [node.type]
                : [],
            ),
          ).toEqual(expected);
          const tracked = applyDocumentOp(document, {
            type: DOCUMENT_OP_TYPES.DELETE_RANGE,
            from: {
              ...position("00000001", from.offset),
              zeroWidthBefore: from.zeroWidthBefore,
            },
            to: {
              ...position("00000001", to.offset),
              zeroWidthBefore: to.zeroWidthBefore,
            },
            revision: { id: 500, author: "Reviewer" },
            newIds: {
              control: Array.from({ length: 40 }, (_, index) => 1000 + index),
              revision: Array.from({ length: 40 }, (_, index) => 2000 + index),
            },
          });
          if (tracked.isErr()) throw tracked.error;
          const accepted = applyDocumentOp(tracked.value.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: tracked.value.revisions,
            decision: "accept",
          });
          if (accepted.isErr()) throw accepted.error;
          expect(accepted.value.document).toStrictEqual(direct.value.document);
          expect(undone(direct.value)).toStrictEqual(document);
        }
      expect([...exercised].sort()).toEqual(
        [
          "0000",
          "1000",
          "0100",
          "0010",
          "0001",
          "1100",
          "0110",
          "0011",
          "1110",
          "0111",
          "1111",
        ].sort(),
      );
    });

    test(`a ${family} deletion removes a complete pair beside another pair's covered endpoint`, () => {
      const document = documentOf(
        paragraph("00000001", [
          run("a"),
          start,
          run("b"),
          end,
          run("c"),
          structuredClone(start),
          run("dX"),
          end,
          run("e"),
        ]),
      );
      if (family === "comment")
        document.package.document.comments = [
          { id: 1, author: "A", content: [paragraph("7FFFFFF0", [])] },
        ];
      expect(validateDocumentModel(document).valid).toBe(true);
      const result = deleteIn({
        document: document,
        blockId: "00000001",
        from: 0,
        to: 4,
      });
      expect(contentOf(result.document, 0)).toEqual([
        start,
        run("X"),
        end,
        run("e"),
      ]);
      expect(undone(result)).toStrictEqual(document);
    });

    test(
      `a ${family} deletion retains source ancestry through acceptance, rejection and JSON inverses`,
      async () => {
        await assertProperty(
          fc.property(
            fc.constantFrom("start", "end", "whole"),
            // Bare comment anchors cannot be wrapped; records holding them can.
            fc.integer({ min: family === "comment" ? 1 : 0, max: 3 }),
            (cut, depth) => {
              const document = documentOf(
                paragraph("00000001", [
                  run("ab"),
                  nestedBoundary({ node: start, depth: depth, id: 10 }),
                  run("cd"),
                  nestedBoundary({ node: end, depth: depth, id: 20 }),
                  run("ef"),
                ]),
              );
              const from = cut === "end" ? 3 : 1;
              const to = cut === "start" ? 3 : 5;
              const direct = deleteIn({
                document: document,
                blockId: "00000001",
                from: from,
                to: to,
              });
              const tracked = applyDocumentOp(document, {
                type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                from: position("00000001", from),
                to: position("00000001", to),
                revision: {
                  id: 500,
                  author: "Reviewer",
                  date: "2026-11-16T12:00:00Z",
                },
                newIds: {
                  control: Array.from(
                    { length: 40 },
                    (_, index) => 1000 + index,
                  ),
                  revision: Array.from(
                    { length: 40 },
                    (_, index) => 2000 + index,
                  ),
                },
              });
              if (tracked.isErr()) throw tracked.error;
              for (const decision of ["accept", "reject"] as const) {
                const resolved = applyDocumentOp(tracked.value.document, {
                  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                  story: OP_STORIES.MAIN,
                  revisionIds: tracked.value.revisions,
                  decision,
                });
                if (resolved.isErr()) throw resolved.error;
                expect(resolved.value.document).toStrictEqual(
                  decision === "accept" ? direct.document : document,
                );
                expect(undone(resolved.value)).toStrictEqual(
                  tracked.value.document,
                );
              }
              expect(undone(direct)).toStrictEqual(document);
            },
          ),
          {
            id: `balanced ${family} deletion source ownership`,
            numRuns: 30,
            examples: [
              ["start", 3],
              ["end", 3],
              ["whole", 3],
            ],
          },
        );
      },
      propertyTestTimeout(30_000),
    );

    test(`a ${family} deletion distinguishes repeated immutable endpoint occurrences`, () => {
      const document = documentOf(
        paragraph("00000001", [
          run("a"),
          start,
          run("b"),
          end,
          run("c"),
          structuredClone(start),
          run("d"),
          end,
          run("e"),
        ]),
      );
      const result = deleteIn({
        document: document,
        blockId: "00000001",
        from: 0,
        to: 3,
      });
      expect(contentOf(result.document, 0)).toEqual([
        start,
        run("d"),
        end,
        run("e"),
      ]);
      expect(undone(result)).toStrictEqual(document);

      const ambiguous = documentOf(
        paragraph("00000001", [
          run("a"),
          start,
          run("b"),
          structuredClone(start),
          run("c"),
          end,
          run("d"),
          end,
          run("e"),
        ]),
      );
      const balancedSubset = deleteIn({
        document: ambiguous,
        blockId: "00000001",
        from: 1,
        to: 4,
      });
      expect(contentOf(balancedSubset.document, 0)).toEqual([
        run("a"),
        start,
        end,
        run("e"),
      ]);
      expect(undone(balancedSubset)).toStrictEqual(ambiguous);

      const nested = documentOf(
        paragraph("00000001", [
          run("a"),
          nestedBoundary({ node: start, depth: 1, id: 10 }),
          run("b"),
          nestedBoundary({ node: end, depth: 1, id: 20 }),
          run("c"),
          nestedBoundary({ node: structuredClone(start), depth: 1, id: 30 }),
          run("d"),
          nestedBoundary({ node: end, depth: 1, id: 40 }),
          run("e"),
        ]),
      );
      const direct = deleteIn({
        document: nested,
        blockId: "00000001",
        from: 0,
        to: 3,
      });
      const tracked = applyDocumentOp(nested, {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: position("00000001", 0),
        to: position("00000001", 3),
        revision: { id: 500, author: "Reviewer", date: "2026-11-16T12:00:00Z" },
        newIds: { control: [1000, 1001, 1002], revision: [2000, 2001, 2002] },
      });
      if (tracked.isErr()) throw tracked.error;
      const accepted = applyDocumentOp(tracked.value.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: tracked.value.revisions,
        decision: "accept",
      });
      if (accepted.isErr()) throw accepted.error;
      expect(accepted.value.document).toStrictEqual(direct.document);
      expect(undone(accepted.value)).toStrictEqual(tracked.value.document);
      expect(undone(direct)).toStrictEqual(nested);
      const rejected = applyDocumentOp(tracked.value.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: tracked.value.revisions,
        decision: "reject",
      });
      if (rejected.isErr()) throw rejected.error;
      expect(rejected.value.document).toStrictEqual(nested);
    });

    test(`a ${family} range's end inside the deletion stays at its edge`, () => {
      // "ab" [start] "cd" [end] "ef": delete "d", the end and "e".
      const document = documentOf(
        paragraph("00000001", [run("ab"), start, run("cd"), end, run("ef")]),
      );
      const original = structuredClone(document);

      const applied = deleteIn({
        document: document,
        blockId: "00000001",
        from: 3,
        to: 5,
      });

      expect(contentOf(applied.document, 0)).toEqual([
        run("ab"),
        start,
        run("c"),
        end,
        run("f"),
      ]);
      expect(undone(applied)).toStrictEqual(original);

      // Break proof: removing every leaf between the gaps, as the deletion
      // did before, leaves the range with a start and no end.
      const unkept = deleteBetween(
        contentOf(original, 0),
        { offset: 3, zeroWidthBefore: 0 },
        { offset: 5, zeroWidthBefore: 0 },
      ).content;
      expect(unkept.some(({ type }) => type === end.type)).toBe(false);
      expect(unkept.some(({ type }) => type === start.type)).toBe(true);
    });

    test(`a ${family} range's start inside the deletion stays at its edge`, () => {
      const document = documentOf(
        paragraph("00000001", [run("ab"), start, run("cd"), end]),
      );
      const original = structuredClone(document);

      const applied = deleteIn({
        document: document,
        blockId: "00000001",
        from: 1,
        to: 3,
      });

      expect(contentOf(applied.document, 0)).toEqual([
        run("a"),
        start,
        run("d"),
        end,
      ]);
      expect(undone(applied)).toStrictEqual(original);
    });

    test(`a ${family} range whose other end is in another paragraph keeps its start`, () => {
      const document = documentOf(
        paragraph("00000001", [run("ab"), start, run("cd")]),
        paragraph("00000002", [run("x"), end]),
      );

      const applied = deleteIn({
        document: document,
        blockId: "00000001",
        from: 1,
        to: 3,
      });

      expect(contentOf(applied.document, 0)).toEqual([
        run("a"),
        start,
        run("d"),
      ]);
      expect(contentOf(applied.document, 1)).toEqual([run("x"), end]);
    });

    test(`a ${family} range the deletion covers whole goes whole`, () => {
      const document = documentOf(
        paragraph("00000001", [run("ab"), start, run("cd"), end, run("ef")]),
      );

      const applied = deleteIn({
        document: document,
        blockId: "00000001",
        from: 1,
        to: 5,
      });

      expect(contentOf(applied.document, 0)).toEqual([run("a"), run("f")]);
    });

    test(`a ${family} range already inverted loses the end the deletion covers`, () => {
      // Not this deletion's range to repair: the end before its start goes
      // as the rest of the deleted content does.
      const document = documentOf(
        paragraph("00000001", [run("ab"), end, run("cd"), start]),
      );

      const applied = deleteIn({
        document: document,
        blockId: "00000001",
        from: 1,
        to: 3,
      });

      expect(contentOf(applied.document, 0)).toEqual([
        run("a"),
        run("d"),
        start,
      ]);
    });
  }
});
