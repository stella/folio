/**
 * A direct deletion and a direct format leave no two runs that could merge
 * meeting at their own points (the deletion's point, the format's edges), so
 * a format and a deletion that meet reach one structure in either order.
 * Runs meeting anywhere else, authored boundaries among them, stay. Every
 * edit's inverse restores the previous document exactly, the runs joined at
 * its points split again.
 */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { equalForStaleness } from "../equality";
import { paragraphLength } from "../offsets";
import { runSeams } from "../runNormalForm";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  type RunPropsPatch,
} from "../types";

const BLOCK_ID = "10000000";
const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: BLOCK_ID,
  offset,
});

const base = (
  texts: readonly string[] = ["Paragraph text here"],
): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: BLOCK_ID,
          content: texts.map((text) => ({
            type: "run" as const,
            content: [{ type: "text" as const, text }],
          })),
        },
      ],
    },
  },
});

const paragraphOf = (document: Document): Paragraph => {
  const found = storyParagraphs(document.package.document).at(0)?.paragraph;
  if (found === undefined) throw new Error("The paragraph is gone.");
  return found;
};

type Edit =
  | { kind: "type"; at: number; text: string }
  | { kind: "erase"; at: number; width: number }
  | { kind: "format"; at: number; width: number; patch: RunPropsPatch };

const PATCHES: readonly RunPropsPatch[] = [
  { bold: true },
  { bold: null },
  { italic: true },
  { italic: null },
];

const editArbitrary: fc.Arbitrary<Edit> = fc.oneof(
  fc.record({
    kind: fc.constant("type" as const),
    at: fc.nat({ max: 40 }),
    text: fc.constantFrom("x", "yz"),
  }),
  fc.record({
    kind: fc.constant("erase" as const),
    at: fc.nat({ max: 40 }),
    width: fc.integer({ min: 1, max: 4 }),
  }),
  fc.record({
    kind: fc.constant("format" as const),
    at: fc.nat({ max: 40 }),
    width: fc.integer({ min: 1, max: 6 }),
    patch: fc.constantFrom(...PATCHES),
  }),
);

/** Where a direct deletion or format meets the runs it left. */
const pointsOf = (op: DocumentOp): number[] => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
      return [op.from.offset];
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
      return [op.from.offset, op.to.offset];
    default:
      return [];
  }
};

/** The operation an edit makes of the paragraph as it stands. */
const opOf = (document: Document, edit: Edit): DocumentOp | undefined => {
  const length = paragraphLength(paragraphOf(document));
  const from = edit.at % (length + 1);
  switch (edit.kind) {
    case "type":
      return {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at(from),
        text: edit.text,
        runProps: INHERIT_RUN_PROPS,
      };
    case "erase":
    case "format": {
      const to = Math.min(length, from + edit.width);
      if (to <= from) return undefined;
      return edit.kind === "erase"
        ? {
            type: DOCUMENT_OP_TYPES.DELETE_RANGE,
            from: at(from),
            to: at(to),
          }
        : {
            type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
            from: at(from),
            to: at(to),
            patch: edit.patch,
          };
    }
    default: {
      const exhaustive: never = edit;
      return exhaustive;
    }
  }
};

describe("run normal form", () => {
  test(
    "a direct deletion or format leaves no mergeable runs at its points, and every inverse restores the previous structure",
    async () => {
      await assertProperty(
        fc.property(
          fc.array(editArbitrary, { minLength: 1, maxLength: 10 }),
          (edits) => {
            let document = base();
            for (const edit of edits) {
              const op = opOf(document, edit);
              if (op === undefined) continue;
              const applied = applyDocumentOps(document, [op]);
              if (applied.isErr()) continue;
              const next = applied.value.document;
              const points = pointsOf(op);
              expect(
                runSeams(paragraphOf(next)).filter(({ at: gap }) =>
                  points.includes(gap.offset),
                ),
              ).toEqual([]);
              const undone = applyDocumentOps(next, applied.value.inverse);
              if (undone.isErr()) throw undone.error;
              expect(equalForStaleness(undone.value.document, document)).toBe(
                true,
              );
              document = next;
            }
          },
        ),
        { id: "run normal form", numRuns: 200 },
      );
    },
    propertyTestTimeout(60_000),
  );

  test("a format inside a deletion leaves the same runs in either order", () => {
    const format: DocumentOp = {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at(10),
      to: at(12),
      patch: { bold: true },
    };
    const deletion: DocumentOp = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at(9),
      to: at(13),
    };
    // Applied after the deletion, the format has nothing left to cover.
    const formatFirst = applyDocumentOps(base(), [format, deletion]);
    const deletionFirst = applyDocumentOps(base(), [deletion]);
    if (formatFirst.isErr()) throw formatFirst.error;
    if (deletionFirst.isErr()) throw deletionFirst.error;
    expect(paragraphOf(formatFirst.value.document)).toEqual(
      paragraphOf(deletionFirst.value.document),
    );
    expect(paragraphOf(formatFirst.value.document).content).toHaveLength(1);
  });

  test("an authored boundary at a deletion's end: the same runs in either order, and undo restores it", () => {
    // Two alike authored runs, "Paragraph " and "text here"; the deletion
    // of "h " ends where the second starts, and the format lies inside it.
    const authored = base(["Paragraph ", "text here"]);
    const format: DocumentOp = {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at(8),
      to: at(10),
      patch: { bold: true },
    };
    const deletion: DocumentOp = {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at(8),
      to: at(10),
    };
    const formatFirst = applyDocumentOps(authored, [format, deletion]);
    const deletionFirst = applyDocumentOps(authored, [deletion]);
    if (formatFirst.isErr()) throw formatFirst.error;
    if (deletionFirst.isErr()) throw deletionFirst.error;
    expect(paragraphOf(formatFirst.value.document)).toEqual(
      paragraphOf(deletionFirst.value.document),
    );
    // The runs meeting at the deletion's point are one; undo puts the
    // deleted text and the authored boundary back.
    expect(paragraphOf(deletionFirst.value.document).content).toHaveLength(1);
    const undone = applyDocumentOps(
      deletionFirst.value.document,
      deletionFirst.value.inverse,
    );
    if (undone.isErr()) throw undone.error;
    expect(equalForStaleness(undone.value.document, authored)).toBe(true);
    // An edit elsewhere leaves the authored boundary alone.
    const elsewhere = applyDocumentOps(authored, [
      {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: at(2),
        to: at(3),
      },
    ]);
    if (elsewhere.isErr()) throw elsewhere.error;
    expect(paragraphOf(elsewhere.value.document).content).toHaveLength(2);
  });
});
