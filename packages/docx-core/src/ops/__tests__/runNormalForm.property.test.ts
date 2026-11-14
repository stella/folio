/**
 * Every edit leaves the paragraph it touched in its run normal form: no two
 * consecutive runs that could merge. For the same text, formatting and ids
 * that structure is unique, so two orders of edits reaching the same content
 * reach the same runs. Every edit's inverse restores the previous document
 * exactly, the runs the normal form joined split again.
 */
import { describe, expect, mock, test } from "bun:test";
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

// Break: no paragraph has a seam, so apply never normalizes.
mock.module("../runNormalForm", () => ({ runSeams: () => [] }));

const BLOCK_ID = "10000000";
const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: BLOCK_ID,
  offset,
});

const base = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: BLOCK_ID,
          content: [
            {
              type: "run",
              content: [{ type: "text", text: "Paragraph text here" }],
            },
          ],
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
    "every edit leaves its paragraph in normal form and its inverse restores the previous structure",
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
              expect(runSeams(paragraphOf(next))).toEqual([]);
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
});
