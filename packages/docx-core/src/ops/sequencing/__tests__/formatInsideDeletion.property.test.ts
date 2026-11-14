import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document, RunContent } from "../../../model/document";
import { equalForStaleness } from "../../equality";
import { compareGaps, type Gap, leafSpans } from "../../leaves";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type TextPosition,
} from "../../types";
import { applyBatch } from "../applyBatch";
import type { DocumentBatch, SequencedBatch } from "../envelope";
import { transformBatch } from "../transform";

/**
 * One leaf of every kind a run holds. The record is exhaustive, so a new
 * kind of run content fails to typecheck here until it has a sample.
 */
const LEAVES: {
  readonly [Kind in RunContent["type"]]: RunContent & { type: Kind };
} = {
  text: { type: "text", text: "x" },
  tab: { type: "tab" },
  break: { type: "break" },
  symbol: { type: "symbol", font: "Symbol", char: "F020" },
  footnoteRef: { type: "footnoteRef", id: 1 },
  endnoteRef: { type: "endnoteRef", id: 1 },
  noteMarker: { type: "noteMarker", kind: "footnote" },
  fieldChar: { type: "fieldChar", charType: "begin" },
  instrText: { type: "instrText", text: " PAGE " },
  softHyphen: { type: "softHyphen" },
  noBreakHyphen: { type: "noBreakHyphen" },
  renderedPageBreak: { type: "renderedPageBreak" },
  preservedXml: { type: "preservedXml", xml: "<w:ptab/>", text: "" },
  drawing: {
    type: "drawing",
    image: {
      type: "image",
      size: { width: 10, height: 10 },
      wrap: { type: "inline" },
    },
  },
  shape: {
    type: "shape",
    shape: {
      type: "shape",
      shapeType: "rect",
      size: { width: 10, height: 10 },
    },
  },
};

const SAMPLES: readonly RunContent[] = Object.values(LEAVES);

const BLOCK = "00000001";

const documentOf = (content: readonly RunContent[]): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: BLOCK,
          content: [{ type: "run", content: [...content] }],
        },
      ],
    },
  },
});

const position = ({ offset, zeroWidthBefore }: Gap): TextPosition => ({
  story: "main",
  blockId: BLOCK,
  offset,
  zeroWidthBefore,
});

/** Every gap of the run, in order. */
const gapsOf = (content: readonly RunContent[]): Gap[] => {
  const all = leafSpans([{ type: "run", content: [...content] }]).flatMap(
    ({ before, after }) => [before, after],
  );
  const seen = new Set<string>();
  return all
    .filter((gap) => {
      const key = `${String(gap.offset)}/${String(gap.zeroWidthBefore)}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    })
    .toSorted(compareGaps);
};

const batchOf = (opId: string, op: DocumentOp): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops: [op],
});

/** `first` sequenced, then `second` rebased over it, or the step refused. */
const inOrder = (
  base: Document,
  first: DocumentBatch,
  second: DocumentBatch,
  forget: boolean,
): { document: Document } | { refused: string } => {
  const applied = applyBatch(base, first.ops);
  if (applied.isErr())
    return { refused: `${first.opId}: ${applied.error.message}` };
  const sequenced: SequencedBatch = {
    ...first,
    revision: 1,
    // The test-only break: a format that records no cut records.
    effects: forget
      ? applied.value.effects.map((effect) =>
          effect.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS
            ? { type: DOCUMENT_OP_TYPES.SET_RUN_PROPS }
            : effect,
        )
      : applied.value.effects,
  };
  const rebased = transformBatch(second, [sequenced]);
  if (rebased.isErr())
    return { refused: `${second.opId} rebased: ${rebased.error.message}` };
  const after = applyBatch(applied.value.document, rebased.value.ops);
  return after.isErr()
    ? {
        refused: `${second.opId} as ${JSON.stringify(rebased.value.ops)}: ${after.error.message}`,
      }
    : { document: after.value.document };
};

/**
 * A run of text and leaves of any kinds, a direct deletion, and a direct
 * format inside it, as four ordered gaps: deletion start, format start,
 * format end, deletion end.
 */
const leaves = fc.array(
  fc.oneof(
    fc.constantFrom<RunContent>(
      { type: "text", text: "ab" },
      { type: "text", text: "cde" },
    ),
    fc.constantFrom(...SAMPLES),
  ),
  { maxLength: 4 },
);

const scenario = fc
  .tuple(leaves, leaves)
  .map(([left, right]): RunContent[] =>
    left.concat([{ type: "text", text: "fgh" }], right),
  )
  .chain((content) => {
    const gaps = gapsOf(content);
    return fc
      .array(fc.nat({ max: gaps.length - 1 }), {
        minLength: 4,
        maxLength: 4,
      })
      .map((picked) => {
        const [deleteFrom, formatFrom, formatTo, deleteTo] = picked
          .toSorted((left, right) => left - right)
          .map((index) => gaps[index] ?? gaps[0]);
        return { content, deleteFrom, formatFrom, formatTo, deleteTo };
      })
      .filter(
        ({ formatFrom, formatTo }) =>
          formatFrom !== undefined &&
          formatTo !== undefined &&
          formatFrom.offset < formatTo.offset,
      );
  });

type Scenario =
  typeof scenario extends fc.Arbitrary<infer Value> ? Value : never;

const EXAMPLES: [Scenario][] = [
  // Separate text on both sides of a formatted tab stays separate.
  [
    {
      content: [
        { type: "text", text: "ab" },
        LEAVES.tab,
        { type: "text", text: "cde" },
      ],
      deleteFrom: { offset: 1, zeroWidthBefore: 0 },
      formatFrom: { offset: 2, zeroWidthBefore: 0 },
      formatTo: { offset: 3, zeroWidthBefore: 0 },
      deleteTo: { offset: 4, zeroWidthBefore: 0 },
    },
  ],
  // A zero-width leaf after the deletion keeps the run one.
  [
    {
      content: [{ type: "text", text: "cde" }, LEAVES.renderedPageBreak],
      deleteFrom: { offset: 1, zeroWidthBefore: 0 },
      formatFrom: { offset: 2, zeroWidthBefore: 0 },
      formatTo: { offset: 3, zeroWidthBefore: 0 },
      deleteTo: { offset: 3, zeroWidthBefore: 0 },
    },
  ],
];

const rejoined = (forget = false) =>
  fc.property(
    scenario,
    ({ content, deleteFrom, formatFrom, formatTo, deleteTo }) => {
      if (
        deleteFrom === undefined ||
        formatFrom === undefined ||
        formatTo === undefined ||
        deleteTo === undefined
      )
        throw new Error("Four gaps are picked.");
      const base = documentOf(content);
      const format = batchOf("format", {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        from: position(formatFrom),
        to: position(formatTo),
        patch: { bold: true },
      });
      const deletion = batchOf("deletion", {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: position(deleteFrom),
        to: position(deleteTo),
      });
      const formatFirst = inOrder(base, format, deletion, forget);
      const deletionFirst = inOrder(base, deletion, format, forget);
      // Declared, a separate class: a position after zero-width leaves
      // is refused against a concurrent deletion holding it, and a
      // deletion bounded there refuses every concurrent operation.
      const zeroWidthBound = [deleteFrom, formatFrom, formatTo, deleteTo].some(
        ({ zeroWidthBefore }) => zeroWidthBefore > 0,
      );
      for (const outcome of [formatFirst, deletionFirst])
        if ("refused" in outcome)
          expect({ zeroWidthBound, refused: outcome.refused }).toEqual({
            zeroWidthBound: true,
            refused: expect.stringMatching(
              /zero-width boundaries require an unchanged range|cannot be mapped without structural context/u,
            ),
          });
      if (!("document" in formatFirst) || !("document" in deletionFirst))
        return;
      expect(
        equalForStaleness(formatFirst.document, deletionFirst.document)
          ? undefined
          : {
              formatFirst: formatFirst.document.package.document.content,
              deletionFirst: deletionFirst.document.package.document.content,
            },
      ).toBeUndefined();
    },
  );

describe("a direct format inside a concurrent direct deletion", () => {
  test(
    "leaves the same document in either order, whatever leaves the run holds",
    () => {
      fc.assert(
        rejoined(),
        propertyConfig({ numRuns: 300, examples: EXAMPLES }),
      );
    },
    propertyTestTimeout(60_000),
  );

  test("a format that records no cut records is caught", () => {
    const result = fc.check(
      rejoined(true),
      // The break is a counterexample on purpose: report nothing.
      propertyConfig({
        numRuns: 1,
        examples: EXAMPLES,
        reporter: () => undefined,
      }),
    );
    expect(result.failed).toBe(true);
  });
});
