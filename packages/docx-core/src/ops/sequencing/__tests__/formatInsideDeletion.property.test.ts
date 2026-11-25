import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document, RunContent } from "../../../model/document";
import { RUN_LEAF_SAMPLES, RUN_LEAVES } from "../../__tests__/runLeafSamples";
import { equalForStaleness } from "../../equality";
import { compareGaps, type Gap, leafSpans } from "../../leaves";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type OpStory,
  type TextPosition,
} from "../../types";
import { applyBatch } from "../applyBatch";
import type { DocumentBatch, SequencedBatch } from "../envelope";
import { transformBatch } from "../transform";

const BLOCK = "00000001";

/**
 * The footnote a run holding a note's own mark lies in: not the note the
 * sampled note references name, which a deletion of one would remove.
 */
const NOTE = 9;

/**
 * Where the run lies: in the body, or in a footnote when it holds a note's
 * own reference mark, which the contract keeps to its note's story.
 */
const placeOf = (
  content: readonly RunContent[],
): { story: OpStory; document: Document } => {
  const paragraph = {
    type: "paragraph" as const,
    paraId: BLOCK,
    content: [{ type: "run" as const, content: [...content] }],
  };
  if (!content.some(({ type }) => type === "noteMarker"))
    return {
      story: OP_STORIES.MAIN,
      document: { package: { document: { content: [paragraph] } } },
    };
  return {
    story: { kind: "footnote", id: NOTE },
    document: {
      package: {
        document: {
          content: [{ type: "paragraph", paraId: "0000000F", content: [] }],
        },
        footnotes: [{ type: "footnote", id: NOTE, content: [paragraph] }],
      },
    },
  };
};

const position = (
  story: OpStory,
  { offset, zeroWidthBefore }: Gap,
): TextPosition => ({
  story,
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

/** The effect a test-only break records nothing in. */
type Forgotten =
  | typeof DOCUMENT_OP_TYPES.SET_RUN_PROPS
  | typeof DOCUMENT_OP_TYPES.DELETE_RANGE
  | "keptMarkOfCut";

const forgetting = (
  effect: NonNullable<SequencedBatch["effects"]>[number],
  forget: Forgotten | undefined,
): NonNullable<SequencedBatch["effects"]>[number] => {
  if (forget === "keptMarkOfCut")
    return effect.type === DOCUMENT_OP_TYPES.SET_RUN_PROPS &&
      effect.cutRecords !== undefined
      ? {
          ...effect,
          cutRecords: effect.cutRecords.map(({ from, to }) => ({ from, to })),
        }
      : effect;
  return effect.type === forget ? { type: forget } : effect;
};

/** `first` sequenced, then `second` rebased over it, or the step refused. */
const inOrder = (
  base: Document,
  first: DocumentBatch,
  second: DocumentBatch,
  forget: Forgotten | undefined,
): { document: Document } | { refused: string } => {
  const applied = applyBatch(base, first.ops);
  if (applied.isErr())
    return { refused: `${first.opId}: ${applied.error.message}` };
  const sequenced: SequencedBatch = {
    ...first,
    revision: 1,
    // The test-only breaks: an effect that records less than it does.
    effects: applied.value.effects.map((effect) => forgetting(effect, forget)),
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
    fc.constantFrom(...RUN_LEAF_SAMPLES),
  ),
  { maxLength: 4 },
);

const scenario = fc
  .tuple(leaves, leaves)
  .map(([left, right]): RunContent[] =>
    left.concat([{ type: "text", text: "fgh" }], right),
  )
  // A note holds one automatic reference mark.
  .filter(
    (content) =>
      content.filter(({ type }) => type === "noteMarker").length <= 1,
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
        RUN_LEAVES.tab,
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
      content: [{ type: "text", text: "cde" }, RUN_LEAVES.renderedPageBreak],
      deleteFrom: { offset: 1, zeroWidthBefore: 0 },
      formatFrom: { offset: 2, zeroWidthBefore: 0 },
      formatTo: { offset: 3, zeroWidthBefore: 0 },
      deleteTo: { offset: 3, zeroWidthBefore: 0 },
    },
  ],
  // A deletion of a note reference inside another note's paragraph.
  [
    {
      content: [
        RUN_LEAVES.noteMarker,
        { type: "text", text: "fgh" },
        RUN_LEAVES.footnoteRef,
        { type: "text", text: "ab" },
      ],
      deleteFrom: { offset: 0, zeroWidthBefore: 1 },
      formatFrom: { offset: 0, zeroWidthBefore: 1 },
      formatTo: { offset: 3, zeroWidthBefore: 0 },
      deleteTo: { offset: 5, zeroWidthBefore: 0 },
    },
  ],
  // A format and a deletion over a note's own mark, which the deletion keeps.
  [
    {
      content: [
        { type: "text", text: "fgh" },
        RUN_LEAVES.noteMarker,
        { type: "text", text: "ab" },
      ],
      deleteFrom: { offset: 0, zeroWidthBefore: 0 },
      formatFrom: { offset: 0, zeroWidthBefore: 0 },
      formatTo: { offset: 5, zeroWidthBefore: 0 },
      deleteTo: { offset: 5, zeroWidthBefore: 0 },
    },
  ],
  // A deletion to the run's end keeps the note's mark after the format:
  // the run's pieces meet at the deletion's point.
  [
    {
      content: [
        RUN_LEAVES.text,
        { type: "text", text: "ab" },
        { type: "text", text: "ab" },
        { type: "text", text: "fgh" },
        RUN_LEAVES.break,
        RUN_LEAVES.endnoteRef,
        RUN_LEAVES.noteMarker,
        RUN_LEAVES.shape,
      ],
      deleteFrom: { offset: 1, zeroWidthBefore: 0 },
      formatFrom: { offset: 1, zeroWidthBefore: 0 },
      formatTo: { offset: 3, zeroWidthBefore: 0 },
      deleteTo: { offset: 11, zeroWidthBefore: 0 },
    },
  ],
  // The kept mark stands before the format: the pieces meet after it.
  [
    {
      content: [
        { type: "text", text: "ab" },
        RUN_LEAVES.noteMarker,
        { type: "text", text: "cde" },
      ],
      deleteFrom: { offset: 1, zeroWidthBefore: 0 },
      formatFrom: { offset: 3, zeroWidthBefore: 0 },
      formatTo: { offset: 4, zeroWidthBefore: 0 },
      deleteTo: { offset: 4, zeroWidthBefore: 0 },
    },
  ],
  // A format and a deletion of all the text, a note's mark right after it.
  [
    {
      content: [{ type: "text", text: "fgh" }, RUN_LEAVES.noteMarker],
      deleteFrom: { offset: 0, zeroWidthBefore: 0 },
      formatFrom: { offset: 0, zeroWidthBefore: 0 },
      formatTo: { offset: 3, zeroWidthBefore: 0 },
      deleteTo: { offset: 3, zeroWidthBefore: 0 },
    },
  ],
];

// The examples in a note, where a deletion keeps the note's own mark: a run
// of the break covers each of them.
const NOTE_EXAMPLES = EXAMPLES.filter(([{ content }]) =>
  content.some(({ type }) => type === "noteMarker"),
);

const rejoined = (forget?: Forgotten) =>
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
      const { story, document: base } = placeOf(content);
      const format = batchOf("format", {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        from: position(story, formatFrom),
        to: position(story, formatTo),
        patch: { bold: true },
      });
      const deletion = batchOf("deletion", {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: position(story, deleteFrom),
        to: position(story, deleteTo),
      });
      const formatFirst = inOrder(base, format, deletion, forget);
      const deletionFirst = inOrder(base, deletion, format, forget);
      // Declared, a separate class: a position after zero-width leaves
      // is refused against a concurrent deletion holding it, and a
      // deletion bounded there refuses an operation with a position at
      // an offset its range spans.
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
              formatFirst: [
                formatFirst.document.package.document.content,
                formatFirst.document.package.footnotes,
              ],
              deletionFirst: [
                deletionFirst.document.package.document.content,
                deletionFirst.document.package.footnotes,
              ],
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

  test(
    "a format that records no cut records is caught",
    () => {
      const result = fc.check(
        rejoined(DOCUMENT_OP_TYPES.SET_RUN_PROPS),
        // The break is a counterexample on purpose: report nothing.
        propertyConfig({
          numRuns: 1,
          examples: EXAMPLES,
          reporter: () => undefined,
        }),
      );
      expect(result.failed).toBe(true);
    },
    propertyTestTimeout(30_000),
  );

  test(
    "a format whose cut records name no kept note mark is caught",
    () => {
      const result = fc.check(
        rejoined("keptMarkOfCut"),
        // The break is a counterexample on purpose: report nothing.
        propertyConfig({
          numRuns: NOTE_EXAMPLES.length,
          examples: NOTE_EXAMPLES,
          reporter: () => undefined,
        }),
      );
      expect(result.failed).toBe(true);
    },
    propertyTestTimeout(30_000),
  );

  test("typing after a deletion that keeps a note's mark lands after the mark in either order", () => {
    // The deletion of 0..5 keeps the note's mark at 3; the typing at 5 is
    // after it in the typing order, and must be after it in the other.
    const content: RunContent[] = [
      { type: "text", text: "fgh" },
      RUN_LEAVES.noteMarker,
      { type: "text", text: "ab" },
    ];
    const leavesOf = (document: Document): string[] =>
      leafSpans(
        document.package.footnotes
          ?.at(0)
          ?.content.flatMap((block) =>
            block.type === "paragraph" ? block.content : [],
          ) ?? [],
      ).map(({ node }) => (node.type === "text" ? node.text : node.type));
    const { story, document: base } = placeOf(content);
    const deletion = batchOf("deletion", {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: position(story, { offset: 0, zeroWidthBefore: 0 }),
      to: position(story, { offset: 5, zeroWidthBefore: 0 }),
    });
    const typing = batchOf("typing", {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: position(story, { offset: 5, zeroWidthBefore: 0 }),
      text: "XY",
      runProps: "inherit",
    });
    const outcomes = [
      inOrder(base, deletion, typing, undefined),
      inOrder(base, typing, deletion, undefined),
    ];
    expect(
      outcomes.map((outcome) =>
        "document" in outcome ? leavesOf(outcome.document) : outcome.refused,
      ),
    ).toEqual([
      ["noteMarker", "XY"],
      ["noteMarker", "XY"],
    ]);
    // Both orders leave one document, runs and all.
    const [deletionFirst, typingFirst] = outcomes;
    expect(
      deletionFirst !== undefined &&
        typingFirst !== undefined &&
        "document" in deletionFirst &&
        "document" in typingFirst &&
        equalForStaleness(deletionFirst.document, typingFirst.document),
    ).toBe(true);
    // The typing rebased over the deletion names the gap after the mark;
    // the test-only break, a deletion recording no kept mark, names the
    // gap before it.
    const rebasedAt = (forget: Forgotten | undefined) => {
      const applied = applyBatch(base, deletion.ops);
      if (applied.isErr()) throw new Error(applied.error.message);
      const rebased = transformBatch(typing, [
        {
          ...deletion,
          revision: 1,
          effects: applied.value.effects.map((effect) =>
            forgetting(effect, forget),
          ),
        },
      ]);
      if (rebased.isErr()) throw new Error(rebased.error.message);
      const [op] = rebased.value.ops;
      return op?.type === DOCUMENT_OP_TYPES.INSERT_TEXT ? op.at : op;
    };
    expect(rebasedAt(undefined)).toEqual(
      position(story, { offset: 0, zeroWidthBefore: 1 }),
    );
    expect(rebasedAt(DOCUMENT_OP_TYPES.DELETE_RANGE)).toEqual(
      position(story, { offset: 0, zeroWidthBefore: 0 }),
    );
  });
});
