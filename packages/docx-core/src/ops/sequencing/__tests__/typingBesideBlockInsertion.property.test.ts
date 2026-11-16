import { Result } from "better-result";
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document } from "../../../model/document";
import { storyBody, storyParagraphs } from "../../blocks";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
  type EditorIntent,
  type EditorIntentMode,
} from "../../editorIntent";
import { equalForStaleness } from "../../equality";
import { formatParaId } from "../../idLease";
import { paragraphLength } from "../../offsets";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type TextPosition,
} from "../../types";
import { applyBatch } from "../applyBatch";
import * as envelope from "../envelope";
import {
  BATCH_REJECTION_REASONS,
  type DocumentBatch,
  type SequencedBatch,
} from "../envelope";
import { isInlineInsertion } from "../points";
import { transformBatch } from "../transform";
import { canonicalBase } from "./canonicalEdits";

/*
 * A paragraph insertion names its anchor by id and moves no offset in it:
 * typing in the anchor paragraph and the insertion pass each other.
 */

/** Typed paragraphs of the base: plain, in a table cell, holding a suggestion, last. */
const PLAIN = formatParaId(0x1000_0001);
const CELL = formatParaId(0x2000_0000);
const SUGGESTED = formatParaId(0x1000_0011);
const LAST = formatParaId(0x1000_0005);

type Transform = typeof transformBatch;

const batchOf = (opId: string, ops: readonly DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops,
});

/** `first` sequenced, then `second` rebased over it, or the step refused. */
const inOrder = (
  base: Document,
  first: DocumentBatch,
  second: DocumentBatch,
  transform: Transform,
): { document: Document } | { refused: string } => {
  const applied = applyBatch(base, first.ops);
  if (applied.isErr())
    return { refused: `${first.opId}: ${applied.error.message}` };
  const sequenced: SequencedBatch = {
    ...first,
    revision: 1,
    effects: applied.value.effects,
  };
  const rebased = transform(second, [sequenced]);
  if (rebased.isErr())
    return { refused: `${second.opId} rebased: ${rebased.error.message}` };
  const after = applyBatch(applied.value.document, rebased.value.ops);
  return after.isErr()
    ? { refused: `${second.opId} applied: ${after.error.message}` }
    : { document: after.value.document };
};

const scenario = fc.record({
  paragraph: fc.constantFrom(PLAIN, CELL, SUGGESTED, LAST),
  offset: fc.nat({ max: 30 }),
  kind: fc.constantFrom("type" as const, "tab" as const),
  suggest: fc.boolean(),
  side: fc.constantFrom("before" as const, "after" as const),
  tracked: fc.boolean(),
});

type Scenario =
  typeof scenario extends fc.Arbitrary<infer Value> ? Value : never;

/** The typing a scenario names, compiled as an editing session does. */
const typing = (
  document: Document,
  { paragraph, offset, kind, suggest }: Scenario,
): readonly DocumentOp[] | undefined => {
  const target = storyParagraphs(storyBody(document, OP_STORIES.MAIN)).find(
    ({ paragraph: candidate }) => candidate.paraId === paragraph,
  )?.paragraph;
  if (target === undefined) throw new Error(`${paragraph} is in the base.`);
  const at: TextPosition = {
    story: OP_STORIES.MAIN,
    blockId: paragraph,
    offset: offset % (paragraphLength(target) + 1),
  };
  const intent: EditorIntent =
    kind === "type"
      ? { type: "replaceText", from: at, to: at, text: "xy" }
      : { type: "insertAtom", from: at, to: at, atom: { type: "tab" } };
  const allocation = createEditorIntentIdAllocator()(document, intent);
  const mode: EditorIntentMode = suggest
    ? {
        type: "suggesting",
        revision: {
          id: allocation.revisionId,
          author: "typist",
          date: "2026-01-02T00:00:00.000Z",
        },
        newIds: allocation.newIds,
      }
    : { type: "editing", newIds: allocation.newIds };
  const compiled = compileEditorIntent(document, { intent, mode });
  return compiled.isOk() && compiled.value.ops.length > 0
    ? compiled.value.ops
    : undefined;
};

/** A paragraph inserted before or after the typed one, direct or tracked. */
const insertion = ({ paragraph, side, tracked }: Scenario): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: OP_STORIES.MAIN,
  at: { type: side, blockId: paragraph },
  blocks: [
    {
      type: "paragraph",
      paraId: "0000A001",
      content: [{ type: "run", content: [{ type: "text", text: "Inserted" }] }],
    },
  ],
  ...(tracked
    ? {
        revision: {
          id: 900,
          author: "inserter",
          date: "2026-01-03T00:00:00.000Z",
        },
        newIds: {
          revision: Array.from({ length: 16 }, (_, index) => 901 + index),
          control: [],
        },
      }
    : {}),
});

const refusedOf = (
  outcome: { document: Document } | { refused: string },
): string | undefined => ("refused" in outcome ? outcome.refused : undefined);

const besideTyping = (transform: Transform) =>
  fc.property(scenario, (picked) => {
    const base = canonicalBase();
    const typed = typing(base, picked);
    if (typed === undefined) return;
    const typingBatch = batchOf("typing", typed);
    const insertBatch = batchOf("insertion", [insertion(picked)]);
    const typingFirst = inOrder(base, typingBatch, insertBatch, transform);
    const insertionFirst = inOrder(base, insertBatch, typingBatch, transform);
    expect({
      typingFirst: refusedOf(typingFirst),
      insertionFirst: refusedOf(insertionFirst),
    }).toEqual({ typingFirst: undefined, insertionFirst: undefined });
    if (!("document" in typingFirst) || !("document" in insertionFirst)) return;
    expect(
      equalForStaleness(typingFirst.document, insertionFirst.document)
        ? undefined
        : {
            typingFirst: typingFirst.document.package.document.content,
            insertionFirst: insertionFirst.document.package.document.content,
          },
    ).toBeUndefined();
  });

const EXAMPLES: [Scenario][] = [
  // The shrunk convergence case: typing with seam joins in a paragraph, a
  // direct paragraph inserted after it.
  [
    {
      paragraph: PLAIN,
      offset: 8,
      kind: "type",
      suggest: false,
      side: "after",
      tracked: false,
    },
  ],
  // A tracked paragraph after the last one takes over its mark; the typing
  // is at its end, after the note reference.
  [
    {
      paragraph: LAST,
      offset: 17,
      kind: "type",
      suggest: false,
      side: "after",
      tracked: true,
    },
  ],
  // A suggested tab in a table cell, a tracked paragraph before it.
  [
    {
      paragraph: CELL,
      offset: 2,
      kind: "tab",
      suggest: true,
      side: "before",
      tracked: true,
    },
  ],
];

/** The test-only break: typing and a paragraph insertion beside it refuse each other. */
const refusingBeside: Transform = (batch, over, options) => {
  const typed = (ops: readonly DocumentOp[]) =>
    ops.some(
      (op) =>
        op.type === DOCUMENT_OP_TYPES.INSERT_TEXT ||
        op.type === DOCUMENT_OP_TYPES.SPLIT_INLINE ||
        op.type === DOCUMENT_OP_TYPES.JOIN_INLINE ||
        isInlineInsertion(op),
    );
  const inserts = (ops: readonly DocumentOp[]) =>
    ops.some(({ type }) => type === DOCUMENT_OP_TYPES.INSERT_BLOCKS);
  const crossing = over.some(
    ({ ops }) =>
      (typed(batch.ops) && inserts(ops)) || (inserts(batch.ops) && typed(ops)),
  );
  return crossing
    ? Result.err(
        new envelope.BatchRejection({
          reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
          message: "Break: no transform between typing and a block insertion.",
        }),
      )
    : transformBatch(batch, over, options);
};

describe("typing beside a concurrent paragraph insertion", () => {
  test(
    "passes it in either order and leaves the same document",
    () => {
      fc.assert(
        besideTyping(transformBatch),
        propertyConfig({ numRuns: 300, examples: EXAMPLES }),
      );
    },
    propertyTestTimeout(60_000),
  );

  test(
    "typing and an insertion that refuse each other are caught",
    () => {
      const result = fc.check(
        besideTyping(refusingBeside),
        // The break is a counterexample on purpose: report nothing.
        propertyConfig({
          numRuns: 1,
          examples: EXAMPLES.slice(0, 1),
          reporter: () => undefined,
        }),
      );
      expect(result.failed).toBe(true);
    },
    propertyTestTimeout(30_000),
  );
});
