import { Result } from "better-result";
import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../../model/document";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
  type EditorIntent,
} from "../../editorIntent";
import { equalForStaleness } from "../../equality";
import { formatParaId } from "../../idLease";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
} from "../../types";
import { applyBatch } from "../applyBatch";
import { createClient } from "../client";
import {
  type DocumentBatch,
  NO_SEQUENCED_EFFECT,
  type SequencedBatch,
  validateDocumentBatch,
} from "../envelope";
import { transformBatch } from "../transform";
import { canonicalBase } from "./canonicalEdits";

/*
 * A revision decision over a concurrent direct paragraph insertion: the
 * insertion's sequenced effect names the revisions it borders (the marks of
 * its anchor and of the paragraph before it, and the table rows and cells
 * holding it), and a decision naming none of them passes it in either order.
 */

const wire = <Value>(value: Value): Value => JSON.parse(JSON.stringify(value));

const DATE = "2026-01-01T00:00:00Z";
/** The base's content revision, a suggested insertion in this paragraph. */
const SUGGESTED = formatParaId(0x1000_0011);
const CONTENT_REVISION = 10;
const PLAIN = formatParaId(0x1000_0001);
const NEXT = formatParaId(0x1000_0002);
const CELL = formatParaId(0x2000_0000);
const MARK_REVISION = 12;
const ROW_REVISION = 13;

type Transform = typeof transformBatch;

const batchOf = (opId: string, ops: readonly DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops,
});

const decisionOf = (
  decision: "accept" | "reject",
  revisionId: number,
): DocumentBatch =>
  batchOf("decision", [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [revisionId],
      decision,
    },
  ]);

/** A direct paragraph insertion, compiled as an editing session does. */
const insertionOf = (
  document: Document,
  at: { type: "before" | "after"; blockId: string },
): DocumentBatch => {
  const intent: EditorIntent = {
    type: "insertParagraphs",
    story: OP_STORIES.MAIN,
    at,
    paragraphs: [
      {
        type: "paragraph",
        content: [{ type: "run", content: [{ type: "text", text: "New" }] }],
      },
    ],
  };
  const allocation = createEditorIntentIdAllocator()(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: allocation.newIds },
  });
  if (compiled.isErr()) throw new Error(compiled.error.message);
  return batchOf("insertion", compiled.value.ops);
};

/** `first` sequenced with its effects, then `second` rebased over it. */
const inOrder = (
  base: Document,
  first: DocumentBatch,
  second: DocumentBatch,
  transform: Transform = transformBatch,
  effectsOf: (sequenced: SequencedBatch) => SequencedBatch = (sequenced) =>
    sequenced,
): { document: Document } | { refused: string } => {
  const applied = applyBatch(base, first.ops);
  if (applied.isErr())
    return { refused: `${first.opId}: ${applied.error.message}` };
  const sequenced = effectsOf({
    ...first,
    revision: 1,
    effects: applied.value.effects,
  });
  const valid = validateDocumentBatch(wire(sequenced));
  if (valid.isErr()) throw valid.error;
  const rebased = transform(second, [sequenced]);
  if (rebased.isErr())
    return { refused: `${second.opId} rebased: ${rebased.error.message}` };
  const after = applyBatch(applied.value.document, rebased.value.ops);
  return after.isErr()
    ? { refused: `${second.opId} applied: ${after.error.message}` }
    : { document: after.value.document };
};

const withParagraph = (
  document: Document,
  paraId: string,
  change: (paragraph: Paragraph) => Paragraph,
): Document => ({
  ...document,
  package: {
    ...document.package,
    document: {
      ...document.package.document,
      content: document.package.document.content.map((block) =>
        block.type === "paragraph" && block.paraId === paraId
          ? change(block)
          : block,
      ),
    },
  },
});

/** The canonical base with an inserted mark on the plain paragraph. */
const markedBase = (): Document =>
  withParagraph(canonicalBase(), PLAIN, (paragraph) => ({
    ...paragraph,
    pPrMark: {
      kind: "ins",
      info: { id: MARK_REVISION, author: "base", date: DATE },
    },
  }));

/** The canonical base with its table's first row a pending insertion. */
const insertedRowBase = (): Document => {
  const document = canonicalBase();
  for (const block of document.package.document.content) {
    const row = block.type === "table" ? block.rows.at(0) : undefined;
    if (row === undefined) continue;
    row.structuralChange = {
      type: "tableRowInsertion",
      info: { id: ROW_REVISION, author: "base", date: DATE },
    };
  }
  return document;
};

const BORDERING_REFUSAL =
  "decision rebased: A concurrent paragraph insertion borders a paragraph mark this decision resolves.";

/** The test-only break: a decision passes every direct paragraph insertion. */
const passingInsertions: Transform = (batch, over, options) => {
  const rebased = transformBatch(batch, over, options);
  return rebased.isErr() &&
    rebased.error.overType === DOCUMENT_OP_TYPES.INSERT_BLOCKS
    ? Result.ok(batch)
    : rebased;
};

describe("a revision decision over a concurrent direct paragraph insertion", () => {
  const anchors = [
    { type: "after", blockId: SUGGESTED },
    { type: "before", blockId: SUGGESTED },
    { type: "after", blockId: PLAIN },
  ] as const;
  for (const at of anchors)
    for (const decision of ["accept", "reject"] as const)
      test(`passes an insertion ${at.type} ${at.blockId} (${decision})`, () => {
        const base = canonicalBase();
        const insertion = insertionOf(base, at);
        const decided = decisionOf(decision, CONTENT_REVISION);
        // Sequenced after the insertion, the decision passes it: the
        // document is the decision's followed by the insertion.
        const insertionFirst = inOrder(base, insertion, decided);
        const decidedThenInserted = applyBatch(
          applyBatch(base, decided.ops).unwrap().document,
          insertion.ops,
        ).unwrap().document;
        if (!("document" in insertionFirst))
          throw new Error(insertionFirst.refused);
        expect(
          equalForStaleness(insertionFirst.document, decidedThenInserted),
        ).toBe(true);
        // Sequenced after the decision, the insertion passes it where the
        // decision left its anchor alone, and is refused, not guessed, where
        // the decision changed it.
        const decisionFirst = inOrder(base, decided, insertion);
        if (at.blockId === SUGGESTED) {
          expect(decisionFirst).toEqual({
            refused:
              "insertion rebased: The operation targets a structurally changed block.",
          });
          return;
        }
        if (!("document" in decisionFirst))
          throw new Error(decisionFirst.refused);
        expect(
          equalForStaleness(decisionFirst.document, decidedThenInserted),
        ).toBe(true);
      });

  test("records the revisions an insertion borders as its effect", () => {
    const effectOf = (document: Document, at: (typeof anchors)[number]) =>
      applyBatch(document, insertionOf(document, at).ops).unwrap().effects;
    expect(effectOf(canonicalBase(), anchors[0])).toEqual([
      { type: DOCUMENT_OP_TYPES.INSERT_BLOCKS, borderingRevisionIds: [] },
    ]);
    // The anchor's mark, and the mark of the paragraph before the anchor.
    expect(effectOf(markedBase(), { type: "after", blockId: PLAIN })).toEqual([
      {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        borderingRevisionIds: [MARK_REVISION],
      },
    ]);
    expect(effectOf(markedBase(), { type: "before", blockId: NEXT })).toEqual([
      {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        borderingRevisionIds: [MARK_REVISION],
      },
    ]);
    // The row holding the anchor.
    expect(
      effectOf(insertedRowBase(), { type: "after", blockId: CELL }),
    ).toEqual([
      {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        borderingRevisionIds: [ROW_REVISION],
      },
    ]);
  });

  test("refuses a decision of a mark the insertion borders, which does not commute", () => {
    for (const at of [
      { type: "after", blockId: PLAIN },
      { type: "before", blockId: NEXT },
    ] as const) {
      const base = markedBase();
      const insertion = insertionOf(base, at);
      const rejection = decisionOf("reject", MARK_REVISION);
      expect(inOrder(base, insertion, rejection)).toEqual({
        refused: BORDERING_REFUSAL,
      });
      // Passed anyway, rejecting the mark joins the paragraph with the
      // inserted one in one order and with the one after it in the other.
      const passed = inOrder(base, insertion, rejection, passingInsertions);
      const decisionFirst = inOrder(base, rejection, insertion);
      expect(
        "document" in passed &&
          "document" in decisionFirst &&
          equalForStaleness(passed.document, decisionFirst.document),
      ).toBe(false);
    }
  });

  test("refuses a decision of a row holding the insertion", () => {
    const base = insertedRowBase();
    const insertion = insertionOf(base, { type: "after", blockId: CELL });
    expect(
      inOrder(base, insertion, decisionOf("reject", ROW_REVISION)),
    ).toEqual({ refused: BORDERING_REFUSAL });
  });

  test("an insertion sequenced before the effect still replays, and refuses a decision over it", () => {
    const base = canonicalBase();
    const insertion = insertionOf(base, { type: "after", blockId: SUGGESTED });
    const applied = applyBatch(base, insertion.ops).unwrap();
    // A journal written before the effect recorded none for it.
    const journaled: SequencedBatch = {
      ...insertion,
      revision: 1,
      effects: [{ type: NO_SEQUENCED_EFFECT }],
    };
    expect(validateDocumentBatch(wire(journaled)).isOk()).toBe(true);
    const replayed = createClient(base);
    expect(replayed.receiveBroadcast(wire(journaled)).isOk()).toBe(true);
    expect(equalForStaleness(replayed.document, applied.document)).toBe(true);
    // A decision pending over it is refused, with a notice, not passed.
    const pending = createClient(base);
    const decided = decisionOf("accept", CONTENT_REVISION);
    expect(pending.enqueue(wire(decided)).isOk()).toBe(true);
    expect(pending.receiveBroadcast(wire(journaled)).isOk()).toBe(true);
    expect(
      pending.notices.map(({ opId, reason }) => [opId, reason.message]),
    ).toEqual([
      [
        "decision",
        "A concurrent paragraph insertion records no revisions it borders.",
      ],
    ]);
    expect(equalForStaleness(pending.document, applied.document)).toBe(true);
  });
});
