import { Result } from "better-result";
import { describe, expect, mock, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document, Paragraph } from "../../../model/document";
import { storyBody, storyParagraphs } from "../../blocks";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
  type EditorIntent,
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
import { createClient } from "../client";
import * as envelope from "../envelope";
import {
  BATCH_REJECTION_REASONS,
  type DocumentBatch,
  type SequencedBatch,
} from "../envelope";
import { isBlockScoped } from "../scope";
import { createSequencer } from "../sequencer";
import * as transformModule from "../transform";
import { canonicalBase } from "./canonicalEdits";

/*
 * A revision decision names revisions, not positions: it passes a concurrent
 * direct positional edit, and is refused only when the edit cut a revision
 * record in two and gave its second half a new id.
 */

const wire = <Value>(value: Value): Value => JSON.parse(JSON.stringify(value));

/** The base's only revision, a suggested insertion in this paragraph. */
const SUGGESTED = formatParaId(0x1000_0011);
const REVISION = 10;
/** Paragraphs of the base without a revision record. */
const PLAIN = formatParaId(0x1000_0001);
const LISTED = formatParaId(0x1000_0010);
const CELL = formatParaId(0x2000_0000);

const CUT_REFUSAL =
  "A concurrent edit cut a revision record this decision resolves.";

type Transform = typeof transformModule.transformBatch;

// Read once: a test-only break replaces the module's export.
const honestTransform: Transform = transformModule.transformBatch;

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

/** The ids of every revision record a document's body holds. */
const revisionIdsIn = (document: Document): number[] => {
  const ids: number[] = [];
  const visit = (node: unknown): void => {
    if (Array.isArray(node)) {
      for (const child of node) visit(child);
      return;
    }
    if (typeof node !== "object" || node === null) return;
    if (
      "type" in node &&
      (node.type === "insertion" || node.type === "deletion") &&
      "info" in node &&
      typeof node.info === "object" &&
      node.info !== null &&
      "id" in node.info &&
      typeof node.info.id === "number"
    )
      ids.push(node.info.id);
    for (const child of Object.values(node)) visit(child);
  };
  visit(document.package.document.content);
  return ids;
};

const paragraphIn = (document: Document, paraId: string) =>
  storyParagraphs(storyBody(document, OP_STORIES.MAIN)).find(
    ({ paragraph }) => paragraph.paraId === paraId,
  )?.paragraph;

const EDIT_KINDS = ["type", "erase", "replace", "tab", "format"] as const;

const scenario = fc.record({
  paragraph: fc.constantFrom(SUGGESTED, PLAIN, LISTED, CELL),
  kind: fc.constantFrom(...EDIT_KINDS),
  from: fc.nat({ max: 30 }),
  span: fc.nat({ max: 6 }),
  decision: fc.constantFrom("accept" as const, "reject" as const),
});

type Scenario =
  typeof scenario extends fc.Arbitrary<infer Value> ? Value : never;

/** The direct edit a scenario names, compiled as an editing session does. */
const directEdit = (
  document: Document,
  { paragraph, kind, from, span }: Scenario,
): readonly DocumentOp[] | undefined => {
  const target = paragraphIn(document, paragraph);
  if (target === undefined) throw new Error(`${paragraph} is in the base.`);
  const length = paragraphLength(target);
  const start = from % (length + 1);
  const end = Math.min(length, start + 1 + span);
  const at = (offset: number): TextPosition => ({
    story: OP_STORIES.MAIN,
    blockId: paragraph,
    offset,
  });
  let intent: EditorIntent;
  switch (kind) {
    case "type":
      intent = {
        type: "replaceText",
        from: at(start),
        to: at(start),
        text: "xy",
      };
      break;
    case "tab":
      intent = {
        type: "insertAtom",
        from: at(start),
        to: at(start),
        atom: { type: "tab" },
      };
      break;
    case "erase":
    case "replace":
      if (start >= length) return undefined;
      intent = {
        type: "replaceText",
        from: at(start),
        to: at(end),
        text: kind === "erase" ? "" : "q",
      };
      break;
    case "format":
      if (start >= length) return undefined;
      intent = {
        type: "formatRun",
        from: at(start),
        to: at(end),
        patch: { bold: true },
      };
      break;
    default: {
      const exhaustive: never = kind;
      return exhaustive;
    }
  }
  const allocation = createEditorIntentIdAllocator()(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: allocation.newIds },
  });
  return compiled.isOk() && compiled.value.ops.length > 0
    ? compiled.value.ops
    : undefined;
};

const refusedOf = (
  outcome: { document: Document } | { refused: string },
): string | undefined => ("refused" in outcome ? outcome.refused : undefined);

const decidedOver = (transform: Transform) =>
  fc.property(scenario, (picked) => {
    const base = canonicalBase();
    const ops = directEdit(base, picked);
    if (ops === undefined) return;
    const edit = batchOf("edit", ops);
    const decision = batchOf("decision", [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [REVISION],
        decision: picked.decision,
      },
    ]);
    const editFirst = inOrder(base, edit, decision, transform);
    if (picked.paragraph !== SUGGESTED) {
      // An edit away from the revision: both orders apply, alike.
      const decisionFirst = inOrder(base, decision, edit, transform);
      expect({
        editFirst: refusedOf(editFirst),
        decisionFirst: refusedOf(decisionFirst),
      }).toEqual({ editFirst: undefined, decisionFirst: undefined });
      if (!("document" in editFirst) || !("document" in decisionFirst)) return;
      expect(
        equalForStaleness(editFirst.document, decisionFirst.document)
          ? undefined
          : {
              editFirst: editFirst.document.package.document.content,
              decisionFirst: decisionFirst.document.package.document.content,
            },
      ).toBeUndefined();
      return;
    }
    // An edit beside or inside the revision: the decision passes it and
    // decides every record of the revision, or is refused because the edit
    // cut a record and gave its second half a new id, which the decision
    // does not name.
    if ("refused" in editFirst) {
      expect(editFirst.refused).toBe(`decision rebased: ${CUT_REFUSAL}`);
      return;
    }
    expect(revisionIdsIn(editFirst.document)).toEqual([]);
  });

const EXAMPLES: [Scenario][] = [
  // Typing in another paragraph.
  [{ paragraph: PLAIN, kind: "type", from: 3, span: 0, decision: "accept" }],
  // An erase in a table cell.
  [{ paragraph: CELL, kind: "erase", from: 1, span: 2, decision: "reject" }],
  // Typing in the revision's paragraph, before the suggested text.
  [
    {
      paragraph: SUGGESTED,
      kind: "type",
      from: 2,
      span: 0,
      decision: "reject",
    },
  ],
  // A tab inside the suggested text cuts its record in two.
  [{ paragraph: SUGGESTED, kind: "tab", from: 8, span: 0, decision: "accept" }],
];

/** The test-only break: a decision does not rebase over a direct edit. */
const refusingDecisions: Transform = (batch, over, options) =>
  batch.ops.some(({ type }) => type === DOCUMENT_OP_TYPES.RESOLVE_REVISION) &&
  over.some(({ ops }) => ops.some((op) => !isBlockScoped(op)))
    ? Result.err(
        new envelope.BatchRejection({
          reason: BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
          message: "Break: no decision over a direct edit.",
        }),
      )
    : honestTransform(batch, over, options);

/** The test-only break: a decision passes an edit that cut its record. */
const passingCuts: Transform = (batch, over, options) => {
  const rebased = honestTransform(batch, over, options);
  return rebased.isErr() && rebased.error.message === CUT_REFUSAL
    ? Result.ok(batch)
    : rebased;
};

const DOCUMENT_DATE = "2026-01-01T00:00:00Z";

/** The canonical base with a second suggested insertion, in its first paragraph. */
const twoRevisions = (): Document => {
  const document = canonicalBase();
  const first: Paragraph = {
    type: "paragraph",
    paraId: formatParaId(0x1000_0000),
    content: [
      { type: "run", content: [{ type: "text", text: "Paragraph " }] },
      {
        type: "insertion",
        info: { id: 11, author: "base", date: DOCUMENT_DATE },
        content: [{ type: "run", content: [{ type: "text", text: "zero" }] }],
      },
      { type: "run", content: [{ type: "text", text: " text" }] },
    ],
  };
  return {
    ...document,
    package: {
      ...document.package,
      document: {
        ...document.package.document,
        content: document.package.document.content.map((block) =>
          block.type === "paragraph" && block.paraId === first.paraId
            ? first
            : block,
        ),
      },
    },
  };
};

/**
 * Peer a accepts the suggestion in one paragraph and is sequenced first;
 * peer b typed in that paragraph and then rejected the suggestion in
 * another. Peer b's typing is dropped against the acceptance; its decision
 * rebases over the typing's inverse and is sequenced.
 */
const rejectionCascade = () => {
  const base = twoRevisions();
  const sequencer = createSequencer(base);
  const peers = { a: createClient(base), b: createClient(base) };
  const enqueue = (
    peer: keyof typeof peers,
    opId: string,
    ops: readonly DocumentOp[],
  ) => {
    const client = peers[peer];
    const queued = client.enqueue(
      wire({
        schema: DOCUMENT_OP_SCHEMA_VERSION,
        opId,
        actor: `actor-${peer}`,
        baseRev: client.headRev,
        ops,
      }),
    );
    if (queued.isErr()) throw queued.error;
  };
  enqueue("a", "accept", [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [REVISION],
      decision: "accept",
    },
  ]);
  enqueue("b", "typing", [
    {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story: OP_STORIES.MAIN, blockId: SUGGESTED, offset: 2 },
      text: "xy",
      runProps: "inherit",
    },
  ]);
  enqueue("b", "reject", [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: [11],
      decision: "reject",
    },
  ]);
  // Each answer follows the broadcasts before it, as the transport orders them.
  const submit = (peer: keyof typeof peers) => {
    const client = peers[peer];
    const next = client.nextSubmission();
    if (next === undefined) return;
    const result = sequencer.submit(wire(next));
    for (const receiver of Object.values(peers))
      for (const sequenced of sequencer.broadcasts.slice(receiver.headRev)) {
        const received = receiver.receiveBroadcast(wire(sequenced));
        if (received.isErr()) throw received.error;
      }
    if (result.type === "ack") client.receiveAck(result);
    else client.receiveReject(result);
  };
  submit("a");
  submit("b");
  return {
    dropped: peers.b.notices.map(({ opId }) => opId),
    sequenced: sequencer.broadcasts.map(({ opId }) => opId),
    converged: Object.values(peers).every(({ document }) =>
      equalForStaleness(document, sequencer.document),
    ),
    revisions: revisionIdsIn(sequencer.document),
  };
};

describe("a revision decision over a concurrent direct edit", () => {
  test(
    "passes it, or is refused only for a record the edit cut",
    () => {
      fc.assert(
        decidedOver(honestTransform),
        propertyConfig({ numRuns: 300, examples: EXAMPLES }),
      );
    },
    propertyTestTimeout(60_000),
  );

  test(
    "a decision that does not rebase over a direct edit is caught",
    () => {
      const result = fc.check(
        decidedOver(refusingDecisions),
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

  test(
    "a decision passing an edit that cut its record is caught",
    () => {
      const result = fc.check(
        decidedOver(passingCuts),
        // The break is a counterexample on purpose: report nothing.
        propertyConfig({
          numRuns: 1,
          examples: EXAMPLES.slice(3),
          reporter: () => undefined,
        }),
      );
      expect(result.failed).toBe(true);
    },
    propertyTestTimeout(30_000),
  );

  test("a decision survives the drop of its peer's unrelated typing before it", () => {
    expect(rejectionCascade()).toEqual({
      dropped: ["typing"],
      sequenced: ["accept", "reject"],
      converged: true,
      revisions: [],
    });
    // The test-only break: decisions do not rebase over a direct edit, so
    // the decision goes with the typing.
    const original = { ...transformModule };
    mock.module("../transform", () => ({
      ...original,
      transformBatch: refusingDecisions,
    }));
    try {
      expect(rejectionCascade().dropped).toEqual(["typing", "reject"]);
    } finally {
      mock.module("../transform", () => original);
    }
  });
});
