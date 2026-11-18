import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../../../test/property-testing";
import type { Document } from "../../../model/document";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
  type EditorIntent,
} from "../../editorIntent";
import { type IdLease, leasedIdCensus } from "../../idLease";
import { paragraphLength } from "../../offsets";
import { revisionIdDemand } from "../../plan";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../../types";
import { applyBatch } from "../applyBatch";
import { createClient } from "../client";
import {
  BATCH_REJECTION_REASONS,
  type DocumentBatch,
  type SequencedBatch,
} from "../envelope";
import { createSequencer } from "../sequencer";
import { trackedOrderPair } from "../trackedTransform";
import { transformBatch } from "../transform";

const BLOCK = "10000000";
const OTHER = "10000001";
const STANDING = "10000002";
const wire = <Value>(value: Value): Value => JSON.parse(JSON.stringify(value));
const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: BLOCK,
  offset,
  zeroWidthBefore: 0,
});

type CompileOptions = {
  document: Document;
  actor: string;
  lease: IdLease;
  intent: EditorIntent;
  serial: number;
};

const suggested = ({
  document,
  actor,
  lease,
  intent,
  serial,
}: CompileOptions) => {
  const allocated = createEditorIntentIdAllocator()(document, intent, lease);
  return compileEditorIntent(document, {
    intent,
    mode: {
      type: "suggesting",
      revision: {
        id: allocated.revisionId,
        author: actor,
        date: new Date(Date.UTC(2026, 0, 1, 0, 0, serial)).toISOString(),
      },
      newIds: allocated.newIds,
      idLease: lease,
    },
  }).unwrap().ops;
};

const batch = (actor: string, ops: readonly DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId: actor,
  actor,
  baseRev: 0,
  ops,
});

const shapes = fc.record({
  prefix: fc.integer({ min: 2, max: 6 }),
  suffix: fc.integer({ min: 2, max: 6 }),
  mirror: fc.boolean(),
  bold: fc.boolean(),
});

type Shape = fc.ArbitraryValue<typeof shapes>;

/** Prior delete/insert records and the replacement are all real compiler output. */
const fixture = ({ prefix, suffix, mirror, bold }: Shape) => {
  const pristine = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: BLOCK,
            content: [
              {
                type: "run",
                formatting: { bold },
                content: [
                  {
                    type: "text",
                    text:
                      "p".repeat(prefix) +
                      (mirror ? "rabc" : "abcr") +
                      "s".repeat(suffix),
                  },
                ],
              },
            ],
          },
          {
            type: "paragraph",
            paraId: OTHER,
            content: [
              { type: "run", content: [{ type: "text", text: "Other" }] },
            ],
          },
          {
            type: "paragraph",
            paraId: STANDING,
            content: [
              { type: "run", content: [{ type: "text", text: "Standing" }] },
            ],
          },
        ],
      },
    },
  } satisfies Document;
  const authority = createSequencer(pristine);
  const oldLease = authority.grantIdLease("old").unwrap();
  const trackedLease = authority.grantIdLease("tracked").unwrap();
  authority.grantIdLease("direct").unwrap();
  let source: Document = pristine;
  const deletedStart = prefix + (mirror ? 1 : 0);
  const insertedAt = prefix + (mirror ? 1 : 3);
  const priorIntents = [
    {
      type: "replaceText",
      from: at(deletedStart),
      to: at(deletedStart + 3),
      text: "",
    },
    {
      type: "replaceText",
      from: at(insertedAt),
      to: at(insertedAt),
      text: "q",
    },
  ] as const satisfies readonly EditorIntent[];
  for (const [serial, intent] of priorIntents.entries()) {
    const ops = suggested({
      document: source,
      actor: "old",
      lease: oldLease,
      intent,
      serial,
    });
    const outcome = authority.submit({
      ...batch("old", ops),
      opId: `old-${serial}`,
      baseRev: serial,
    });
    if (outcome.type === "reject") throw outcome.reason;
    source = authority.document;
  }
  const standingAt = { ...at(0), blockId: STANDING };
  const standingOps = suggested({
    document: source,
    actor: "old",
    lease: oldLease,
    intent: {
      type: "replaceText",
      from: standingAt,
      to: standingAt,
      text: "pending",
    },
    serial: 2,
  });
  const standing = standingOps.find(
    (op) => op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT,
  );
  if (standing?.revision === undefined)
    panic("The independent paragraph has an older revision.");
  const standingOutcome = authority.submit({
    ...batch("old", standingOps),
    opId: "old-standing",
    baseRev: 2,
  });
  if (standingOutcome.type === "reject") throw standingOutcome.reason;
  source = authority.document;
  const from = prefix + (mirror ? 0 : 2);
  const to = from + 3;
  const ops = suggested({
    document: source,
    actor: "tracked",
    lease: trackedLease,
    intent: { type: "replaceText", from: at(from), to: at(to), text: "q" },
    serial: 4,
  });
  const deletion = ops.find((op) => op.type === DOCUMENT_OP_TYPES.DELETE_RANGE);
  if (
    deletion?.revision === undefined ||
    deletion.newIds?.revision?.length !== 1
  )
    panic("The source creates exactly two new deletion fragments.");
  expect(revisionIdDemand(source, deletion).unwrap()).toBe(1);
  const extra = deletion.newIds.revision.at(0);
  if (extra === undefined)
    panic("The second fragment has its own supplied id.");
  return {
    source,
    leases: authority.idLeases,
    trackedLease,
    standingId: standing.revision.id,
    tracked: batch("tracked", ops),
    firstId: deletion.revision.id,
    extraId: extra,
    ranges: {
      interior: [insertedAt, insertedAt + 1],
      leading: [from, insertedAt + 1],
      trailing: [insertedAt, to],
      crossesStart: [from - 1, insertedAt + 1],
      crossesEnd: [insertedAt, to + 1],
      equal: [from, to],
      covering: [from - 1, to + 1],
    } satisfies Record<string, readonly [number, number]>,
  };
};

type DirectDeleteOptions = { from: number; to: number };

const directDelete = ({ from, to }: DirectDeleteOptions) =>
  batch("direct", [
    {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at(from),
      to: at(to),
      newIds: { revision: [], control: [] },
    },
  ]);

type ProtocolOptions = {
  scenario: ReturnType<typeof fixture>;
  direct: DocumentBatch;
  first: "direct" | "tracked";
  id: number;
  decision: "accept" | "reject";
};

/** A decision is authored against the optimistic fragment, before either answer. */
const protocol = ({
  scenario,
  direct,
  first,
  id,
  decision,
}: ProtocolOptions) => {
  const authority = createSequencer(scenario.source, {
    idLeases: scenario.leases,
  });
  const peers = {
    tracked: createClient(scenario.source),
    direct: createClient(scenario.source),
  };
  const tracked = scenario.tracked;
  peers.tracked.enqueue(wire(tracked)).unwrap();
  peers.direct.enqueue(wire(direct)).unwrap();
  const review = {
    ...batch("tracked", [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [id],
        decision,
      },
    ]),
    opId: "review",
  };
  peers.tracked.enqueue(wire(review)).unwrap();
  const originalTracked = peers.tracked.document;
  const firstApplied = applyBatch(
    scenario.source,
    wire(first === "tracked" ? tracked.ops : direct.ops),
  ).unwrap();
  expect(
    applyBatch(firstApplied.document, wire(firstApplied.inverse)).unwrap()
      .document,
  ).toEqual(scenario.source);
  const outcome = authority.submit(
    wire(first === "tracked" ? tracked : direct),
  );
  if (outcome.type !== "ack") throw outcome.reason;
  const accepted = authority.document;
  expect(accepted).toEqual(firstApplied.document);
  const leasesAfterFirst = authority.idLeases;
  const headAfterFirst = authority.headRev;
  const second = authority.submit(wire(first === "tracked" ? direct : tracked));
  expect(second.type).toBe("reject");
  if (second.type !== "reject")
    panic("The overlapping fragment edit must refuse.");
  expect(second.reason.reason).toBe(BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR);
  expect(authority.document).toEqual(accepted);
  expect(authority.idLeases).toEqual(leasesAfterFirst);
  expect(authority.headRev).toBe(headAfterFirst);

  for (const peer of Object.values(peers)) {
    for (const broadcast of authority.broadcasts) {
      const before = peer.document;
      const changes = peer.receiveBroadcast(wire(broadcast)).unwrap();
      let replayed = before;
      for (const change of changes)
        replayed = applyBatch(replayed, wire(change.batch.ops)).unwrap()
          .document;
      expect(replayed).toEqual(peer.document);
    }
  }
  if (first === "direct") {
    expect(peers.tracked.notices.map(({ opId }) => opId)).toEqual([
      "tracked",
      "review",
    ]);
    expect(peers.tracked.pending).toHaveLength(0);
    expect(peers.tracked.document).toEqual(authority.document);
  } else {
    expect(peers.tracked.notices).toEqual([]);
    expect(peers.tracked.document).toEqual(originalTracked);
    const next = peers.tracked.nextSubmission();
    if (next?.opId !== "review")
      panic("Only the original-target review remains.");
    expect(next.ops).toEqual(review.ops);
    const reviewed = authority.submit(wire(next));
    if (reviewed.type !== "ack") throw reviewed.reason;
    for (const peer of Object.values(peers))
      for (const broadcast of authority.broadcasts.slice(peer.headRev))
        peer.receiveBroadcast(wire(broadcast)).unwrap();
    expect(authority.document).toEqual(
      applyBatch(
        applyBatch(scenario.source, tracked.ops).unwrap().document,
        review.ops,
      ).unwrap().document,
    );
  }
  for (const peer of Object.values(peers)) {
    expect(peer.pending).toHaveLength(0);
    expect(peer.document).toEqual(authority.document);
  }
};

test(
  "overlapping erasures cannot silently retarget a pending fragment decision",
  () => {
    fc.assert(
      fc.property(
        shapes,
        fc.constantFrom(...Object.values(REVISION_DECISIONS)),
        fc.boolean(),
        (shape, decision, extra) => {
          const scenario = fixture(shape);
          for (const [from, to] of Object.values(scenario.ranges))
            for (const first of ["tracked", "direct"] as const)
              protocol({
                scenario,
                direct: directDelete({ from: from, to: to }),
                first,
                id: extra ? scenario.extraId : scenario.firstId,
                decision,
              });
        },
      ),
      propertyConfig({
        numRuns: 20,
        examples: [false, true].flatMap((mirror) =>
          Object.values(REVISION_DECISIONS).flatMap((decision) =>
            [false, true].map(
              (extra) =>
                [
                  { prefix: 3, suffix: 3, mirror, bold: true },
                  decision,
                  extra,
                ] satisfies [Shape, RevisionDecision, boolean],
            ),
          ),
        ),
      }),
    );
  },
  propertyTestTimeout(30_000),
);

test("fragment-risk refusal is symmetric for either transform tie order", () => {
  for (const mirror of [false, true]) {
    const scenario = fixture({ prefix: 3, suffix: 3, mirror, bold: true });
    for (const [from, to] of Object.values(scenario.ranges)) {
      const direct = directDelete({ from: from, to: to });
      const directOp = direct.ops.at(0);
      if (directOp === undefined) panic("The direct deletion is stated.");
      for (const [first, second] of [
        [scenario.tracked, direct],
        [direct, scenario.tracked],
      ]) {
        if (first === undefined || second === undefined)
          panic("Both batches are present.");
        const applied = applyBatch(scenario.source, first.ops).unwrap();
        const sequenced: SequencedBatch = {
          ...first,
          revision: 1,
          effects: applied.effects,
        };
        for (const order of ["before", "after"] as const) {
          const moved = transformBatch(second, [sequenced], { order });
          expect(moved.isErr()).toBe(true);
          if (moved.isErr())
            expect(moved.error.reason).toBe(
              BATCH_REJECTION_REASONS.UNSUPPORTED_PAIR,
            );
        }
        for (const op of scenario.tracked.ops) {
          if (op.type !== DOCUMENT_OP_TYPES.DELETE_RANGE) continue;
          expect(trackedOrderPair(op, directOp)).toBeUndefined();
          expect(trackedOrderPair(directOp, op)).toBeUndefined();
        }
      }
    }
  }
});

test("ordinary unfragmented interior deletion still commutes with exact identities", () => {
  const source = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: BLOCK,
            content: [
              {
                type: "run",
                content: [{ type: "text", text: "prefix word suffix" }],
              },
            ],
          },
        ],
      },
    },
  } satisfies Document;
  const authority = createSequencer(source);
  const lease = authority.grantIdLease("tracked").unwrap();
  const tracked = batch(
    "tracked",
    suggested({
      document: source,
      actor: "tracked",
      lease,
      intent: { type: "replaceText", from: at(5), to: at(11), text: "" },
      serial: 0,
    }),
  );
  const deletion = tracked.ops.find(
    (op) => op.type === DOCUMENT_OP_TYPES.DELETE_RANGE,
  );
  if (deletion?.revision === undefined)
    panic("The compiler records the deletion.");
  expect(deletion.newIds?.revision ?? []).toEqual([]);
  const direct = directDelete({ from: 7, to: 9 });
  const results = [];
  for (const [first, second] of [
    [tracked, direct],
    [direct, tracked],
  ]) {
    if (first === undefined || second === undefined)
      panic("Both batches are present.");
    const applied = applyBatch(source, wire(first.ops)).unwrap();
    const sequenced = { ...first, revision: 1, effects: applied.effects };
    const moved = transformBatch(wire(second), [wire(sequenced)]).unwrap();
    const final = applyBatch(applied.document, moved.ops).unwrap();
    expect(
      applyBatch(final.document, wire(final.inverse)).unwrap().document,
    ).toEqual(applied.document);
    results.push(final.document);
  }
  expect(results.at(0)).toEqual(results.at(1));
  for (const decision of Object.values(REVISION_DECISIONS)) {
    const review = [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [deletion.revision.id],
        decision,
      },
    ] satisfies DocumentOp[];
    const left = results.at(0);
    const right = results.at(1);
    if (left === undefined || right === undefined)
      panic("Both accepted orders are present.");
    expect(applyBatch(left, review).unwrap().document).toEqual(
      applyBatch(right, review).unwrap().document,
    );
  }
});

test("fragment pools do not refuse disjoint or touching erasures", () => {
  const scenario = fixture({
    prefix: 4,
    suffix: 4,
    mirror: false,
    bold: false,
  });
  const deletion = scenario.tracked.ops.find(
    (op) => op.type === DOCUMENT_OP_TYPES.DELETE_RANGE,
  );
  if (deletion === undefined) panic("The replacement deletes a source span.");
  const directOps = [
    { ...directDelete({ from: 1, to: 2 }), opId: "disjoint" },
    directDelete({ from: deletion.from.offset - 1, to: deletion.from.offset }),
    directDelete({ from: deletion.to.offset, to: deletion.to.offset + 1 }),
    batch("other", [
      {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: { ...at(0), blockId: OTHER },
        to: { ...at(1), blockId: OTHER },
      },
    ]),
  ];
  for (const direct of directOps)
    for (const [first, second] of [
      [scenario.tracked, direct],
      [direct, scenario.tracked],
    ]) {
      if (first === undefined || second === undefined)
        panic("Both batches are present.");
      const applied = applyBatch(scenario.source, first.ops).unwrap();
      const moved = transformBatch(second, [
        { ...first, revision: 1, effects: applied.effects },
      ]);
      expect(moved.isOk()).toBe(true);
      if (moved.isErr()) throw moved.error;
      expect(applyBatch(applied.document, moved.value.ops).isOk()).toBe(true);
    }
});

test.each(Object.values(REVISION_DECISIONS))(
  "only actual dropped-created revision targets cascade to pending %s decisions",
  (decision) => {
    const scenario = fixture({
      prefix: 3,
      suffix: 3,
      mirror: false,
      bold: true,
    });
    const unused = scenario.extraId + 100;
    const tracked = {
      ...scenario.tracked,
      ops: scenario.tracked.ops.map((op) =>
        op.type === DOCUMENT_OP_TYPES.DELETE_RANGE
          ? {
              ...op,
              newIds: {
                ...op.newIds,
                revision: [...(op.newIds?.revision ?? []), unused],
              },
            }
          : op,
      ),
    };
    const authority = createSequencer(scenario.source, {
      idLeases: scenario.leases,
    });
    const peer = createClient(scenario.source);
    peer.enqueue(wire(tracked)).unwrap();
    expect(leasedIdCensus(peer.document).revision.has(unused)).toBe(false);
    const target = (id: number) =>
      ({
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [id],
        decision,
      }) satisfies DocumentOp;
    const other = (offset: number) => ({ ...at(offset), blockId: OTHER });
    const laterOps = suggested({
      document: peer.document,
      actor: "tracked",
      lease: scenario.trackedLease,
      intent: {
        type: "replaceText",
        from: other(0),
        to: other(0),
        text: "later",
      },
      serial: 5,
    });
    const laterInsertion = laterOps.find(
      (op) => op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT,
    );
    if (laterInsertion?.revision === undefined)
      panic("The later batch introduces a new revision.");
    peer
      .enqueue({
        ...batch("tracked", [target(scenario.firstId), ...laterOps]),
        opId: "atomic-dependent",
      })
      .unwrap();
    peer
      .enqueue({
        ...batch("tracked", [target(laterInsertion.revision.id)]),
        opId: "cascade",
      })
      .unwrap();
    const preexisting = scenario.standingId;
    peer
      .enqueue({
        ...batch("tracked", [target(preexisting)]),
        opId: "preexisting",
      })
      .unwrap();
    const otherParagraph = peer.document.package.document.content.find(
      (block) => block.type === "paragraph" && block.paraId === OTHER,
    );
    if (otherParagraph?.type !== "paragraph")
      panic("The unrelated typing has its paragraph.");
    peer
      .enqueue({
        ...batch("tracked", [
          {
            type: DOCUMENT_OP_TYPES.INSERT_TEXT,
            at: other(paragraphLength(otherParagraph)),
            text: "!",
            runProps: "inherit",
          },
        ]),
        opId: "unrelated",
      })
      .unwrap();
    peer
      .enqueue({ ...batch("tracked", [target(unused)]), opId: "unused-pool" })
      .unwrap();
    const [from, to] = scenario.ranges.interior;
    const submitted = authority.submit(
      wire(directDelete({ from: from, to: to })),
    );
    if (submitted.type !== "ack") throw submitted.reason;
    const first = authority.broadcasts.at(0);
    if (first === undefined) panic("The direct erase is sequenced.");
    const before = peer.document;
    const changes = peer.receiveBroadcast(wire(first)).unwrap();
    let replayed = before;
    for (const change of changes)
      replayed = applyBatch(replayed, wire(change.batch.ops)).unwrap().document;
    expect(replayed).toEqual(peer.document);
    expect(peer.notices.map(({ opId }) => opId)).toEqual([
      "tracked",
      "atomic-dependent",
      "cascade",
    ]);
    expect(peer.notices.slice(1).map(({ reason }) => reason.reason)).toEqual([
      BATCH_REJECTION_REASONS.CONFLICT,
      BATCH_REJECTION_REASONS.CONFLICT,
    ]);
    expect(peer.pending.map(({ opId }) => opId)).toEqual([
      "preexisting",
      "unrelated",
      "unused-pool",
    ]);
    expect(peer.pending.at(0)?.ops).toEqual([target(preexisting)]);
    expect(peer.pending.at(-1)?.ops).toEqual([target(unused)]);
    for (let index = 0; index < 3; index += 1) {
      const next = peer.nextSubmission();
      if (next === undefined)
        panic("Every independent pending batch survives.");
      const result = authority.submit(wire(next));
      if (result.type !== "ack") throw result.reason;
      for (const broadcast of authority.broadcasts.slice(peer.headRev))
        peer.receiveBroadcast(wire(broadcast)).unwrap();
      peer.receiveAck(result);
    }
    expect(peer.pending).toHaveLength(0);
    expect(peer.document).toEqual(authority.document);
    expect(authority.document).toEqual(
      applyBatch(
        applyBatch(
          scenario.source,
          directDelete({ from: from, to: to }).ops,
        ).unwrap().document,
        [
          target(preexisting),
          {
            type: DOCUMENT_OP_TYPES.INSERT_TEXT,
            at: other(5),
            text: "!",
            runProps: "inherit",
          },
          target(unused),
        ],
      ).unwrap().document,
    );
  },
);
