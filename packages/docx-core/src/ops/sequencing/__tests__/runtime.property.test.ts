import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../../../test/property-testing";
import type { Document, Paragraph } from "../../../model/document";
import { documentArbitrary, independentCopy } from "../../__tests__/documentArbitraries";
import { applyDocumentOps } from "../../apply";
import { storyParagraphs } from "../../blocks";
import { equalForStaleness } from "../../equality";
import { paragraphLength, paragraphLogicalText } from "../../offsets";
import { zeroWidthLeavesAt } from "../../leaves";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  type DocumentOp,
} from "../../types";
import { createClient } from "../client";
import { BatchRejection, type DocumentBatch, type SequencedBatch } from "../envelope";
import { createSequencer, type BatchSubmission } from "../sequencer";
import { transformBatch } from "../transform";

const FIRST = "70000001";
const SECOND = "70000002";
const REVIEW = "70000003";
const REVIEW_ID = 1_000_000;
const NUM_RUNS = 100;

const generatedDocument = documentArbitrary.map((source): Document => {
  const paragraphs: Paragraph[] = [
    {
      type: "paragraph",
      paraId: FIRST,
      content: [
        { type: "run", content: [{ type: "text", text: "ab😀cd" }, { type: "tab" }] },
        { type: "bookmarkStart", id: 1_000_000, name: "position" },
        { type: "bookmarkEnd", id: 1_000_000 },
        {
          type: "simpleField",
          instruction: "PAGE",
          fieldType: "PAGE",
          content: [{ type: "run", content: [{ type: "text", text: "123" }] }],
        },
      ],
    },
    {
      type: "paragraph",
      paraId: SECOND,
      content: [{ type: "run", content: [{ type: "text", text: "éxyz" }] }],
    },
    {
      type: "paragraph",
      paraId: REVIEW,
      content: [
        {
          type: "insertion",
          info: { id: REVIEW_ID, author: "review", date: "2026-01-01T00:00:00Z" },
          content: [{ type: "run", content: [{ type: "text", text: "reviewed" }] }],
        },
      ],
    },
  ];
  const { sections: _sections, ...body } = source.package.document;
  return Object.assign({}, source, {
    package: Object.assign({}, source.package, {
      document: Object.assign({}, body, { content: [...paragraphs, ...body.content] }),
    }),
  });
});

const GROUPS = [
  "differentBlocks",
  "insertInsert",
  "insertDelete",
  "deleteDelete",
  "trackedText",
  "trackedDeleteDirect",
  "runText",
  "runRun",
  "paragraphParagraph",
  "textSplit",
  "textJoin",
  "splitSplit",
  "insertBlocks",
  "resolveRevision",
] as const;
type Group = (typeof GROUPS)[number];

type OperationOptions = {
  document: Document;
  group: Group;
  role: number;
  seed: number;
  id: number;
  coordinate?: "all";
};
const operation = ({
  document,
  group,
  role,
  seed,
  id,
  coordinate,
}: OperationOptions): DocumentOp => {
  const blockId =
    (group === "differentBlocks" && role % 2 === 1) || (group === "textJoin" && role % 2 === 1)
      ? SECOND
      : FIRST;
  const paragraph = storyParagraphs(document.package.document).find(
    ({ paragraph: candidate }) => candidate.paraId === blockId,
  )?.paragraph;
  if (paragraph === undefined) {
    return {
      type: "setParagraphProps",
      story: OP_STORIES.MAIN,
      blockId,
      patch: { keepNext: true },
    };
  }
  const length = paragraphLength(paragraph);
  const text = paragraphLogicalText(paragraph);
  const boundaries = Array.from({ length: length + 1 }, (_, offset) => offset).filter(
    (offset) =>
      !(
        /[\uD800-\uDBFF]/u.test(text.charAt(offset - 1)) &&
        /[\uDC00-\uDFFF]/u.test(text.charAt(offset))
      ),
  );
  const positions =
    group === "textJoin" && role % 2 === 1 ? boundaries.filter((offset) => offset > 0) : boundaries;
  const offset = positions.at(seed % positions.length) ?? 0;
  const available = zeroWidthLeavesAt(paragraph.content, offset).length;
  const base = { story: OP_STORIES.MAIN, blockId, offset };
  const at =
    seed % 3 === 0
      ? { ...base, zeroWidthBefore: coordinate === "all" ? seed % (available + 1) : 0 }
      : base;
  const from = {
    story: OP_STORIES.MAIN,
    blockId,
    offset: 0,
    ...(seed % 3 === 0 ? { zeroWidthBefore: 0 } : {}),
  };
  const to = { ...from, offset: length };
  const insert = {
    type: "insertText",
    at,
    text: role % 2 === 0 ? "X" : "ü",
    runProps: INHERIT_RUN_PROPS,
  } as const;
  switch (group) {
    case "differentBlocks":
    case "insertInsert":
      return insert;
    case "insertDelete":
      return role % 2 === 0 ? insert : { type: "deleteRange", from, to };
    case "deleteDelete":
      return { type: "deleteRange", from: { ...from, offset: seed % 2 }, to };
    case "trackedText":
      return role % 2 === 0
        ? {
            type: "deleteRange",
            from,
            to: { ...to, offset: Math.min(1, length) },
            revision: { id: 2_000_000 + id, author: "tracked", date: "2026-01-01T00:00:00Z" },
          }
        : { ...insert, at: { ...at, offset: length } };
    case "trackedDeleteDirect":
      return role % 2 === 0
        ? {
            type: "deleteRange",
            from,
            to: { ...to, offset: 1 },
            revision: { id: 2_000_000 + id, author: "tracked", date: "2026-01-01T00:00:00Z" },
          }
        : { type: "deleteRange", from, to: { ...to, offset: 1 } };
    case "runText":
      return role % 2 === 0 ? { type: "setRunProps", from, to, patch: { bold: true } } : insert;
    case "runRun":
      if (role % 2 === 0) return { type: "setRunProps", from, to, patch: { bold: true } };
      if (seed % 2 === 0) return { type: "setRunProps", from, to, patch: { bold: false } };
      return { type: "setRunProps", from, to, patch: { italic: true } };
    case "paragraphParagraph":
      return {
        type: "setParagraphProps",
        story: OP_STORIES.MAIN,
        blockId,
        patch: role % 2 === 0 ? { keepNext: true } : { alignment: "center" },
      };
    case "textSplit":
      return role % 2 === 0
        ? {
            type: "splitBlock",
            at,
            newBlockId: (0x71000000 + id).toString(16).toUpperCase(),
            newHalf: "second",
          }
        : insert;
    case "textJoin":
      return role % 2 === 0
        ? {
            type: "joinBlocks",
            story: OP_STORIES.MAIN,
            blockId: FIRST,
            nextBlockId: SECOND,
            survivor: "second",
          }
        : insert;
    case "splitSplit":
      return {
        type: "splitBlock",
        at,
        newBlockId: (0x71000000 + id).toString(16).toUpperCase(),
        newHalf: "second",
      };
    case "insertBlocks":
      return {
        type: "insertBlocks",
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId: FIRST },
        blocks: [
          {
            type: "paragraph",
            paraId: (0x71000000 + id).toString(16).toUpperCase(),
            content: [{ type: "run", content: [{ type: "text", text: "inserted" }] }],
          },
        ],
      };
    case "resolveRevision":
      return {
        type: "resolveRevision",
        story: OP_STORIES.MAIN,
        revisionIds: [REVIEW_ID],
        decision: "accept",
      };
    default: {
      const exhaustive: never = group;
      return exhaustive;
    }
  }
};

type Delivery =
  | { client: number; type: "result"; result: BatchSubmission }
  | { client: number; type: "broadcast"; batch: SequencedBatch };
type SimulationOptions = {
  document: Document;
  groups: readonly Group[];
  clientsCount: number;
  schedule: readonly number[];
  enqueueLater: boolean;
};
const simulate = ({
  document,
  groups,
  clientsCount,
  schedule,
  enqueueLater,
}: SimulationOptions) => {
  const sequencer = createSequencer(independentCopy(document));
  const clients = Array.from({ length: clientsCount }, () =>
    createClient(independentCopy(document)),
  );
  const submitted = new Set<string>();
  const admitted = new Set<string>();
  const rejected = new Map<string, { client: number }>();
  const rejectionReasons = new Map<string, BatchRejection>();
  const delivery: Delivery[] = [];
  let id = 1;
  let scheduleIndex = 0;
  let accepted = 0;
  const choice = () => schedule.at(scheduleIndex++ % schedule.length) ?? 0;
  for (const [index, client] of clients.entries()) {
    const group = groups.at(index % groups.length) ?? "insertInsert";
    const batch: DocumentBatch = {
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      opId: `operation-${id}`,
      actor: `actor-${index}`,
      baseRev: client.headRev,
      ops: [
        operation({
          document: client.document,
          group,
          role: index,
          seed: choice(),
          id: id++,
          ...(enqueueLater ? { coordinate: "all" as const } : {}),
        }),
      ],
    };
    expect(client.enqueue(batch).isOk()).toBe(true);
    admitted.add(batch.opId);
    if (enqueueLater) {
      const later: DocumentBatch = {
        ...batch,
        opId: `operation-${id}`,
        ops: [
          operation({
            document: client.document,
            group: "differentBlocks",
            role: 1,
            seed: choice(),
            id: id++,
          }),
        ],
      };
      expect(client.enqueue(later).isOk()).toBe(true);
      admitted.add(later.opId);
    }
  }
  const submit = (index: number): void => {
    const client = clients.at(index);
    const batch = client?.nextSubmission();
    if (batch === undefined) return;
    const before = sequencer.broadcasts.length;
    const result = sequencer.submit(batch);
    submitted.add(batch.opId);
    expect(sequencer.submit(batch)).toEqual(result);
    expect(sequencer.broadcasts.length).toBe(before + (result.type === "ack" ? 1 : 0));
    delivery.push({ client: index, type: "result", result });
    if (result.type === "ack") {
      accepted += 1;
      const broadcast = sequencer.broadcasts.at(-1);
      expect(broadcast?.revision).toBe(result.rev);
      if (broadcast !== undefined)
        for (let target = 0; target < clientsCount; target += 1)
          delivery.push({ client: target, type: "broadcast", batch: broadcast });
    } else expect(result.headRev).toBe(sequencer.headRev);
    if (result.type === "reject") rejected.set(batch.opId, { client: index });
  };
  const order = clients
    .map((_, index) => ({ index, priority: choice() }))
    .sort((a, b) => a.priority - b.priority);
  for (const { index } of order) submit(index);
  let steps = 0;
  while (delivery.length > 0) {
    expect(steps++).toBeLessThan(1000);
    const event = delivery.splice(choice() % delivery.length, 1).at(0);
    if (event === undefined) continue;
    const client = clients.at(event.client);
    if (client === undefined) continue;
    if (event.type === "broadcast") {
      const received = client.receiveBroadcast(event.batch);
      expect(received.isOk()).toBe(true);
      expect(client.receiveBroadcast(event.batch).isOk()).toBe(true);
    } else if (event.result.type === "ack") client.receiveAck(event.result);
    else {
      if (client.pending.some(({ opId }) => opId === event.result.opId))
        rejectionReasons.set(event.result.opId, event.result.reason);
      client.receiveReject(event.result);
      const noticeCount = client.notices.length;
      client.receiveReject(event.result);
      expect(client.notices).toHaveLength(noticeCount);
    }
    for (let index = 0; index < clientsCount; index += 1) submit(index);
  }
  const sequencedIds = new Set(sequencer.broadcasts.map(({ opId }) => opId));
  const droppedIds: string[] = [];
  for (const [index, client] of clients.entries()) {
    expect(client.headRev).toBe(sequencer.headRev);
    expect(client.pending).toHaveLength(0);
    expect(equalForStaleness(client.document, sequencer.document)).toBe(true);
    for (const notice of client.notices) {
      expect(notice.ops).toBeDefined();
      droppedIds.push(notice.opId);
    }
    for (const [opId, rejection] of rejected) {
      if (rejection.client !== index) continue;
      const notices = client.notices.filter((notice) => notice.opId === opId);
      expect(notices).toHaveLength(1);
      const reason = rejectionReasons.get(opId);
      if (reason !== undefined) expect(notices.at(0)?.reason).toBe(reason);
    }
  }
  expect(new Set(droppedIds).size).toBe(droppedIds.length);
  expect(new Set([...sequencedIds, ...droppedIds])).toEqual(admitted);
  for (const opId of rejected.keys()) expect(droppedIds).toContain(opId);
  return { accepted, submitted: submitted.size };
};

const scheduleArbitrary = fc.array(fc.nat({ max: 100_000 }), { minLength: 10, maxLength: 100 });

describe("reference sequencing convergence", () => {
  for (const group of GROUPS) {
    test(
      group,
      () => {
        let accepted = 0;
        assertProperty(
          fc.property(
            generatedDocument,
            fc.integer({ min: 2, max: 4 }),
            scheduleArbitrary,
            (document, clientsCount, schedule) => {
              const result = simulate({
                document,
                groups: [group],
                clientsCount,
                schedule,
                enqueueLater: false,
              });
              accepted += result.accepted;
              expect(result.submitted).toBe(clientsCount);
              expect(result.accepted).toBeGreaterThanOrEqual(2);
            },
          ),
          { numRuns: NUM_RUNS },
        );
        expect(accepted).toBeGreaterThanOrEqual(NUM_RUNS * 2);
      },
      propertyTestTimeout(30_000),
    );
  }

  test(
    "mixed operations with causal pending queues",
    () => {
      assertProperty(
        fc.property(
          generatedDocument,
          fc.integer({ min: 2, max: 4 }),
          fc.array(fc.constantFrom(...GROUPS), { minLength: 2, maxLength: 4 }),
          scheduleArbitrary,
          (document, clientsCount, groups, schedule) => {
            const result = simulate({
              document,
              groups,
              clientsCount,
              schedule,
              enqueueLater: true,
            });
            expect(result.accepted).toBeGreaterThan(0);
          },
        ),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(30_000),
  );

  test(
    "inverse submissions over foreign edits succeed or are dropped atomically",
    () => {
      let accepted = 0;
      let dropped = 0;
      assertProperty(
        fc.property(generatedDocument, fc.nat(), (document, seed) => {
          const forward = operation({ document, group: "insertInsert", role: 0, seed, id: 1 });
          const applied = applyDocumentOps(document, [forward]);
          expect(applied.isOk()).toBe(true);
          if (applied.isErr()) return;
          const sequencer = createSequencer(document);
          const batch: DocumentBatch = {
            schema: DOCUMENT_OP_SCHEMA_VERSION,
            opId: "forward",
            actor: "first",
            baseRev: 0,
            ops: [forward],
          };
          expect(sequencer.submit(batch).type).toBe("ack");
          const foreign: DocumentBatch = {
            ...batch,
            opId: "foreign",
            actor: "second",
            baseRev: 1,
            ops: [
              operation({
                document: sequencer.document,
                group: seed % 2 === 0 ? "differentBlocks" : "insertInsert",
                role: 1,
                seed: seed + 1,
                id: 2,
              }),
            ],
          };
          expect(sequencer.submit(foreign).type).toBe("ack");
          const before = sequencer.document;
          const inverse: DocumentBatch = {
            ...batch,
            opId: "inverse",
            baseRev: 1,
            ops: applied.value.inverse,
          };
          const transformed = transformBatch(inverse, sequencer.broadcasts.slice(1));
          const result = sequencer.submit(inverse);
          if (result.type === "reject") {
            dropped += 1;
            expect(equalForStaleness(sequencer.document, before)).toBe(true);
          } else {
            accepted += 1;
            expect(transformed.isOk()).toBe(true);
            if (transformed.isOk())
              expect(applyDocumentOps(before, transformed.value.ops).isOk()).toBe(true);
          }
        }),
        { numRuns: NUM_RUNS },
      );
      expect(accepted).toBeGreaterThan(0);
      expect(dropped).toBeGreaterThan(0);
    },
    propertyTestTimeout(30_000),
  );

  test(
    "later pending insertions retain their point after a foreign join",
    () => {
      const textArbitrary = fc
        .array(fc.constantFrom("X", "é", "😀"), { minLength: 1, maxLength: 3 })
        .map((parts) => parts.join(""));
      assertProperty(
        fc.property(
          generatedDocument,
          textArbitrary,
          textArbitrary,
          (document, firstText, secondText) => {
            const client = createClient(document);
            const sequencer = createSequencer(document);
            const firstParagraph = storyParagraphs(document.package.document).find(
              ({ paragraph }) => paragraph.paraId === FIRST,
            )?.paragraph;
            const secondParagraph = storyParagraphs(document.package.document).find(
              ({ paragraph }) => paragraph.paraId === SECOND,
            )?.paragraph;
            expect(firstParagraph).toBeDefined();
            expect(secondParagraph).toBeDefined();
            if (firstParagraph === undefined || secondParagraph === undefined) return;
            const firstBefore = paragraphLogicalText(firstParagraph);
            const secondBefore = paragraphLogicalText(secondParagraph);
            const first: DocumentBatch = {
              schema: DOCUMENT_OP_SCHEMA_VERSION,
              opId: "pending-first",
              actor: "first",
              baseRev: 0,
              ops: [
                {
                  type: "insertText",
                  at: { story: OP_STORIES.MAIN, blockId: FIRST, offset: 1 },
                  text: firstText,
                  runProps: INHERIT_RUN_PROPS,
                },
              ],
            };
            const second: DocumentBatch = {
              ...first,
              opId: "pending-second",
              ops: [
                {
                  type: "insertText",
                  at: { story: OP_STORIES.MAIN, blockId: SECOND, offset: 1 },
                  text: secondText,
                  runProps: INHERIT_RUN_PROPS,
                },
              ],
            };
            expect(client.enqueue(first).isOk()).toBe(true);
            expect(client.enqueue(second).isOk()).toBe(true);
            const joined: DocumentBatch = {
              ...first,
              opId: "foreign-join",
              actor: "foreign",
              ops: [
                {
                  type: "joinBlocks",
                  story: OP_STORIES.MAIN,
                  blockId: FIRST,
                  nextBlockId: SECOND,
                  survivor: "second",
                },
              ],
            };
            expect(sequencer.submit(joined).type).toBe("ack");
            const broadcast = sequencer.broadcasts.at(0);
            expect(broadcast).toBeDefined();
            if (broadcast === undefined) return;
            expect(client.receiveBroadcast(broadcast).isOk()).toBe(true);
            const after = storyParagraphs(client.document.package.document).find(
              ({ paragraph }) => paragraph.paraId === SECOND,
            )?.paragraph;
            expect(after).toBeDefined();
            if (after === undefined) return;
            expect(paragraphLogicalText(after)).toBe(
              `${firstBefore.slice(0, 1)}${firstText}${firstBefore.slice(1)}${secondBefore.slice(0, 1)}${secondText}${secondBefore.slice(1)}`,
            );
            expect(client.pending).toHaveLength(2);
          },
        ),
        { numRuns: NUM_RUNS },
      );
    },
    propertyTestTimeout(30_000),
  );
});

test("rejection preserves refused text and removes optimistic state", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: FIRST,
            content: [{ type: "run", content: [{ type: "text", text: "before" }] }],
          },
        ],
      },
    },
  };
  const client = createClient(document);
  const batch: DocumentBatch = {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "rejected",
    actor: "first",
    baseRev: 0,
    ops: [
      {
        type: "insertText",
        at: { story: OP_STORIES.MAIN, blockId: FIRST, offset: 3 },
        text: "copy me",
        runProps: INHERIT_RUN_PROPS,
      },
    ],
  };
  expect(client.enqueue(batch).isOk()).toBe(true);
  client.nextSubmission();
  client.receiveReject({
    type: "reject",
    opId: batch.opId,
    headRev: 0,
    reason: new BatchRejection({ reason: "conflict", message: "Refused edit." }),
  });
  expect(equalForStaleness(client.document, document)).toBe(true);
  expect(client.notices.at(0)?.ops).toEqual(batch.ops);
});

test("an insertion inside a tracked deletion is refused without altering the journal", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: FIRST,
            content: [{ type: "run", content: [{ type: "text", text: "abcd" }] }],
          },
        ],
      },
    },
  };
  const sequencer = createSequencer(document);
  const at = { story: OP_STORIES.MAIN, blockId: FIRST, offset: 0 };
  const deleted: DocumentBatch = {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "tracked-delete",
    actor: "first",
    baseRev: 0,
    ops: [
      {
        type: "deleteRange",
        from: at,
        to: { ...at, offset: 2 },
        revision: { id: REVIEW_ID, author: "first", date: "2026-01-01T00:00:00Z" },
      },
    ],
  };
  expect(sequencer.submit(deleted).type).toBe("ack");
  const before = sequencer.document;
  const inserted: DocumentBatch = {
    ...deleted,
    opId: "inside-delete",
    actor: "second",
    ops: [
      {
        type: "insertText",
        at: { ...at, offset: 1 },
        text: "copy me",
        runProps: INHERIT_RUN_PROPS,
      },
    ],
  };
  expect(sequencer.submit(inserted).type).toBe("reject");
  expect(equalForStaleness(sequencer.document, before)).toBe(true);
  expect(sequencer.headRev).toBe(1);
});

test("acknowledgment before broadcast preserves the optimistic projection and drops pending", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: FIRST,
            content: [{ type: "run", content: [{ type: "text", text: "ab" }] }],
          },
        ],
      },
    },
  };
  const sequencer = createSequencer(document);
  const client = createClient(document);
  const batch: DocumentBatch = {
    schema: DOCUMENT_OP_SCHEMA_VERSION,
    opId: "ack-first",
    actor: "first",
    baseRev: 0,
    ops: [
      {
        type: "insertText",
        at: { story: OP_STORIES.MAIN, blockId: FIRST, offset: 1 },
        text: "X",
        runProps: INHERIT_RUN_PROPS,
      },
    ],
  };
  expect(client.enqueue(batch).isOk()).toBe(true);
  const submission = client.nextSubmission();
  expect(submission).toBeDefined();
  if (submission === undefined) return;
  const outcome = sequencer.submit(submission);
  expect(outcome.type).toBe("ack");
  if (outcome.type !== "ack") return;
  const optimistic = client.document;
  client.receiveAck(outcome);
  expect(client.pending).toHaveLength(0);
  expect(client.document).toBe(optimistic);
  expect(client.nextSubmission()).toBeUndefined();
  const broadcast = sequencer.broadcasts.at(0);
  expect(broadcast).toBeDefined();
  if (broadcast === undefined) return;
  expect(client.receiveBroadcast(broadcast).isOk()).toBe(true);
  expect(equalForStaleness(client.document, sequencer.document)).toBe(true);
  expect(client.receiveBroadcast({ ...broadcast, opId: "conflicting-revision" }).isErr()).toBe(
    true,
  );
});
