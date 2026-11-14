/**
 * Two peers in one room: one removes a comment whose range opens between
 * blocks, or undoes that removal, while the other splits, joins, deletes or
 * types into the paragraph the marker stands beside. Through the real
 * sequencer and optimistic clients, every peer ends on the authority's
 * document, the document keeps whole comment ranges, a restoration either
 * puts the marker back beside the text it opened on or is dropped with a
 * notice, and nothing is lost silently.
 */
import { describe, expect, test } from "bun:test";

import { setDocumentBodyComments } from "../../model/canonical";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { storyParagraphs } from "../blocks";
import { commentAnchorsIn, commentDocumentIssue } from "../comments";
import { equalForStaleness } from "../equality";
import {
  formatParaId,
  freshLeasedIds,
  ID_LEASE_KINDS,
  type IdLease,
} from "../idLease";
import { applyBatch } from "../sequencing/applyBatch";
import { createClient } from "../sequencing/client";
import type { DocumentBatch } from "../sequencing/envelope";
import { createSequencer } from "../sequencing/sequencer";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "../types";

const OPENING = "00000001";
const COMMENTED = "00000002";
const CLOSING = "00000003";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (paraId: string, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});

const PLACEMENTS = {
  "between paragraphs": (commented: Paragraph): BlockContent[] => [
    { type: "commentRangeStart", id: 1 },
    commented,
  ],
  "beside a content control": (commented: Paragraph): BlockContent[] => [
    { type: "commentRangeStart", id: 1, displacedByCustomXml: "next" },
    { type: "blockSdt", properties: {}, content: [commented] },
  ],
} as const;

const base = (placement: keyof typeof PLACEMENTS): Document => {
  const document: Document = {
    package: {
      document: {
        content: [
          paragraph(OPENING, [run("Opening words")]),
          ...PLACEMENTS[placement](
            paragraph(COMMENTED, [
              run("Please remit funds."),
              { type: "commentRangeEnd", id: 1 },
              { type: "commentReference", id: 1 },
            ]),
          ),
          paragraph(CLOSING, [run("Closing words")]),
        ],
      },
    },
  };
  setDocumentBodyComments(document.package.document, [
    {
      id: 1,
      author: "Reviewer",
      date: "2024-01-01T00:00:00Z",
      content: [paragraph("00000011", [run("Check the account")])],
    },
  ]);
  return document;
};

/** The neighbour's edits, each compiled against the peer's document. */
const EDITS: Readonly<Record<string, (lease: IdLease) => DocumentOp[]>> = {
  split: (lease: IdLease): DocumentOp[] => [
    {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: { story: OP_STORIES.MAIN, blockId: COMMENTED, offset: 6 },
      newBlockId: formatParaId(
        freshLeasedIds({
          lease,
          kind: ID_LEASE_KINDS.PARAGRAPH,
          count: 1,
          taken: () => false,
        }).at(0) ?? 0,
      ),
    },
  ],
  join: (): DocumentOp[] => [
    {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: OPENING,
      nextBlockId: COMMENTED,
    },
  ],
  delete: (): DocumentOp[] => [
    {
      type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
      story: OP_STORIES.MAIN,
      blockIds: [COMMENTED],
    },
  ],
  type: (): DocumentOp[] => [
    {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: { story: OP_STORIES.MAIN, blockId: COMMENTED, offset: 0 },
      text: "Now ",
      runProps: INHERIT_RUN_PROPS,
    },
  ],
};

const wire = <Value>(value: Value): Value => JSON.parse(JSON.stringify(value));

const room = (document: Document) => {
  const sequencer = createSequencer(document);
  const actors = ["remover", "neighbour"] as const;
  const leases = actors.map((actor) => sequencer.grantIdLease(actor).unwrap());
  const clients = actors.map(() => createClient(document));
  let serial = 0;
  const enqueued: string[] = [];
  /** Queue an edit; its inverse, or undefined when the peer refuses it. */
  const enqueue = (
    peer: 0 | 1,
    ops: readonly DocumentOp[],
  ): readonly DocumentOp[] | undefined => {
    const client = clients[peer];
    if (client === undefined) return undefined;
    const before = client.document;
    const batch: DocumentBatch = {
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      opId: `op-${serial++}`,
      actor: actors[peer],
      baseRev: client.headRev,
      ops,
    };
    const queued = client.enqueue(wire(batch));
    if (queued.isErr()) return undefined;
    enqueued.push(batch.opId);
    const applied = applyBatch(before, queued.value.ops);
    return applied.isOk() ? applied.value.inverse : undefined;
  };
  const submit = (peer: 0 | 1): void => {
    const batch = clients[peer]?.nextSubmission();
    if (batch === undefined) return;
    const result = sequencer.submit(wire(batch));
    const sender = clients[peer];
    if (result.type === "ack") sender?.receiveAck(result);
    else sender?.receiveReject(result);
    const broadcast = sequencer.broadcasts.at(-1);
    if (result.type !== "ack" || broadcast === undefined) return;
    for (const client of clients) {
      const received = client.receiveBroadcast(wire(broadcast));
      if (received.isErr()) throw received.error;
    }
  };
  const drain = (order: readonly (0 | 1)[]): void => {
    let rounds = 0;
    while (clients.some(({ pending }) => pending.length > 0)) {
      rounds += 1;
      expect(rounds).toBeLessThan(20);
      for (const peer of order) submit(peer);
    }
  };
  return { sequencer, clients, leases, enqueue, drain, enqueued };
};

/** The paragraph holding a comment's start, by its text. */
const startText = (document: Document): string | undefined => {
  const start = commentAnchorsIn(document).find(
    ({ node }) => node.type === "commentRangeStart" && node.id === 1,
  );
  if (start === undefined) return undefined;
  const found = storyParagraphs(document.package.document).find(
    ({ paragraph: value }) => value.paraId === start.blockId,
  );
  return JSON.stringify(found?.paragraph.content ?? []);
};

const expectConverged = (current: ReturnType<typeof room>): Document => {
  const authority = current.sequencer.document;
  for (const client of current.clients)
    expect(equalForStaleness(client.document, authority)).toBe(true);
  expect(commentDocumentIssue(authority, "all")).toBeUndefined();
  // An edit the authority did not sequence comes back as a notice.
  const sequenced = new Set(
    current.sequencer.broadcasts.map(({ opId }) => opId),
  );
  const noticed = new Set(
    current.clients.flatMap(({ notices }) => notices.map(({ opId }) => opId)),
  );
  for (const opId of current.enqueued)
    expect(sequenced.has(opId) || noticed.has(opId)).toBe(true);
  return authority;
};

const hasComment = (document: Document): boolean =>
  (document.package.document.comments ?? []).some(({ id }) => id === 1);

describe("a comment opening between blocks under concurrent edits beside it", () => {
  for (const placement of Object.keys(
    PLACEMENTS,
  ) as (keyof typeof PLACEMENTS)[])
    for (const [edit, compile] of Object.entries(EDITS))
      for (const order of [
        [0, 1],
        [1, 0],
      ] as const) {
        test(`${placement}: removal beside a concurrent ${edit}, ${order[0] === 0 ? "removal" : edit} first`, () => {
          const current = room(base(placement));
          current.enqueue(0, [
            { type: DOCUMENT_OP_TYPES.DELETE_COMMENT, id: 1, scope: "thread" },
          ]);
          current.enqueue(1, compile(current.leases[1] ?? { blocks: [] }));
          current.drain(order);
          const authority = expectConverged(current);
          // A removal that was sequenced leaves no marker of the comment.
          if (!hasComment(authority))
            expect(
              commentAnchorsIn(authority).filter(
                ({ node }) => "id" in node && node.id === 1,
              ),
            ).toEqual([]);
        });

        test(`${placement}: undoing the removal beside a concurrent ${edit}, ${order[0] === 0 ? "undo" : edit} first`, () => {
          const current = room(base(placement));
          const undo = current.enqueue(0, [
            { type: DOCUMENT_OP_TYPES.DELETE_COMMENT, id: 1, scope: "thread" },
          ]);
          current.drain([0, 1]);
          expect(hasComment(current.sequencer.document)).toBe(false);
          expect(undo).toBeDefined();
          current.enqueue(0, undo ?? []);
          current.enqueue(1, compile(current.leases[1] ?? { blocks: [] }));
          current.drain(order);
          const authority = expectConverged(current);
          // Restored, the start stands beside the text it opened on; not
          // restored, no marker of the comment is left behind.
          if (hasComment(authority))
            expect(startText(authority)).toContain("Please");
          else
            expect(
              commentAnchorsIn(authority).filter(
                ({ node }) => "id" in node && node.id === 1,
              ),
            ).toEqual([]);
        });
      }
});
