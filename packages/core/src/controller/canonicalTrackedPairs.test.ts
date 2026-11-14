import { panic } from "better-result";
import { describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";

import { serializeDocumentToDocx } from "@stll/docx-core";
import {
  applyDocumentOps,
  type ClientDocumentChange,
  createClient,
  createSequencer,
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentBatch,
  type DocumentOp,
  ID_LEASE_KINDS,
  type IdLease,
  OP_STORIES,
  paragraphLogicalText,
  type SequencedBatch,
  storyParagraphs,
  type TextPosition,
} from "@stll/docx-core/ops";

import {
  type CanonicalHistoryDrop,
  type CanonicalSelection,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";

/*
 * Concurrent tracked changes of one paragraph whose two orders record other
 * cuts, or leave shared content to another deletion, are accepted: the
 * authority's order decides, every peer replays to it, and a peer that
 * applied its own edit first gets the document as a replacement.
 */

const TEXT = "Shared paragraph text here";
const DATE = "2026-01-01T00:00:00Z";

const seed = async (): Promise<Uint8Array> =>
  (
    await prepareCanonicalDocxInput(
      await createDocx({
        package: {
          document: {
            content: [
              {
                type: "paragraph",
                content: [
                  { type: "run", content: [{ type: "text", text: TEXT }] },
                ],
              },
            ],
          },
        },
      }),
    )
  ).unwrap();

const open = async (bytes: Uint8Array): Promise<Document> =>
  parseDocx(bytes, { preloadFonts: false, detectVariables: false });

const firstParagraph = (document: Document) =>
  storyParagraphs(document.package.document).at(0)?.paragraph ??
  panic("The document has a paragraph.");

const revisionId = (lease: IdLease): number =>
  lease.blocks.find(({ kind }) => kind === ID_LEASE_KINDS.REVISION)?.start ??
  panic("The lease holds revision ids.");

const at = (blockId: string, offset: number): TextPosition => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

const trackedDeletion = (
  blockId: string,
  [from, to]: readonly [number, number],
  id: number,
  author: string,
): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_RANGE,
  from: at(blockId, from),
  to: at(blockId, to),
  revision: { id, author, date: DATE },
});

const batch = (
  actor: string,
  opId: string,
  baseRev: number,
  ops: readonly DocumentOp[],
): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor,
  baseRev,
  ops,
});

/** The text a paragraph shows once every tracked change left in it is accepted. */
const acceptedText = (document: Document, ids: readonly number[]): string => {
  const accepted =
    ids.length === 0
      ? document
      : applyDocumentOps(document, [
          {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: [...ids],
            decision: "accept",
          },
        ]).unwrap().document;
  return paragraphLogicalText(firstParagraph(accepted));
};

/** The ids of the tracked deletions a paragraph holds, at any depth. */
const deletionIds = (document: Document): Set<number> => {
  const ids = new Set<number>();
  const visit = (nodes: readonly unknown[]): void => {
    for (const node of nodes) {
      if (typeof node !== "object" || node === null) continue;
      if (
        "type" in node &&
        node.type === "deletion" &&
        "info" in node &&
        typeof node.info === "object" &&
        node.info !== null &&
        "id" in node.info &&
        typeof node.info.id === "number"
      )
        ids.add(node.info.id);
      if ("content" in node && Array.isArray(node.content)) visit(node.content);
    }
  };
  visit(firstParagraph(document).content);
  return ids;
};

/** `TEXT` without the characters of the given ranges. */
const without = (ranges: readonly (readonly [number, number])[]): string =>
  [...TEXT]
    .filter(
      (_, index) => !ranges.some(([from, to]) => from <= index && index < to),
    )
    .join("");

/** The ranges of the second deletion the first one does not hold. */
const outside = (
  [from, to]: readonly [number, number],
  [heldFrom, heldTo]: readonly [number, number],
): (readonly [number, number])[] =>
  [
    [from, Math.min(to, heldFrom)] as const,
    [Math.max(from, heldTo), to] as const,
  ].filter(([start, end]) => start < end);

const OVERLAPS = {
  nested: [
    [2, 12],
    [5, 8],
  ],
  equal: [
    [3, 9],
    [3, 9],
  ],
  crossing: [
    [2, 9],
    [6, 14],
  ],
} as const satisfies Record<
  string,
  readonly [readonly [number, number], readonly [number, number]]
>;

describe("overlapping tracked deletions", () => {
  for (const [name, [first, second]] of Object.entries(OVERLAPS))
    test(`${name}: the deletion sequenced first holds the shared content, and every decision saves and reopens`, async () => {
      const base = createCanonicalSession(await open(await seed())).unwrap()
        .document;
      const blockId = firstParagraph(base).paraId ?? panic("Identified.");
      const room = createSequencer(base);
      const actors = ["first", "second", "watcher"];
      const leases = actors.map((actor) => room.grantIdLease(actor).unwrap());
      const clients = actors.map(() => createClient(base));
      const ids = leases.slice(0, 2).map(revisionId);
      // Both suggest at once; neither has seen the other's deletion.
      for (const [index, range] of [first, second].entries()) {
        const actor = actors[index] ?? panic("An actor per deletion.");
        const id = ids[index] ?? panic("A revision id per deletion.");
        const client = clients[index] ?? panic("A client per actor.");
        client
          .enqueue(
            batch(actor, `${actor}-0`, 0, [
              trackedDeletion(blockId, range, id, actor),
            ]),
          )
          .unwrap();
      }
      const broadcasts: SequencedBatch[] = [];
      for (const client of clients.slice(0, 2)) {
        const submitted = client.nextSubmission() ?? panic("A batch to send.");
        const answer = room.submit(submitted);
        // Neither is refused: the later one wraps only what is not deleted.
        expect(answer.type).toBe("ack");
        if (answer.type === "ack") client.receiveAck(answer);
        broadcasts.push(room.broadcasts.at(-1) ?? panic("Sequenced."));
      }
      for (const client of clients)
        for (const sequenced of broadcasts)
          client.receiveBroadcast(sequenced).unwrap();
      for (const client of clients) {
        expect(client.pending).toHaveLength(0);
        expect(client.notices).toEqual([]);
        expect(client.document).toEqual(room.document);
      }
      const [firstId, secondId] = ids;
      if (firstId === undefined || secondId === undefined)
        return panic("Two revision ids.");
      // Accepting everything removes both ranges; the shared content went to
      // the first deletion, so rejecting it alone keeps the shared content.
      const secondOnly = outside(second, first);
      // A deletion inside one sequenced before it has nothing left to wrap.
      const present = deletionIds(room.document);
      expect(present.has(firstId)).toBe(true);
      expect(present.has(secondId)).toBe(secondOnly.length > 0);
      const decisions = [
        { decision: "accept", revisionIds: [firstId] },
        { decision: "accept", revisionIds: [secondId] },
        { decision: "accept", revisionIds: [firstId, secondId] },
        { decision: "reject", revisionIds: [firstId] },
        { decision: "reject", revisionIds: [secondId] },
        { decision: "reject", revisionIds: [firstId, secondId] },
      ] as const;
      const removedOnAcceptingTheRest = {
        accept: [first, ...secondOnly],
        reject: {
          first: secondOnly,
          second: [first],
          both: [],
        },
      };
      for (const { decision, revisionIds } of decisions) {
        if (!revisionIds.every((id) => present.has(id))) continue;
        const decided = applyDocumentOps(room.document, [
          {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: [...revisionIds],
            decision,
          },
        ]).unwrap().document;
        const left = ids.filter(
          (id) => present.has(id) && !revisionIds.includes(id),
        );
        const rejected = (): (readonly [number, number])[] => {
          if (revisionIds.length === 2)
            return removedOnAcceptingTheRest.reject.both;
          return revisionIds[0] === firstId
            ? removedOnAcceptingTheRest.reject.first
            : removedOnAcceptingTheRest.reject.second;
        };
        const expected =
          decision === "accept" ? removedOnAcceptingTheRest.accept : rejected();
        expect(acceptedText(decided, left)).toBe(without(expected));
        // It saves, and the saved file reopens to the same text.
        const reopened = createCanonicalSession(
          await open(await serializeDocumentToDocx(decided)),
        ).unwrap().document;
        expect(acceptedText(reopened, left)).toBe(without(expected));
      }
    });
});

describe("a replaced change in a session", () => {
  test("drops the local history through the history drop signal and maps the selection", async () => {
    const bytes = await seed();
    const parsed = createCanonicalSession(await open(bytes)).unwrap().document;
    const room = createSequencer(parsed);
    const sessions = await Promise.all(
      ["local", "remote"].map(async (actor) =>
        createCanonicalSession(await open(bytes), undefined, {
          idLease: room.grantIdLease(actor).unwrap(),
        }).unwrap(),
      ),
    );
    const [local, remote] = sessions;
    if (local === undefined || remote === undefined)
      return panic("Two sessions.");
    const blockId =
      firstParagraph(local.document).paraId ?? panic("Identified.");
    const state = (session: typeof local, offset: number) => {
      const doc = session.projection.doc;
      return EditorState.create({
        doc,
        selection: TextSelection.create(
          doc,
          session.projection.positionAt(at(blockId, offset)).unwrap(),
        ),
      });
    };
    const caret = (offset: number): CanonicalSelection => ({
      type: "text",
      anchor: at(blockId, offset),
      head: at(blockId, offset),
    });
    const drops: CanonicalHistoryDrop[] = [];
    local.onHistoryDrop((drop) => drops.push(drop));
    const localClient = createClient(local.document);
    const leaseOf = (actor: string) => room.grantIdLease(actor).unwrap();

    // The local peer suggests deleting 4..5 and keeps its caret at 16; the
    // edit stays pending.
    let localState = state(local, 16);
    const localCommit = local
      .prepareOps(
        localState,
        [
          trackedDeletion(
            blockId,
            [4, 5],
            revisionId(leaseOf("local")),
            "local",
          ),
        ],
        caret(16),
      )
      .unwrap();
    localState = publishCanonicalProjection({
      state: localState,
      commit: localCommit,
      session: local,
    }).unwrap().state;
    localClient.enqueue(batch("local", "local-0", 0, localCommit.ops)).unwrap();
    expect(local.canUndo).toBe(true);

    // The other peer suggests deleting 5..6, against the local one, and
    // types two characters at the paragraph's start; it is sequenced first.
    const remoteOps: DocumentOp[] = [
      trackedDeletion(blockId, [5, 6], revisionId(leaseOf("remote")), "remote"),
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at(blockId, 0),
        text: "xy",
      },
    ];
    const answer = room.submit(batch("remote", "remote-0", 0, remoteOps));
    expect(answer.type).toBe("ack");
    const sequenced = room.broadcasts.at(-1) ?? panic("Sequenced.");

    const changes = localClient.receiveBroadcast(sequenced).unwrap();
    expect(changes.map(({ type }) => type)).toEqual(["replaced"]);
    const change: ClientDocumentChange = changes.at(0) ?? panic("One change.");
    // Only the cuts differ: positions follow the rebased change.
    expect(change.type === "replaced" && change.moves !== undefined).toBe(true);

    const prepared = local.prepareRemote(localState, change).unwrap();
    localState = publishCanonicalProjection({
      state: localState,
      commit: prepared,
      session: local,
    }).unwrap().state;
    // The host hears that the local step is gone.
    expect(drops).toEqual([{ undo: 1, redo: 0 }]);
    expect(local.canUndo).toBe(false);
    // The caret moved past the two characters typed before it, rather than
    // staying at its old offset.
    const selection = local.projection.selectionAt(localState).unwrap();
    expect(selection.head).toEqual(at(blockId, 18));
    expect(selection.anchor).toEqual(at(blockId, 18));
    expect(local.document).toEqual(localClient.document);

    // The local edit is sequenced after the remote one, and every peer
    // ends on the authority's document.
    const submitted = localClient.nextSubmission() ?? panic("Pending.");
    const localAnswer = room.submit(submitted);
    expect(localAnswer.type).toBe("ack");
    if (localAnswer.type === "ack") localClient.receiveAck(localAnswer);
    localClient
      .receiveBroadcast(room.broadcasts.at(-1) ?? panic("Sequenced."))
      .unwrap();
    expect(localClient.pending).toHaveLength(0);
    expect(localClient.document).toEqual(room.document);
  });
});
