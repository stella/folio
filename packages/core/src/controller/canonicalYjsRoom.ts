import { Result, TaggedError, type TaggedErrorClass } from "better-result";
import * as v from "valibot";
import type { Awareness } from "y-protocols/awareness";
import type * as Y from "yjs";

import {
  ID_LEASE_KINDS,
  type IdLeaseKind,
  type IdLeaseState,
  type SequencedBatch,
  validateDocumentBatch,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import type { CanonicalPresenceChannel } from "./canonicalPresence";
import {
  type CanonicalRoomRecord,
  decodeRoomRecord,
  restoreCanonicalRoomAuthority,
} from "./canonicalRoomAuthority";
import type {
  CanonicalRoomSubmission,
  CanonicalRoomTransport,
} from "./canonicalRoomTransport";
import {
  decodeLeaseGrant,
  decodeLeaseRenewal,
  decodeRoomSubmission,
  encodeLeaseGrant,
  encodeLeaseRenewal,
  encodeRoomSubmission,
  toWireJson,
} from "./canonicalRoomWire";

/*
 * A canonical room carried by a Yjs document, for hosts that already sync
 * one (a y-websocket or Hocuspocus server). The document is a transport, not
 * the shared model: clients append submissions and lease requests to an
 * inbox array; one authority, which the host server runs
 * (`serveCanonicalYjsRoom`), sequences them, appends every record to a
 * journal array and answers each request in an answers map, keyed by actor
 * and operation id. Only the authority writes the journal: a client reads
 * an entry another Yjs client wrote as foreign, refuses it and reports it.
 *
 * This module touches the host's own Yjs document and awareness and imports
 * no Yjs code: core keeps no eager Yjs dependency.
 *
 * A room needs its authority. A serverless room (y-webrtc) has none: give it
 * one, a peer that is always present or a server that joins the room.
 */

/** The Yjs client id the authority writes the journal under, unless the host picks another. */
export const CANONICAL_YJS_AUTHORITY = 0x46_4f_4c_49;

const INBOX = "folio.canonical.inbox";
const JOURNAL = "folio.canonical.journal";
const ANSWERS = "folio.canonical.answers";
const ROOM = "folio.canonical.room";
const SNAPSHOT_KEY = "snapshot";
const AWARENESS_FIELD = "folioCanonical";

const CanonicalYjsRoomErrorBase: TaggedErrorClass<"CanonicalYjsRoomError"> =
  TaggedError("CanonicalYjsRoomError");

/**
 * What went wrong in a Yjs room. `snapshotMismatch`: the room was seeded
 * from another DOCX. `restoreFailed`: the journal does not rebuild the
 * room's sequencer. `foreignJournalEntry`: a client other than the authority
 * wrote the journal; the entry is refused and the room goes on.
 * `journalRewritten`: an authority entry was deleted. `journalConflict`: two
 * authority entries claim one revision. `invalidJournalEntry`: an authority
 * entry is malformed. `invalidInboxEntry`: the authority could not read a
 * request and dropped it. `invalidAnswer`: an answer is malformed. A client
 * stops following the room on every reason but `foreignJournalEntry` and
 * `invalidAnswer`.
 */
export class CanonicalYjsRoomError extends CanonicalYjsRoomErrorBase<{
  message: string;
  reason:
    | "snapshotMismatch"
    | "restoreFailed"
    | "foreignJournalEntry"
    | "journalRewritten"
    | "journalConflict"
    | "invalidJournalEntry"
    | "invalidInboxEntry"
    | "invalidAnswer";
  cause?: unknown;
}> {}

const yjsRoomError = (
  reason: CanonicalYjsRoomError["reason"],
  message: string,
  cause?: unknown,
) =>
  new CanonicalYjsRoomError({
    reason,
    message,
    ...(cause === undefined ? {} : { cause }),
  });

const NAME = v.pipe(v.string(), v.minLength(1));
const INBOX_ENTRY = v.variant("type", [
  v.object({ type: v.literal("submit"), batch: v.unknown() }),
  v.object({
    type: v.literal("lease"),
    actor: NAME,
    request: NAME,
    kind: v.exactOptional(v.picklist(Object.values(ID_LEASE_KINDS))),
  }),
]);
type InboxEntry = v.InferOutput<typeof INBOX_ENTRY>;
const SUBMITTED = v.object({
  type: v.literal("submit"),
  batch: v.object({ actor: v.string(), opId: v.string() }),
});

const submitKey = (actor: string, opId: string) =>
  JSON.stringify(["submit", actor, opId]);
const leaseKey = (actor: string, request: string) =>
  JSON.stringify(["lease", actor, request]);

type RoomTypes = {
  inbox: Y.Array<unknown>;
  journal: Y.Array<unknown>;
  answers: Y.Map<unknown>;
  room: Y.Map<unknown>;
};
const roomTypes = (doc: Y.Doc): RoomTypes => ({
  inbox: doc.getArray(INBOX),
  journal: doc.getArray(JOURNAL),
  answers: doc.getMap(ANSWERS),
  room: doc.getMap(ROOM),
});

/** The live items of an array, in order. */
const itemsOf = (array: Y.Array<unknown>): Y.Item[] => {
  const items: Y.Item[] = [];
  for (let item = array._start; item !== null; item = item.right)
    items.push(item);
  return items;
};
/** The values an item holds, if it is a live entry. */
const valuesOf = (item: Y.Item): readonly unknown[] =>
  item.deleted || !item.countable ? [] : item.content.getContent();

export type ServeCanonicalYjsRoomOptions = {
  /** The server's copy of the room's Yjs document, with its stored state loaded. */
  doc: Y.Doc;
  /** `canonicalSnapshotId` of the DOCX the room was seeded from. */
  snapshot: string;
  /** Revision 0: the document a canonical session reads from that DOCX. */
  document: Document;
  idBlockSizes?: IdLeaseState["blockSizes"];
  /** The Yjs client id the journal is written under; the clients name the same. */
  authority?: number;
  onError?: (error: CanonicalYjsRoomError) => void;
};

export type CanonicalYjsAuthority = {
  readonly document: Document;
  readonly headRev: number;
  readonly journal: readonly SequencedBatch[];
  /** Stop answering; the document keeps what was written. */
  stop(): void;
};

/**
 * Run a room's authority on the host server's copy of its Yjs document. It
 * rebuilds the sequencer from the journal's authority entries, then takes
 * every request in the inbox, in order: sequences a submission, leases ids,
 * and in one transaction appends the records, writes the answer and removes
 * the request. The host persists the document's updates before it sends
 * them on, so a restarted authority finds every entry a client saw.
 */
export const serveCanonicalYjsRoom = ({
  doc,
  snapshot,
  document,
  idBlockSizes,
  authority = CANONICAL_YJS_AUTHORITY,
  onError = () => {},
}: ServeCanonicalYjsRoomOptions): Result<
  CanonicalYjsAuthority,
  CanonicalYjsRoomError
> => {
  const { inbox, journal, answers, room } = roomTypes(doc);
  const seeded = room.get(SNAPSHOT_KEY);
  if (seeded !== undefined && seeded !== snapshot)
    return Result.err(
      yjsRoomError(
        "snapshotMismatch",
        "The room was seeded from another document.",
      ),
    );
  const records = itemsOf(journal).flatMap((item) =>
    item.id.client === authority ? valuesOf(item) : [],
  );
  const restored = restoreCanonicalRoomAuthority({
    document,
    records,
    ...(idBlockSizes === undefined ? {} : { idBlockSizes }),
  });
  if (restored.isErr())
    return Result.err(
      yjsRoomError("restoreFailed", restored.error.message, restored.error),
    );
  const sequencer = restored.value;
  const origin = { serveCanonicalYjsRoom: true };
  // Yjs gives a document a new client id when another client's update
  // touches its own; the journal is always written under the authority's.
  const write = (change: () => void): void => {
    doc.clientID = authority;
    doc.transact(change, origin);
  };
  const append = (added: readonly CanonicalRoomRecord[]): void => {
    if (added.length > 0)
      journal.push(added.map((record) => toWireJson(record)));
  };

  const take = (value: unknown): void => {
    const parsed = v.safeParse(INBOX_ENTRY, value);
    if (!parsed.success) {
      onError(
        yjsRoomError("invalidInboxEntry", "An inbox entry is malformed."),
      );
      return;
    }
    const entry: InboxEntry = parsed.output;
    switch (entry.type) {
      case "submit": {
        const batch = validateDocumentBatch(entry.batch);
        if (batch.isErr()) {
          onError(
            yjsRoomError("invalidInboxEntry", batch.error.message, batch.error),
          );
          return;
        }
        const { answer, records: added } = sequencer.submit(batch.value);
        append(added);
        answers.set(
          submitKey(batch.value.actor, batch.value.opId),
          toWireJson(encodeRoomSubmission(answer)),
        );
        return;
      }
      case "lease": {
        const { actor, request, kind } = entry;
        if (kind === undefined) {
          const granted = sequencer.grantIdLease(actor);
          append(granted.records);
          answers.set(
            leaseKey(actor, request),
            toWireJson(encodeLeaseGrant(granted.answer)),
          );
          return;
        }
        const renewed = sequencer.renewIdLease(actor, kind);
        append(renewed.records);
        answers.set(
          leaseKey(actor, request),
          toWireJson(encodeLeaseRenewal(renewed.answer)),
        );
        return;
      }
      default: {
        const exhaustive: never = entry;
        return exhaustive;
      }
    }
  };

  let taking = false;
  let stopped = false;
  const takeInbox = (): void => {
    if (taking || stopped) return;
    taking = true;
    while (inbox.length > 0) {
      const entries = inbox.toArray();
      write(() => {
        for (const entry of entries) take(entry);
        inbox.delete(0, entries.length);
      });
    }
    taking = false;
  };
  const watchJournal = (event: Y.YArrayEvent<unknown>): void => {
    for (const item of event.changes.added) {
      if (item.id.client === authority) continue;
      for (const _ of valuesOf(item))
        onError(
          yjsRoomError(
            "foreignJournalEntry",
            `Yjs client ${item.id.client} wrote the journal.`,
          ),
        );
    }
  };

  if (seeded === undefined) write(() => room.set(SNAPSHOT_KEY, snapshot));
  inbox.observe(takeInbox);
  journal.observe(watchJournal);
  takeInbox();

  return Result.ok({
    get document() {
      return sequencer.document;
    },
    get headRev() {
      return sequencer.headRev;
    },
    get journal() {
      return sequencer.journal;
    },
    stop: () => {
      stopped = true;
      inbox.unobserve(takeInbox);
      journal.unobserve(watchJournal);
    },
  });
};

/** A provider's connection state (y-websocket's `status`, for example). */
export type CanonicalYjsConnection = {
  readonly connected: boolean;
  /** Returns the unsubscribe. */
  onChange(listener: (connected: boolean) => void): () => void;
};

export type OpenCanonicalYjsRoomOptions = {
  /** This client's copy of the room's Yjs document. */
  doc: Y.Doc;
  /** `canonicalSnapshotId` of the DOCX this client opened. */
  snapshot: string;
  /** The Yjs client id the room's authority writes under. */
  authority?: number;
  /** Without one the room counts as connected throughout. */
  connection?: CanonicalYjsConnection;
  /** The provider's awareness, to carry presence. */
  awareness?: Awareness;
  onError?: (error: CanonicalYjsRoomError) => void;
};

export type CanonicalYjsRoom = CanonicalRoomTransport & {
  /** Why the room stopped following the journal, if it did. */
  readonly error: CanonicalYjsRoomError | undefined;
  /** Stop reading and writing the document. */
  close(): void;
};

/** A client's transport to a room carried by `doc`. */
export const openCanonicalYjsRoom = ({
  doc,
  snapshot,
  authority = CANONICAL_YJS_AUTHORITY,
  connection,
  awareness,
  onError = () => {},
}: OpenCanonicalYjsRoomOptions): CanonicalYjsRoom => {
  const { inbox, journal, answers, room } = roomTypes(doc);
  const origin = { openCanonicalYjsRoom: true };
  const sequenced: SequencedBatch[] = [];
  const waiting = new Map<number, SequencedBatch>();
  const refused = new WeakSet<Y.Item>();
  const submissions = new Map<
    string,
    (outcome: CanonicalRoomSubmission) => void
  >();
  const leaseRequests = new Map<string, (answer: unknown) => boolean>();
  const subscriptions = new Set<(batch: SequencedBatch) => void>();
  const connectionListeners = new Set<(connected: boolean) => void>();
  let failure: CanonicalYjsRoomError | undefined;
  let closed = false;
  const connected = () =>
    !closed && failure === undefined && (connection?.connected ?? true);

  const loseConnection = (): void => {
    const lost = [...submissions];
    submissions.clear();
    subscriptions.clear();
    for (const [, resolve] of lost) resolve({ type: "disconnected", opId: "" });
    for (const listener of [...connectionListeners]) listener(false);
  };
  const fail = (error: CanonicalYjsRoomError): void => {
    if (failure !== undefined) return;
    const wasConnected = connected();
    failure = error;
    onError(error);
    if (wasConnected) loseConnection();
  };

  const checkSnapshot = (): void => {
    const seeded = room.get(SNAPSHOT_KEY);
    if (seeded !== undefined && seeded !== snapshot)
      fail(
        yjsRoomError(
          "snapshotMismatch",
          "The room was seeded from another document.",
        ),
      );
  };

  const readEntries = (items: Iterable<Y.Item>): void => {
    for (const item of items) {
      if (failure !== undefined) return;
      const values = valuesOf(item);
      if (values.length === 0) continue;
      if (item.id.client !== authority) {
        if (refused.has(item)) continue;
        refused.add(item);
        for (const _ of values)
          onError(
            yjsRoomError(
              "foreignJournalEntry",
              `Yjs client ${item.id.client} wrote the journal; its entry is refused.`,
            ),
          );
        continue;
      }
      for (const value of values) {
        const record = decodeRoomRecord(value);
        if (record === undefined) {
          fail(
            yjsRoomError(
              "invalidJournalEntry",
              "A journal entry is malformed.",
            ),
          );
          return;
        }
        if (record.type === "leased") continue;
        const { batch } = record;
        const known =
          sequenced.at(batch.revision - 1) ?? waiting.get(batch.revision);
        if (known !== undefined) {
          if (known.opId !== batch.opId)
            fail(
              yjsRoomError(
                "journalConflict",
                `Two journal entries claim revision ${batch.revision}.`,
              ),
            );
          continue;
        }
        waiting.set(batch.revision, batch);
      }
    }
    for (
      let next = waiting.get(sequenced.length + 1);
      next !== undefined && failure === undefined;
      next = waiting.get(sequenced.length + 1)
    ) {
      waiting.delete(next.revision);
      sequenced.push(next);
      for (const listener of [...subscriptions]) listener(next);
    }
  };

  const settle = (key: string): void => {
    const value = answers.get(key);
    if (value === undefined) return;
    const submission = submissions.get(key);
    if (submission !== undefined) {
      const outcome = decodeRoomSubmission(value);
      if (outcome === undefined) {
        onError(yjsRoomError("invalidAnswer", "An answer is malformed."));
        return;
      }
      submissions.delete(key);
      submission(outcome);
      return;
    }
    const lease = leaseRequests.get(key);
    if (lease !== undefined && lease(value)) leaseRequests.delete(key);
    else if (lease !== undefined)
      onError(yjsRoomError("invalidAnswer", "A lease answer is malformed."));
  };

  const watchJournal = (event: Y.YArrayEvent<unknown>): void => {
    if (closed || failure !== undefined) return;
    for (const item of event.changes.deleted)
      if (item.id.client === authority) {
        fail(
          yjsRoomError("journalRewritten", "An authority entry was deleted."),
        );
        return;
      }
    readEntries(event.changes.added);
  };
  const watchAnswers = (event: Y.YMapEvent<unknown>): void => {
    if (closed) return;
    for (const key of event.keysChanged)
      if (typeof key === "string") settle(key);
  };
  const watchRoom = (): void => {
    if (!closed) checkSnapshot();
  };

  const requestLease = <T>(
    actor: string,
    kind: IdLeaseKind | undefined,
    decode: (answer: unknown) => T | undefined,
  ): Promise<T> =>
    new Promise((resolve) => {
      const request = crypto.randomUUID();
      const key = leaseKey(actor, request);
      leaseRequests.set(key, (answer) => {
        const decoded = decode(answer);
        if (decoded === undefined) return false;
        resolve(decoded);
        return true;
      });
      doc.transact(() => {
        inbox.push([
          {
            type: "lease",
            actor,
            request,
            ...(kind === undefined ? {} : { kind }),
          },
        ]);
      }, origin);
    });

  checkSnapshot();
  readEntries(itemsOf(journal));
  journal.observe(watchJournal);
  answers.observe(watchAnswers);
  room.observe(watchRoom);
  const offConnection =
    connection?.onChange((next) => {
      if (closed || failure !== undefined) return;
      if (!next) {
        loseConnection();
        return;
      }
      for (const listener of [...connectionListeners]) listener(true);
    }) ?? (() => {});

  const presence: CanonicalPresenceChannel | undefined =
    awareness === undefined
      ? undefined
      : {
          publish: (state) => {
            if (!closed) awareness.setLocalStateField(AWARENESS_FIELD, state);
          },
          subscribe: (listener) => {
            const stateOf = (id: number): unknown =>
              awareness.getStates().get(id)?.[AWARENESS_FIELD] ?? null;
            const onChange = ({
              added,
              updated,
              removed,
            }: {
              added: number[];
              updated: number[];
              removed: number[];
            }): void => {
              for (const id of [...added, ...updated])
                if (id !== awareness.clientID)
                  listener(String(id), stateOf(id));
              for (const id of removed)
                if (id !== awareness.clientID) listener(String(id), null);
            };
            awareness.on("change", onChange);
            let subscribed = true;
            queueMicrotask(() => {
              if (!subscribed) return;
              for (const id of awareness.getStates().keys())
                if (id !== awareness.clientID)
                  listener(String(id), stateOf(id));
            });
            return () => {
              subscribed = false;
              awareness.off("change", onChange);
            };
          },
        };

  return {
    snapshot,
    get headRev() {
      return sequenced.length;
    },
    get connected() {
      return connected();
    },
    get error() {
      return failure;
    },
    ...(presence === undefined ? {} : { presence }),
    submit: (batch) => {
      if (!connected())
        return Promise.resolve({ type: "disconnected", opId: batch.opId });
      // What the authority would answer first, without the trip.
      const valid = validateDocumentBatch(batch);
      if (valid.isErr())
        return Promise.resolve({
          type: "reject",
          opId: batch.opId,
          reason: valid.error,
          headRev: sequenced.length,
        });
      const key = submitKey(batch.actor, batch.opId);
      return new Promise((resolve) => {
        submissions.get(key)?.({ type: "disconnected", opId: batch.opId });
        submissions.set(key, (outcome) =>
          resolve(
            outcome.type === "disconnected"
              ? { type: "disconnected", opId: batch.opId }
              : outcome,
          ),
        );
        if (answers.has(key)) {
          queueMicrotask(() => settle(key));
          return;
        }
        // Submitted again after a disconnection: the request already waits.
        const queued = inbox.toArray().some((entry) => {
          const parsed = v.safeParse(SUBMITTED, entry);
          return (
            parsed.success &&
            parsed.output.batch.actor === batch.actor &&
            parsed.output.batch.opId === batch.opId
          );
        });
        if (!queued)
          doc.transact(() => {
            inbox.push([toWireJson({ type: "submit", batch: valid.value })]);
          }, origin);
      });
    },
    subscribe: (afterRev, onSequenced) => {
      if (!connected()) return () => {};
      subscriptions.add(onSequenced);
      const replay = sequenced.slice(afterRev);
      queueMicrotask(() => {
        if (!subscriptions.has(onSequenced)) return;
        for (const batch of replay) onSequenced(batch);
      });
      return () => {
        subscriptions.delete(onSequenced);
      };
    },
    onConnection: (listener) => {
      connectionListeners.add(listener);
      return () => {
        connectionListeners.delete(listener);
      };
    },
    grantIdLease: (actor) => requestLease(actor, undefined, decodeLeaseGrant),
    renewIdLease: (actor, kind) =>
      requestLease(actor, kind, decodeLeaseRenewal),
    close: () => {
      if (closed) return;
      const wasConnected = connected();
      if (awareness !== undefined)
        awareness.setLocalStateField(AWARENESS_FIELD, null);
      closed = true;
      journal.unobserve(watchJournal);
      answers.unobserve(watchAnswers);
      room.unobserve(watchRoom);
      offConnection();
      if (wasConnected) loseConnection();
    },
  };
};
