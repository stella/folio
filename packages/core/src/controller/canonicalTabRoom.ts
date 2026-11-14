import { Result, TaggedError, type TaggedErrorClass } from "better-result";
import * as v from "valibot";

import {
  ID_LEASE_KINDS,
  type IdLeaseKind,
  type IdLeaseState,
  type SequencedBatch,
  validateDocumentBatch,
  validateSequencedBatch,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import {
  type CanonicalPresenceChannel,
  type CanonicalPresenceState,
  parseCanonicalPresence,
} from "./canonicalPresence";
import {
  type CanonicalRoomAuthority,
  type CanonicalRoomRecord,
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
  UNAVAILABLE_LEASE_ANSWER,
} from "./canonicalRoomWire";

/*
 * A room shared by the tabs of one origin. The authority runs in exactly one
 * tab, the one holding the room's Web Lock; every tab, that one included,
 * reaches it over a BroadcastChannel. When the leading tab goes away the lock
 * passes to a waiting tab, which rebuilds the sequencer from the durable
 * journal (every record, in order) before it announces itself; clients then
 * count the old leader's unanswered submissions as lost to a disconnection,
 * and the controller submits them again under the same operation id.
 */

/**
 * The durable store a room's leaders share, in this origin: what one leader
 * stored, the next one reads. A leader stores each record before it
 * broadcasts or answers anything that depends on it.
 */
export type CanonicalTabJournal = {
  /** The records stored so far, in order, and the last leadership term. */
  load(): Promise<{ records: readonly unknown[]; term: number }>;
  /** Store the term of a leader taking over; resolves once it is durable. */
  storeTerm(term: number): Promise<void>;
  /** Store records after the stored ones; resolves once they are durable. */
  append(records: readonly CanonicalRoomRecord[]): Promise<void>;
};

/** A message channel to every other tab of the room (a BroadcastChannel). */
export type CanonicalTabChannel = {
  post(message: string): void;
  /** Returns the unsubscribe. */
  onMessage(listener: (message: unknown) => void): () => void;
  close(): void;
};

/** What a tab room needs from the browser. */
export type CanonicalTabEnvironment = {
  openChannel(name: string): CanonicalTabChannel;
  /**
   * Run `lead` once this tab holds the exclusive lock `name`, and hold it
   * until `lead` settles or the tab goes away (`navigator.locks.request`).
   * Aborting `signal` gives up waiting. Settles once the lock is released or
   * the request is given up.
   */
  requestLock(
    name: string,
    signal: AbortSignal,
    lead: () => Promise<void>,
  ): Promise<void>;
};

/** BroadcastChannel and the Web Locks API, as a browser tab has them. */
export const browserTabEnvironment = (): CanonicalTabEnvironment => ({
  openChannel: (name) => {
    const channel = new BroadcastChannel(name);
    return {
      // oxlint-disable-next-line unicorn/require-post-message-target-origin -- a BroadcastChannel has no target origin
      post: (message) => channel.postMessage(message),
      onMessage: (listener) => {
        const onMessage = (event: MessageEvent) => listener(event.data);
        channel.addEventListener("message", onMessage);
        return () => channel.removeEventListener("message", onMessage);
      },
      close: () => channel.close(),
    };
  },
  requestLock: async (name, signal, lead) => {
    await navigator.locks.request(name, { signal }, lead);
  },
});

const CanonicalTabRoomErrorBase: TaggedErrorClass<"CanonicalTabRoomError"> =
  TaggedError("CanonicalTabRoomError");

/**
 * What went wrong in a tab room. `journalUnavailable`: the durable journal
 * could not be read or written; a leader that cannot store steps down
 * without answering. `restoreFailed`: the stored records do not rebuild the
 * room's sequencer. `lockUnavailable`: the tab could not ask for the room's
 * lock. `invalidMessage`: a message on the room's channel was malformed and
 * was ignored.
 */
export class CanonicalTabRoomError extends CanonicalTabRoomErrorBase<{
  message: string;
  reason:
    | "journalUnavailable"
    | "restoreFailed"
    | "lockUnavailable"
    | "invalidMessage";
  cause?: unknown;
}> {}

export type CanonicalTabRoomOptions = {
  /** The room's name in this origin, for example the document's id. */
  room: string;
  /** `canonicalSnapshotId` of the DOCX every tab of the room opened. */
  snapshot: string;
  /** Revision 0: the document a canonical session reads from that DOCX. */
  document: Document;
  /** The durable journal every leader of the room reads and appends to. */
  journal: CanonicalTabJournal;
  /** How many ids one lease block holds, per kind; the same in every tab. */
  idBlockSizes?: IdLeaseState["blockSizes"];
  /** The browser by default. */
  environment?: CanonicalTabEnvironment;
  /**
   * How often a tab publishes its presence again, and a third of how long
   * a silent tab's presence is kept (a closed tab says goodbye, a crashed one
   * cannot). 0 publishes only on change and keeps presence until goodbye.
   */
  presenceRefreshMs?: number;
  onError?: (error: CanonicalTabRoomError) => void;
};

export type CanonicalTabRoom = CanonicalRoomTransport & {
  readonly presence: CanonicalPresenceChannel;
  /** Whether this tab runs the room's authority now. */
  readonly leading: boolean;
  /** Leave the room: lead no more, answer nothing more, say goodbye. */
  close(): void;
};

const TERM = v.pipe(v.number(), v.safeInteger(), v.minValue(1));
const REVISION = v.pipe(v.number(), v.safeInteger(), v.minValue(0));
const NAME = v.pipe(v.string(), v.minLength(1));
const MESSAGE = v.variant("type", [
  v.object({ type: v.literal("discover"), from: NAME }),
  v.object({ type: v.literal("leader"), term: TERM, headRev: REVISION }),
  v.object({
    type: v.literal("submit"),
    term: TERM,
    from: NAME,
    batch: v.unknown(),
  }),
  v.object({
    type: v.literal("answer"),
    term: TERM,
    to: NAME,
    outcome: v.unknown(),
  }),
  v.object({
    type: v.literal("catchUp"),
    term: TERM,
    from: NAME,
    afterRev: REVISION,
  }),
  v.object({
    type: v.literal("sequenced"),
    term: TERM,
    to: v.exactOptional(NAME),
    batches: v.array(v.unknown()),
  }),
  v.object({
    type: v.literal("lease"),
    term: TERM,
    from: NAME,
    request: NAME,
    actor: NAME,
    kind: v.exactOptional(v.picklist(Object.values(ID_LEASE_KINDS))),
  }),
  v.object({
    type: v.literal("leaseAnswer"),
    term: TERM,
    to: NAME,
    request: NAME,
    answer: v.unknown(),
  }),
  v.object({ type: v.literal("presence"), from: NAME, state: v.unknown() }),
  v.object({ type: v.literal("presenceQuery"), from: NAME }),
]);
type Message = v.InferOutput<typeof MESSAGE>;

/** A lease call waiting for the leader, sent again to each new one. */
type LeaseRequest = {
  message: (term: number) => Message;
  /** Settle with the leader's answer; false when it is malformed. */
  settle: (answer: unknown) => boolean;
};

const PRESENCE_TIMEOUTS = 3;
const DEFAULT_PRESENCE_REFRESH_MS = 10_000;

/**
 * Join this tab to the room the same-origin tabs share. The tab runs the
 * room's authority while it holds the room's lock and is one of the room's
 * clients throughout.
 */
export const openCanonicalTabRoom = ({
  room,
  snapshot,
  document,
  journal,
  idBlockSizes,
  environment = browserTabEnvironment(),
  presenceRefreshMs = DEFAULT_PRESENCE_REFRESH_MS,
  onError = () => {},
}: CanonicalTabRoomOptions): CanonicalTabRoom => {
  const name = `folio-canonical-room:${room}:${snapshot}`;
  const tab = crypto.randomUUID();
  const channel = environment.openChannel(name);
  const abort = new AbortController();
  let closed = false;

  const report = (
    reason: CanonicalTabRoomError["reason"],
    message: string,
    cause?: unknown,
  ): void =>
    onError(
      new CanonicalTabRoomError({
        reason,
        message,
        ...(cause === undefined ? {} : { cause }),
      }),
    );

  // Every message also reaches this tab's own handlers: a channel never
  // delivers what it posted to itself.
  const send = (message: Message): void => {
    if (closed) return;
    const data = JSON.stringify(message);
    channel.post(data);
    queueMicrotask(() => receive(data));
  };

  // The client side, in every tab.
  let term = 0;
  let headRev = 0;
  const answers = new Map<string, (outcome: CanonicalRoomSubmission) => void>();
  const subscriptions = new Set<(batch: SequencedBatch) => void>();
  const connectionListeners = new Set<(connected: boolean) => void>();
  const leaseRequests = new Map<string, LeaseRequest>();
  // Whether a leader of `term` is known to be leading.
  let following = false;
  const connected = () => !closed && following;

  const notifyConnection = (next: boolean): void => {
    for (const listener of [...connectionListeners]) listener(next);
  };
  const loseConnection = (): void => {
    const lost = [...answers];
    answers.clear();
    subscriptions.clear();
    for (const [opId, resolve] of lost) resolve({ type: "disconnected", opId });
    notifyConnection(false);
  };

  const followLeader = (next: number, announcedHead: number): void => {
    if (next <= term) return;
    if (following) loseConnection();
    term = next;
    following = true;
    headRev = Math.max(headRev, announcedHead);
    notifyConnection(true);
    for (const request of leaseRequests.values()) send(request.message(term));
  };

  const deliver = (batches: readonly unknown[]): void => {
    for (const value of batches) {
      const batch = validateSequencedBatch(value);
      if (batch.isErr()) {
        report("invalidMessage", batch.error.message, batch.error);
        continue;
      }
      headRev = Math.max(headRev, batch.value.revision);
      for (const listener of [...subscriptions]) listener(batch.value);
    }
  };

  const settleAnswer = (value: unknown): void => {
    const outcome = decodeRoomSubmission(value);
    if (outcome === undefined) {
      report("invalidMessage", "A submission's answer is malformed.");
      return;
    }
    if (outcome.type === "ack") headRev = Math.max(headRev, outcome.rev);
    const resolve = answers.get(outcome.opId);
    if (resolve === undefined) return;
    answers.delete(outcome.opId);
    resolve(outcome);
  };

  const requestLease = <T>(
    actor: string,
    kind: IdLeaseKind | undefined,
    decode: (answer: unknown) => T | undefined,
  ): Promise<T> =>
    new Promise((resolve) => {
      const request = crypto.randomUUID();
      const pending: LeaseRequest = {
        message: (current) => ({
          type: "lease",
          term: current,
          from: tab,
          request,
          actor,
          ...(kind === undefined ? {} : { kind }),
        }),
        settle: (answer) => {
          const decoded = decode(answer);
          if (decoded === undefined) return false;
          leaseRequests.delete(request);
          resolve(decoded);
          return true;
        },
      };
      if (closed) {
        pending.settle(UNAVAILABLE_LEASE_ANSWER);
        return;
      }
      leaseRequests.set(request, pending);
      if (connected()) send(pending.message(term));
    });

  // The authority side, in the tab holding the lock.
  let leading: number | undefined;
  let authority: CanonicalRoomAuthority | undefined;
  let release = (): void => {};
  let work = Promise.resolve();
  /** Run leader work one request at a time, in arrival order. */
  const serially = (task: () => Promise<void>): void => {
    work = work.then(task);
  };

  const stepDown = (): void => {
    leading = undefined;
    authority = undefined;
    release();
  };

  /** Store `records` durably; a leader that cannot steps down unanswered. */
  const store = async (
    records: readonly CanonicalRoomRecord[],
  ): Promise<boolean> => {
    if (records.length === 0) return true;
    const stored = await Result.tryPromise({
      try: () => journal.append(records),
      catch: (error: unknown) => error,
    });
    if (stored.isOk()) return true;
    report(
      "journalUnavailable",
      "The room's journal could not be written.",
      stored.error,
    );
    const stepped = leading;
    stepDown();
    // This tab no longer follows the term it led: its unanswered calls are
    // lost, and it stands again, so a later leader (this tab or another)
    // rebuilds from the journal and announces a higher term.
    if (following && term === stepped) {
      following = false;
      loseConnection();
    }
    if (!closed) void contend();
    return false;
  };

  /** The authority, if this tab leads in `current`. */
  const leaderIn = (current: number) =>
    leading === current ? authority : undefined;

  const lead = async (): Promise<void> => {
    if (closed) return;
    const loaded = await Result.tryPromise({
      try: () => journal.load(),
      catch: (error: unknown) => error,
    });
    if (loaded.isErr()) {
      report(
        "journalUnavailable",
        "The room's journal could not be read.",
        loaded.error,
      );
      return;
    }
    // Every record replays before the first submission is taken.
    const restored = restoreCanonicalRoomAuthority({
      document,
      records: loaded.value.records,
      ...(idBlockSizes === undefined ? {} : { idBlockSizes }),
    });
    if (restored.isErr()) {
      report("restoreFailed", restored.error.message, restored.error);
      return;
    }
    const next = loaded.value.term + 1;
    const stored = await Result.tryPromise({
      try: () => journal.storeTerm(next),
      catch: (error: unknown) => error,
    });
    if (stored.isErr()) {
      report(
        "journalUnavailable",
        "The room's journal could not be written.",
        stored.error,
      );
      return;
    }
    if (closed) return;
    const released = new Promise<void>((resolve) => {
      release = resolve;
    });
    authority = restored.value;
    leading = next;
    send({ type: "leader", term: next, headRev: restored.value.headRev });
    await released;
  };

  const onLeaderMessage = (message: Message): void => {
    switch (message.type) {
      case "discover": {
        const current = leading;
        if (current !== undefined && authority !== undefined)
          send({ type: "leader", term: current, headRev: authority.headRev });
        return;
      }
      case "submit":
        serially(async () => {
          const sequencer = leaderIn(message.term);
          if (sequencer === undefined) return;
          const batch = validateDocumentBatch(message.batch);
          if (batch.isErr()) {
            report("invalidMessage", batch.error.message, batch.error);
            return;
          }
          const { answer, records } = sequencer.submit(batch.value);
          if (!(await store(records))) return;
          const batches = records.flatMap((record) =>
            record.type === "sequenced" ? [record.batch] : [],
          );
          if (batches.length > 0)
            send({ type: "sequenced", term: message.term, batches });
          send({
            type: "answer",
            term: message.term,
            to: message.from,
            outcome: encodeRoomSubmission(answer),
          });
        });
        return;
      case "catchUp":
        serially(() => {
          const sequencer = leaderIn(message.term);
          if (sequencer !== undefined)
            send({
              type: "sequenced",
              term: message.term,
              to: message.from,
              batches: sequencer.journal.slice(message.afterRev),
            });
          return Promise.resolve();
        });
        return;
      case "lease":
        serially(async () => {
          const sequencer = leaderIn(message.term);
          if (sequencer === undefined) return;
          const { kind } = message;
          const answer =
            kind === undefined
              ? (() => {
                  const granted = sequencer.grantIdLease(message.actor);
                  return {
                    records: granted.records,
                    wire: encodeLeaseGrant(granted.answer),
                  };
                })()
              : (() => {
                  const renewed = sequencer.renewIdLease(message.actor, kind);
                  return {
                    records: renewed.records,
                    wire: encodeLeaseRenewal(renewed.answer),
                  };
                })();
          if (!(await store(answer.records))) return;
          send({
            type: "leaseAnswer",
            term: message.term,
            to: message.from,
            request: message.request,
            answer: answer.wire,
          });
        });
        return;
      default:
        return;
    }
  };

  // Presence, beside the room.
  const presenceStates = new Map<string, { state: unknown; heard: number }>();
  const presenceListeners = new Set<(peer: string, state: unknown) => void>();
  let ownPresence: CanonicalPresenceState | null = null;
  const notifyPresence = (peer: string, state: unknown): void => {
    for (const listener of [...presenceListeners]) listener(peer, state);
  };
  /**
   * Whether a state arriving from a tab is older than the one held from it:
   * the same sender and instance, at a clock no later. Deliveries reorder,
   * and the held state is what a listener subscribing later is handed, so
   * it must be the sender's latest, as a viewer keeps it.
   */
  const olderThanHeld = (held: unknown, arriving: unknown): boolean => {
    const before = parseCanonicalPresence(held);
    const after = parseCanonicalPresence(arriving);
    return (
      before !== undefined &&
      after !== undefined &&
      before.actor === after.actor &&
      before.instance === after.instance &&
      after.clock <= before.clock
    );
  };
  const postPresence = (): void => {
    if (ownPresence !== null)
      send({ type: "presence", from: tab, state: ownPresence });
  };

  const onClientMessage = (message: Message): void => {
    switch (message.type) {
      case "leader":
        followLeader(message.term, message.headRev);
        return;
      case "sequenced":
        if (message.term !== term) return;
        if (message.to !== undefined && message.to !== tab) return;
        deliver(message.batches);
        return;
      case "answer":
        if (message.term === term && message.to === tab)
          settleAnswer(message.outcome);
        return;
      case "leaseAnswer": {
        if (message.term !== term || message.to !== tab) return;
        const request = leaseRequests.get(message.request);
        if (request !== undefined && !request.settle(message.answer))
          report("invalidMessage", "A lease answer is malformed.");
        return;
      }
      case "presence": {
        if (message.from === tab) return;
        const held = presenceStates.get(message.from);
        if (message.state === null) presenceStates.delete(message.from);
        else if (
          held !== undefined &&
          olderThanHeld(held.state, message.state)
        ) {
          // The tab is alive; what it said last still stands.
          held.heard = Date.now();
          return;
        } else
          presenceStates.set(message.from, {
            state: message.state,
            heard: Date.now(),
          });
        notifyPresence(message.from, message.state);
        return;
      }
      case "presenceQuery":
        if (message.from !== tab) postPresence();
        return;
      default:
        return;
    }
  };

  const receive = (data: unknown): void => {
    if (closed) return;
    if (typeof data !== "string") return;
    const json = Result.try((): unknown => JSON.parse(data));
    const parsed = v.safeParse(MESSAGE, json.isOk() ? json.value : undefined);
    if (!parsed.success) {
      report("invalidMessage", "A message on the room's channel is malformed.");
      return;
    }
    onLeaderMessage(parsed.output);
    onClientMessage(parsed.output);
  };

  const offMessage = channel.onMessage(receive);

  const refresh =
    presenceRefreshMs > 0
      ? setInterval(() => {
          postPresence();
          const silentSince =
            Date.now() - presenceRefreshMs * PRESENCE_TIMEOUTS;
          for (const [peer, { heard }] of presenceStates)
            if (heard < silentSince) {
              presenceStates.delete(peer);
              notifyPresence(peer, null);
            }
        }, presenceRefreshMs)
      : undefined;

  send({ type: "discover", from: tab });
  send({ type: "presenceQuery", from: tab });
  const contend = async (): Promise<void> => {
    const held = await Result.tryPromise({
      try: () => environment.requestLock(name, abort.signal, lead),
      catch: (error: unknown) => error,
    });
    // Giving up the request on close rejects it: that is no error.
    if (held.isErr() && !abort.signal.aborted)
      report(
        "lockUnavailable",
        "The room's lock could not be requested.",
        held.error,
      );
  };
  void contend();

  return {
    snapshot,
    get headRev() {
      return headRev;
    },
    get connected() {
      return connected();
    },
    get leading() {
      return leading !== undefined;
    },
    submit: (batch) => {
      if (!connected())
        return Promise.resolve({ type: "disconnected", opId: batch.opId });
      // What the leader would answer first, without the trip.
      const valid = validateDocumentBatch(batch);
      if (valid.isErr())
        return Promise.resolve({
          type: "reject",
          opId: batch.opId,
          reason: valid.error,
          headRev,
        });
      return new Promise((resolve) => {
        answers.get(batch.opId)?.({ type: "disconnected", opId: batch.opId });
        answers.set(batch.opId, resolve);
        send({ type: "submit", term, from: tab, batch: valid.value });
      });
    },
    subscribe: (afterRev, onSequenced) => {
      if (!connected()) return () => {};
      subscriptions.add(onSequenced);
      send({ type: "catchUp", term, from: tab, afterRev });
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
    presence: {
      publish: (state) => {
        ownPresence = state;
        send({ type: "presence", from: tab, state });
      },
      subscribe: (listener) => {
        presenceListeners.add(listener);
        queueMicrotask(() => {
          if (!presenceListeners.has(listener)) return;
          for (const [peer, { state }] of presenceStates) listener(peer, state);
        });
        return () => {
          presenceListeners.delete(listener);
        };
      },
    },
    close: () => {
      if (closed) return;
      if (ownPresence !== null)
        send({ type: "presence", from: tab, state: null });
      const wasConnected = connected();
      closed = true;
      abort.abort();
      stepDown();
      if (refresh !== undefined) clearInterval(refresh);
      if (wasConnected) loseConnection();
      const waiting = [...leaseRequests.values()];
      leaseRequests.clear();
      for (const request of waiting) request.settle(UNAVAILABLE_LEASE_ANSWER);
      offMessage();
      channel.close();
    },
  };
};
