import type { Result } from "better-result";

import {
  type BatchSubmission,
  createSequencer,
  type DocumentBatch,
  type IdBlock,
  type IdLease,
  type IdLeaseKind,
  type IdLeaseRefusal,
  type IdLeaseState,
  type SequencedBatch,
  type SequencerOptions,
} from "@stll/docx-core/ops";

import type { Document } from "../types/document";
import type { CanonicalPresenceChannel } from "./canonicalPresence";

/**
 * A submission the transport lost to a disconnection: it may or may not have
 * reached the authority. The sender submits the same batch, under the same
 * operation id, once connected again; the authority answers a repeated id
 * from the same actor with its first outcome.
 */
export type CanonicalRoomUnanswered = { type: "disconnected"; opId: string };
export type CanonicalRoomSubmission = BatchSubmission | CanonicalRoomUnanswered;

/**
 * How one actor reaches a room's authority (a sequencer) and its journal. It
 * carries batches and leases only; the controller owns the protocol. Neither
 * `submit` nor `subscribe` calls back before returning.
 */
export type CanonicalRoomTransport = {
  /**
   * The room's base snapshot: `canonicalSnapshotId` of the DOCX the room was
   * seeded from, which revision 0 is.
   */
  readonly snapshot: string;
  /** The last revision the authority had sequenced, as last heard. */
  readonly headRev: number;
  readonly connected: boolean;
  /** Ask the authority to sequence `batch`. */
  submit(batch: DocumentBatch): Promise<CanonicalRoomSubmission>;
  /**
   * Deliver every sequenced batch after `afterRev`, then each new one. A
   * delivery may repeat or arrive out of order; the receiver orders them by
   * revision. A disconnection ends the subscription: subscribe again once
   * connected. Returns the unsubscribe.
   */
  subscribe(
    afterRev: number,
    onSequenced: (batch: SequencedBatch) => void,
  ): () => void;
  /** Called on every connection change; returns the unsubscribe. */
  onConnection(listener: (connected: boolean) => void): () => void;
  /** The actor's remaining id lease; the first call grants one block per kind. */
  grantIdLease(actor: string): Promise<Result<IdLease, IdLeaseRefusal>>;
  /** One more block of `kind` for an actor running low. */
  renewIdLease(
    actor: string,
    kind: IdLeaseKind,
  ): Promise<Result<IdBlock, IdLeaseRefusal>>;
  /** Where the room's peers share their selections, when the transport has one. */
  readonly presence?: CanonicalPresenceChannel;
};

/** A room's base snapshot identity: the SHA-256 of the DOCX bytes, in hex. */
export const canonicalSnapshotId = async (
  docx: ArrayBuffer | Uint8Array,
): Promise<string> => {
  const digest = await crypto.subtle.digest("SHA-256", new Uint8Array(docx));
  return [...new Uint8Array(digest)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
};

export type CanonicalMemoryConnection = CanonicalRoomTransport & {
  /** Lose the connection: undelivered batches and unanswered submissions are dropped. */
  disconnect(): void;
  reconnect(): void;
};

export type CanonicalMemoryRoom = {
  /** One actor's transport to the room, connected. */
  connect(): CanonicalMemoryConnection;
  readonly document: Document;
  readonly headRev: number;
  readonly journal: readonly SequencedBatch[];
  /** The lease table to persist with the journal. */
  readonly idLeases: IdLeaseState;
};

export type CanonicalMemoryRoomOptions = {
  /** Revision 0: the document a canonical session reads from the room's DOCX. */
  document: Document;
  /** `canonicalSnapshotId` of that DOCX. */
  snapshot: string;
  sequencer?: SequencerOptions;
  /**
   * Runs each hop between the room and a connection: a submission reaching
   * the authority, its answer, a sequenced batch reaching a subscriber. The
   * default runs it on a microtask; a test may delay, reorder or repeat them.
   */
  schedule?: (hop: () => void) => void;
};

/**
 * A room in memory around `createSequencer`, for tests and for hosts whose
 * actors share one process. Lease calls are answered directly.
 */
export const createCanonicalMemoryRoom = ({
  document,
  snapshot,
  sequencer: sequencerOptions,
  schedule = queueMicrotask,
}: CanonicalMemoryRoomOptions): CanonicalMemoryRoom => {
  const sequencer = createSequencer(document, sequencerOptions);
  const broadcastListeners = new Set<(batch: SequencedBatch) => void>();

  const connect = (): CanonicalMemoryConnection => {
    let connected = true;
    let epoch = 0;
    const subscriptions = new Set<(batch: SequencedBatch) => void>();
    const connectionListeners = new Set<(connected: boolean) => void>();
    const unanswered = new Set<() => void>();
    // A hop runs only if the connection it was sent on is still up.
    const hop = (run: () => void): void => {
      const sentOn = epoch;
      schedule(() => {
        if (connected && epoch === sentOn) run();
      });
    };
    const fromRoom = (batch: SequencedBatch): void => {
      for (const listener of subscriptions) hop(() => listener(batch));
    };
    broadcastListeners.add(fromRoom);
    const setConnected = (next: boolean): void => {
      if (connected === next) return;
      connected = next;
      epoch += 1;
      if (!next) {
        subscriptions.clear();
        for (const answer of unanswered) answer();
        unanswered.clear();
      }
      for (const listener of connectionListeners) listener(next);
    };
    return {
      snapshot,
      get headRev() {
        return sequencer.headRev;
      },
      get connected() {
        return connected;
      },
      submit: (batch) => {
        if (!connected)
          return Promise.resolve({ type: "disconnected", opId: batch.opId });
        return new Promise((resolve) => {
          const lost = () =>
            resolve({ type: "disconnected", opId: batch.opId });
          unanswered.add(lost);
          hop(() => {
            const before = sequencer.headRev;
            const outcome = sequencer.submit(batch);
            const sequenced = sequencer.broadcasts.at(-1);
            if (sequencer.headRev > before && sequenced !== undefined)
              for (const listener of broadcastListeners) listener(sequenced);
            hop(() => {
              unanswered.delete(lost);
              resolve(outcome);
            });
          });
        });
      },
      subscribe: (afterRev, onSequenced) => {
        if (!connected) return () => {};
        for (const batch of sequencer.broadcasts.slice(afterRev))
          hop(() => onSequenced(batch));
        subscriptions.add(onSequenced);
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
      grantIdLease: (actor) => Promise.resolve(sequencer.grantIdLease(actor)),
      renewIdLease: (actor, kind) =>
        Promise.resolve(sequencer.renewIdLease(actor, kind)),
      disconnect: () => setConnected(false),
      reconnect: () => setConnected(true),
    };
  };

  return {
    connect,
    get document() {
      return sequencer.document;
    },
    get headRev() {
      return sequencer.headRev;
    },
    get journal() {
      return sequencer.broadcasts;
    },
    get idLeases() {
      return sequencer.idLeases;
    },
  };
};
