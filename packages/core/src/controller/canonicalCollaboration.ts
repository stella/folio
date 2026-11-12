import { Result, TaggedError, type TaggedErrorClass } from "better-result";
import type { EditorState } from "prosemirror-state";

import {
  type BatchRejection,
  type ClientDocumentChange,
  createClient,
  DOCUMENT_OP_SCHEMA_VERSION,
  type DocumentBatch,
  type DocumentOp,
  ID_LEASE_KINDS,
  ID_LEASE_REFUSAL_REASONS,
  type IdLease,
  type IdLeaseKind,
  type IdLeaseRefusal,
  idLeaseRemaining,
  type SequencedBatch,
} from "@stll/docx-core/ops";

import { readModel } from "../types/documentModel";
import {
  type CanonicalPresenceUser,
  type CanonicalRemotePresence,
  createPresenceTracker,
} from "./canonicalPresence";
import type {
  CanonicalRoomSubmission,
  CanonicalRoomTransport,
} from "./canonicalRoomTransport";
import type {
  CanonicalCommit,
  CanonicalCommittedBatch,
  CanonicalSession,
  CanonicalSessionError,
} from "./canonicalSession";

const CanonicalCollaborationErrorBase: TaggedErrorClass<"CanonicalCollaborationError"> =
  TaggedError("CanonicalCollaborationError");

/**
 * Why a session could not join a room or stopped following it.
 * `snapshotMismatch`: the room was seeded from another DOCX. `edited`: the
 * session was edited before it joined. `leaseRefused`: the authority granted
 * no id lease. `sessionRefused`: the host could not open the session.
 * `diverged`: the session refused an edit the room sequenced, or a local
 * commit the room could not take; it no longer follows the room.
 */
export class CanonicalCollaborationError extends CanonicalCollaborationErrorBase<{
  message: string;
  reason:
    | "snapshotMismatch"
    | "edited"
    | "leaseRefused"
    | "sessionRefused"
    | "diverged";
  cause?: unknown;
}> {}

/** Where the session's projection is shown. */
export type CanonicalCollaborationEditor = {
  /** The editor state the session's projection is shown in now. */
  state: () => EditorState;
  /**
   * Publish a staged commit (`publishCanonicalProjection`) and show the state
   * it produced, as for a local commit.
   */
  publish: (commit: CanonicalCommit) => Result<void, CanonicalSessionError>;
};

/**
 * What the host is told. `localEditRefused`: a local edit the room could not
 * merge, taken back from the document. `synced`: the session caught up with
 * the revision the room had when it joined. `leaseRefused`: no more ids of a
 * kind; edits minting them are refused until the room is compacted.
 * `stopped`: the session no longer follows the room.
 */
export type CanonicalCollaborationEvent =
  | {
      type: "localEditRefused";
      opId: string;
      reason: BatchRejection;
      ops: readonly DocumentOp[];
    }
  | { type: "synced"; rev: number }
  | { type: "leaseRefused"; refusal: IdLeaseRefusal }
  | { type: "stopped"; error: CanonicalCollaborationError };

export type CanonicalCollaborationPeer = {
  session: CanonicalSession;
  editor: CanonicalCollaborationEditor;
};

export type CanonicalCollaborationOptions = {
  transport: CanonicalRoomTransport;
  /** This actor's identity in the room, stable across its reloads. */
  actor: string;
  /** `canonicalSnapshotId` of the DOCX this actor opened. */
  snapshot: string;
  /**
   * Create the session over the opened DOCX with the granted lease
   * (`createCanonicalSession(document, styles, { idLease })`) and the editor
   * that shows it. It must not be edited before it joins.
   */
  open: (
    idLease: IdLease,
  ) => Result<CanonicalCollaborationPeer, CanonicalSessionError>;
  onEvent?: (event: CanonicalCollaborationEvent) => void;
  /** Ask for another block of a kind once fewer ids than this remain. */
  renewBelow?: number;
  /**
   * Share this peer's selection with the room's other peers and follow
   * theirs, over the transport's presence channel.
   */
  presence?: { user: CanonicalPresenceUser };
};

/** Another peer's presence, as this session shows it. */
export type CanonicalRemotePeer = CanonicalRemotePresence & {
  /** Editor positions in the session's projection, when both ends show there. */
  shown: { anchor: number; head: number } | undefined;
};

/** The room's other peers' selections, and this peer's own. */
export type CanonicalCollaborationPresence = {
  /**
   * Publish this peer's selection as the editor shows it now. The host
   * calls it when the selection moves; the controller publishes after each
   * remote edit and each revision the room confirms.
   */
  update(): void;
  /** Every other peer that published, moved into this session's document. */
  peers(): readonly CanonicalRemotePeer[];
  /** Called when what `peers` returns may have changed; returns the unsubscribe. */
  onChange(listener: () => void): () => void;
};

export type CanonicalCollaboration = {
  readonly session: CanonicalSession;
  /** The last revision the session applied. */
  readonly headRev: number;
  /** Local batches the room has not sequenced yet. */
  readonly pending: readonly DocumentBatch[];
  readonly stopped: boolean;
  /** Undefined unless the host asked for presence and the transport carries it. */
  readonly presence: CanonicalCollaborationPresence | undefined;
  /**
   * Apply what waited for a composition. The controller also catches up on
   * its next delivery.
   */
  flush(): void;
  /** Stop following the room; pending batches stay unsent. */
  leave(): void;
};

const DEFAULT_RENEW_BELOW = 64;

/**
 * A deliberate defect, for tests that must show their property detects it.
 * @internal
 */
export type CanonicalCollaborationFault =
  "unrebasedRemote" | "freshOpIdOnResubmit";

type Inbound =
  | { type: "sequenced"; batch: SequencedBatch }
  | { type: "answer"; outcome: CanonicalRoomSubmission };

const collaborationError = (
  reason: CanonicalCollaborationError["reason"],
  message: string,
  cause?: unknown,
): CanonicalCollaborationError =>
  new CanonicalCollaborationError({
    reason,
    message,
    ...(cause === undefined ? {} : { cause }),
  });
const failure = (
  reason: CanonicalCollaborationError["reason"],
  message: string,
  cause?: unknown,
) => Result.err(collaborationError(reason, message, cause));

/**
 * Join one canonical session to a sequenced room.
 *
 * The room starts from its base DOCX at revision 0; the session opens that
 * same DOCX (checked by snapshot identity, refused on a mismatch) and replays
 * the journal from revision 0. Each local commit is enqueued on the
 * optimistic client (a fresh operation id, the client's head revision) and
 * submitted, one at a time. A sequenced batch goes through the client first,
 * which rebases the pending local batches over it, then into the session
 * through `prepareRemote` with the change the client reported. A refused
 * local batch is taken back the same way and reported as `localEditRefused`.
 * After a disconnection the unanswered batch is submitted again under the
 * same operation id, and the journal is read again after the last applied
 * revision.
 */
export const joinCanonicalRoom = (
  options: CanonicalCollaborationOptions,
): Promise<Result<CanonicalCollaboration, CanonicalCollaborationError>> =>
  connectCanonicalRoom(options);

/** @internal `joinCanonicalRoom` with a deliberate defect, for tests. */
export const connectCanonicalRoom = async (
  {
    transport,
    actor,
    snapshot,
    open,
    onEvent = () => {},
    renewBelow = DEFAULT_RENEW_BELOW,
    presence: presenceOptions,
  }: CanonicalCollaborationOptions,
  fault?: CanonicalCollaborationFault,
): Promise<Result<CanonicalCollaboration, CanonicalCollaborationError>> => {
  if (transport.snapshot !== snapshot)
    return failure(
      "snapshotMismatch",
      "The room was seeded from another document.",
    );
  const granted = await transport.grantIdLease(actor);
  if (granted.isErr())
    return failure("leaseRefused", granted.error.message, granted.error);
  const opened = open(granted.value);
  if (opened.isErr())
    return failure("sessionRefused", opened.error.message, opened.error);
  const { session, editor } = opened.value;
  if (session.version !== 0)
    return failure("edited", "The session was edited before it joined.");

  // The client only reads the document: every batch it applies makes a new one.
  const client = createClient(readModel(session.document));
  const joinedAt = transport.headRev;
  const inbox: Inbound[] = [];
  const renewing = new Set<IdLeaseKind>();
  const unrebased = new Map<number, SequencedBatch>();
  // The batch submitted last, until the room answers it or sequences it.
  let awaiting: DocumentBatch | undefined;
  let resubmissions = 0;
  let reported = 0;
  let synced = false;
  let draining = false;
  let stopped = false;
  let unsubscribe = (): void => {};
  const presence =
    presenceOptions === undefined || transport.presence === undefined
      ? undefined
      : followPresence({
          channel: transport.presence,
          actor,
          user: presenceOptions.user,
          session,
          editor,
          headRev: () => client.headRev,
        });

  const stop = (error: CanonicalCollaborationError): void => {
    if (stopped) return;
    leave();
    onEvent({ type: "stopped", error });
  };
  const diverge = (message: string, cause?: unknown): void => {
    stop(collaborationError("diverged", message, cause));
  };

  const reportRefusals = (): void => {
    const notices = client.notices;
    for (const notice of notices.slice(reported))
      onEvent({
        type: "localEditRefused",
        opId: notice.opId,
        reason: notice.reason,
        ops: notice.ops,
      });
    reported = notices.length;
  };

  const show = (change: ClientDocumentChange): boolean => {
    const prepared = session.prepareRemote(editor.state(), change);
    if (prepared.isErr()) {
      if (prepared.error.reason === "noChange") return true;
      diverge(prepared.error.message, prepared.error);
      return false;
    }
    const published = editor.publish(prepared.value);
    if (published.isErr()) {
      diverge(published.error.message, published.error);
      return false;
    }
    return true;
  };

  const receive = (item: Inbound): boolean => {
    if (item.type === "sequenced") {
      const { batch } = item;
      // A repeated delivery of a revision already applied.
      if (batch.revision <= client.headRev) return true;
      if (awaiting?.opId === batch.opId) awaiting = undefined;
      if (fault === "unrebasedRemote") unrebased.set(batch.revision, batch);
      const before = client.headRev;
      presence?.received(batch);
      const received = client.receiveBroadcast(batch);
      if (received.isErr()) {
        diverge(received.error.message, received.error);
        return false;
      }
      const shown = received.value.every((change) => {
        const sequenced = unrebased.get(change.batch.revision);
        // The defect: the sequenced ops as sequenced, not as rebased over
        // the pending local batches.
        return show(
          change.type === "rebased" && sequenced !== undefined
            ? { type: "rebased", batch: sequenced }
            : change,
        );
      });
      // Peers' selections move once the editor shows the batches, so a host
      // reading them on change resolves them in the projection it shows.
      if (shown) presence?.applied(before, client.headRev);
      return shown;
    }
    const { outcome } = item;
    if (outcome.type === "disconnected") return true;
    if (awaiting?.opId === outcome.opId) awaiting = undefined;
    if (outcome.type === "ack") {
      client.receiveAck(outcome);
      return true;
    }
    return client.receiveReject(outcome).every(show);
  };

  const answer = async (batch: DocumentBatch): Promise<void> => {
    const outcome = await transport.submit(batch);
    inbox.push({ type: "answer", outcome });
    drain();
  };
  const send = (batch: DocumentBatch): void => {
    void answer(batch);
  };

  const pump = (): void => {
    if (stopped || !transport.connected || awaiting !== undefined) return;
    const next = client.nextSubmission();
    if (next === undefined) return;
    awaiting = next;
    send(next);
  };

  const drain = (): void => {
    if (draining || stopped) return;
    draining = true;
    try {
      // A remote edit waits while a composition is pending.
      while (inbox.length > 0 && !session.isComposing) {
        const item = inbox.shift();
        if (item === undefined || !receive(item)) break;
        reportRefusals();
      }
    } finally {
      draining = false;
    }
    if (stopped) return;
    if (!synced && client.headRev >= joinedAt) {
      synced = true;
      onEvent({ type: "synced", rev: client.headRev });
    }
    pump();
    presence?.publish();
  };

  const renewKind = async (kind: IdLeaseKind): Promise<void> => {
    renewing.add(kind);
    const renewed = await transport.renewIdLease(actor, kind);
    renewing.delete(kind);
    if (stopped) return;
    if (renewed.isOk()) {
      session.extendIdLease({ blocks: [renewed.value] });
      return;
    }
    // Not due: the authority still counts unsequenced ids as unused.
    if (renewed.error.reason !== ID_LEASE_REFUSAL_REASONS.NOT_DUE)
      onEvent({ type: "leaseRefused", refusal: renewed.error });
  };
  const renew = (): void => {
    const lease = session.idLease;
    if (lease === undefined) return;
    for (const kind of Object.values(ID_LEASE_KINDS))
      if (!renewing.has(kind) && idLeaseRemaining(lease, kind) < renewBelow)
        void renewKind(kind);
  };

  const onCommitted = (committed: CanonicalCommittedBatch): void => {
    if (stopped || committed.origin === "remote") return;
    const enqueued = client.enqueue({
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      opId: committed.opId,
      actor,
      baseRev: client.headRev,
      ops: committed.ops,
    });
    if (enqueued.isErr()) {
      diverge(enqueued.error.message, enqueued.error);
      return;
    }
    renew();
    // Called inside the host's publish: only submit, never touch the editor.
    pump();
  };

  const subscribe = (): void => {
    unsubscribe();
    unsubscribe = transport.subscribe(client.headRev, (batch) => {
      inbox.push({ type: "sequenced", batch });
      drain();
    });
  };

  const onConnection = (connected: boolean): void => {
    if (stopped || !connected) return;
    subscribe();
    if (awaiting === undefined) {
      pump();
      return;
    }
    // The same batch under the same id: the room answers it once.
    send(
      fault === "freshOpIdOnResubmit"
        ? { ...awaiting, opId: `${awaiting.opId}~${(resubmissions += 1)}` }
        : awaiting,
    );
  };

  const offCommitted = session.onCommitted(onCommitted);
  const offConnection = transport.onConnection(onConnection);

  const leave = (): void => {
    if (stopped) return;
    stopped = true;
    unsubscribe();
    offCommitted();
    offConnection();
    presence?.leave();
  };

  if (transport.connected) subscribe();
  renew();
  drain();

  return Result.ok({
    session,
    get headRev() {
      return client.headRev;
    },
    get pending() {
      return client.pending;
    },
    get stopped() {
      return stopped;
    },
    presence: presence?.api,
    flush: drain,
    leave,
  });
};

type FollowPresenceOptions = {
  channel: NonNullable<CanonicalRoomTransport["presence"]>;
  actor: string;
  user: CanonicalPresenceUser;
  session: CanonicalSession;
  editor: CanonicalCollaborationEditor;
  headRev: () => number;
};

/**
 * This peer's presence on the channel and every other peer's, moved over
 * each sequenced batch the session applies.
 */
const followPresence = ({
  channel,
  actor,
  user,
  session,
  editor,
  headRev,
}: FollowPresenceOptions) => {
  const tracker = createPresenceTracker(actor);
  // Sequenced batches received and not applied yet, by revision.
  const unapplied = new Map<number, SequencedBatch>();
  const listeners = new Set<() => void>();
  const instance = crypto.randomUUID();
  let clock = 0;
  let published: string | undefined;
  let left = false;
  const changed = (): void => {
    for (const listener of listeners) listener();
  };

  const publish = (): void => {
    if (left) return;
    const shown = session.projection.selectionAt(editor.state());
    const selection = shown.isOk()
      ? { anchor: shown.value.anchor, head: shown.value.head }
      : null;
    const rev = headRev();
    const key = JSON.stringify([rev, selection]);
    if (key === published) return;
    published = key;
    clock += 1;
    channel.publish({ actor, instance, clock, user, rev, selection });
  };

  const shownAt = (
    selection: CanonicalRemotePresence["selection"],
  ): CanonicalRemotePeer["shown"] => {
    if (selection === undefined) return undefined;
    const anchor = session.projection.positionAt(selection.anchor);
    const head = session.projection.positionAt(selection.head);
    return anchor.isOk() && head.isOk()
      ? { anchor: anchor.value, head: head.value }
      : undefined;
  };

  const unsubscribe = channel.subscribe((peer, state) => {
    if (!left && tracker.receive(peer, state)) changed();
  });

  return {
    received(batch: SequencedBatch): void {
      unapplied.set(batch.revision, batch);
    },
    /** The client went from revision `from` to `to`. */
    applied(from: number, to: number): void {
      let moved = false;
      for (let revision = from + 1; revision <= to; revision += 1) {
        const batch = unapplied.get(revision);
        unapplied.delete(revision);
        if (batch !== undefined && tracker.applied(batch)) moved = true;
      }
      if (moved) changed();
    },
    publish,
    leave(): void {
      if (left) return;
      left = true;
      unsubscribe();
      channel.publish(null);
    },
    api: {
      update: publish,
      peers: () =>
        tracker
          .peers()
          .map(({ actor: peerActor, user: peerUser, selection }) => ({
            actor: peerActor,
            user: peerUser,
            selection,
            shown: shownAt(selection),
          })),
      onChange: (listener: () => void) => {
        listeners.add(listener);
        return () => {
          listeners.delete(listener);
        };
      },
    } satisfies CanonicalCollaborationPresence,
  };
};
