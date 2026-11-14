/**
 * Peers for the room adapter properties: each peer is a real canonical
 * session joined to a room through `joinCanonicalRoom`, over whatever
 * transport the property builds, editing with `localEdit`'s concurrent
 * edits (./canonicalPeerEdits.ts). The adapter properties differ in their
 * networks; joining, editing and the final checks are shared here.
 */

import { panic, Result } from "better-result";
import { expect } from "bun:test";

import type { SequencedBatch } from "@stll/docx-core/ops";

import {
  type CanonicalCommit,
  type CanonicalSessionError,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../../canonical/canonicalSession";
import { parseDocx } from "../../docx/parser";
import type { Document } from "../../types/document";
import {
  type CanonicalCollaboration,
  type CanonicalCollaborationEvent,
  joinCanonicalRoom,
} from "../canonicalCollaboration";
import type {
  CanonicalRoomSubmission,
  CanonicalRoomTransport,
} from "../canonicalRoomTransport";
import { canonicalSnapshotId } from "../canonicalRoomTransport";
import {
  type BatchMeta,
  caretOf,
  documentText,
  type EditingPeer,
  expectIdenticalSaves,
  expectSequencedKeystrokes,
  homeState,
  localEdit,
  type LocalKind,
  paragraphs,
  peerSeed,
} from "./canonicalPeerEdits";

export type RoomPeer = EditingPeer & {
  actor: string;
  collaboration: CanonicalCollaboration;
  events: CanonicalCollaborationEvent[];
  refused: Set<string>;
  /** Commits this incarnation made, by operation id. */
  local: Set<string>;
  /** A pending batch was dropped: the caret's continuity is not promised. */
  replaced: boolean;
};

export type RoomCoverage = {
  remoteApplied: number;
  undoAfterRemote: number;
  resubmitted: number;
  lateJoins: number;
};

export const emptyRoomCoverage = (): RoomCoverage => ({
  remoteApplied: 0,
  undoAfterRemote: 0,
  resubmitted: 0,
  lateJoins: 0,
});

/** Let every promise continuation run. */
export const settle = () =>
  new Promise<void>((resolve) => setImmediate(resolve));

export const ROOM_BLOCK_SIZES = {
  paragraph: 8,
  revision: 64,
  control: 16,
  comment: 16,
} as const;

/**
 * A transport that also records every acknowledgement it carried while its
 * peer was alive: a simulated dead tab's code still runs, a real one's does
 * not.
 */
const recordingAcks = (
  transport: CanonicalRoomTransport,
  acked: Map<string, number>,
  alive: () => boolean,
): CanonicalRoomTransport => ({
  get snapshot() {
    return transport.snapshot;
  },
  get headRev() {
    return transport.headRev;
  },
  get connected() {
    return transport.connected;
  },
  submit: async (batch): Promise<CanonicalRoomSubmission> => {
    const outcome = await transport.submit(batch);
    if (outcome.type === "ack" && alive()) acked.set(outcome.opId, outcome.rev);
    return outcome;
  },
  subscribe: (afterRev, onSequenced) =>
    transport.subscribe(afterRev, onSequenced),
  onConnection: (listener) => transport.onConnection(listener),
  grantIdLease: (actor) => transport.grantIdLease(actor),
  renewIdLease: (actor, kind) => transport.renewIdLease(actor, kind),
  ...(transport.presence === undefined ? {} : { presence: transport.presence }),
});

/** The seed document, the room's base, and every peer's bookkeeping. */
export const createRoomPeers = async (coverage: RoomCoverage) => {
  const bytes = await peerSeed();
  const parse = () =>
    parseDocx(bytes, { preloadFonts: false, detectVariables: false });
  const snapshot = await canonicalSnapshotId(bytes);
  const base = createCanonicalSession(await parse()).unwrap().document;
  const batches = new Map<string, BatchMeta>();
  const acked = new Map<string, number>();
  const typed = new Map<number, number>();
  let nextMeta: BatchMeta | undefined;

  /** Join peer `index` as `actor`; resolves once the room granted its lease. */
  const join = async ({
    index,
    incarnation,
    transport,
    alive = () => true,
  }: {
    index: number;
    incarnation: number;
    transport: CanonicalRoomTransport;
    /** False once the simulated context running this peer is gone. */
    alive?: () => boolean;
  }): Promise<RoomPeer> => {
    const actor = `actor-${index}`;
    const document = await parse();
    const refused = new Set<string>();
    const events: CanonicalCollaborationEvent[] = [];
    let ops = 0;
    const self: { peer?: RoomPeer; state?: EditingPeer["state"] } = {};
    const joined = await joinCanonicalRoom({
      transport: recordingAcks(transport, acked, alive),
      actor,
      snapshot,
      renewBelow: ROOM_BLOCK_SIZES.paragraph / 2,
      presence: { user: { name: actor, color: `#00000${index}` } },
      onEvent: (event) => {
        events.push(event);
        if (event.type !== "localEditRefused") return;
        refused.add(event.opId);
        if (self.peer !== undefined) self.peer.replaced = true;
      },
      open: (idLease) =>
        createCanonicalSession(document, undefined, {
          idLease,
          newOpId: () => `${actor}.${incarnation}-${ops++}`,
        }).map((session) => {
          const home = paragraphs(session.document).at(index)?.paraId;
          if (home === undefined)
            return panic("The seed has a paragraph per peer.");
          self.state = homeState(session, home);
          return {
            session,
            editor: {
              state: () => self.peer?.state ?? self.state ?? panic("No state."),
              publish: (commit: CanonicalCommit) => {
                const state = self.peer?.state ?? self.state;
                if (state === undefined) return panic("No state.");
                const published = publishCanonicalProjection({
                  state,
                  commit,
                  session,
                });
                if (published.isErr()) return published;
                if (self.peer === undefined) self.state = published.value.state;
                else self.peer.state = published.value.state;
                if (commit.origin === "remote") coverage.remoteApplied += 1;
                return Result.ok(undefined);
              },
            },
          };
        }),
    });
    const collaboration = joined.unwrap();
    const { session } = collaboration;
    const home = paragraphs(base).at(index)?.paraId;
    if (home === undefined || self.state === undefined)
      return panic("The peer opened its session.");
    const peer: RoomPeer = {
      index,
      actor,
      home,
      session,
      state: self.state,
      history: [],
      typed: typed.get(index) ?? 0,
      collaboration,
      events,
      refused,
      local: new Set(),
      replaced: incarnation > 0,
    };
    self.peer = peer;
    session.onCommitted((committed) => {
      if (committed.origin === "remote") {
        for (const step of peer.history) step.rebased = true;
        return;
      }
      if (nextMeta === undefined) return panic("A commit nobody made.");
      batches.set(committed.opId, nextMeta);
      peer.local.add(committed.opId);
    });
    session.onHistoryDrop(({ undo }) => {
      peer.history.splice(0, undo);
      peer.replaced = true;
    });
    collaboration.presence?.update();
    return peer;
  };

  const commitLocal = (
    peer: RoomPeer,
    prepared: Result<CanonicalCommit, CanonicalSessionError>,
    meta: BatchMeta,
  ): CanonicalCommit | undefined => {
    if (prepared.isErr()) return undefined;
    nextMeta = meta;
    peer.state = publishCanonicalProjection({
      state: peer.state,
      commit: prepared.value,
      session: peer.session,
    }).unwrap().state;
    nextMeta = undefined;
    return prepared.value;
  };

  /** One local edit by `peer`, then its presence, as a host publishes it. */
  const edit = (
    peer: RoomPeer,
    peers: readonly RoomPeer[],
    kind: LocalKind,
    seed: number,
  ): void => {
    localEdit({
      peer,
      peers,
      kind,
      choice: seed,
      commit: commitLocal,
      onUndoAfterRemote: () => {
        coverage.undoAfterRemote += 1;
      },
    });
    typed.set(peer.index, peer.typed);
    peer.collaboration.presence?.update();
  };

  /**
   * The checks once the network is drained: every live peer follows the
   * room, holds the authority's text and paragraph ids and saves the same
   * bytes; every keystroke the room sequenced (less the undone ones)
   * survives; no acknowledged edit is missing from the journal; every edit
   * of a live peer was sequenced or reported refused; and every peer sees
   * every other peer's selection where its owner holds it.
   */
  const expectConverged = async ({
    peers,
    peersCount,
    document,
    journal,
    contiguous,
  }: {
    peers: readonly RoomPeer[];
    peersCount: number;
    /** The authority's document. */
    document: Document;
    journal: readonly SequencedBatch[];
    contiguous: (index: number) => boolean;
  }): Promise<void> => {
    const finalText = documentText(document);
    const finalIds = paragraphs(document).map(({ paraId }) => paraId);
    for (const peer of peers) {
      expect(peer.events.filter(({ type }) => type === "stopped")).toEqual([]);
      expect(peer.collaboration.pending).toHaveLength(0);
      expect(peer.collaboration.headRev).toBe(journal.length);
      expect(documentText(peer.session.document)).toBe(finalText);
      expect(
        paragraphs(peer.session.document).map(({ paraId }) => paraId),
      ).toEqual(finalIds);
    }
    await expectIdenticalSaves(peers.map(({ session }) => session));
    expectSequencedKeystrokes({
      finalText,
      journal,
      batches,
      peers: peersCount,
      contiguous: (index) =>
        contiguous(index) &&
        peers.find((peer) => peer.index === index)?.replaced === false,
    });
    for (const [opId, rev] of acked)
      expect(journal.at(rev - 1)?.opId).toBe(opId);
    const sequenced = new Set(journal.map(({ opId }) => opId));
    for (const peer of peers)
      for (const opId of peer.local)
        expect(sequenced.has(opId) || peer.refused.has(opId)).toBe(true);
    // oxlint-disable-next-line no-console
    console.log(
      `TABDIAG ${JSON.stringify({
        journal: journal.map(({ revision, opId, actor, ops }) => ({
          revision,
          opId,
          actor,
          ops: ops.map((op) => JSON.stringify(op).slice(0, 220)),
        })),
        peers: peers.map((peer) => ({
          actor: peer.actor,
          index: peer.index,
          held: caretOf(peer),
          replaced: peer.replaced,
          refused: [...peer.refused],
          local: [...peer.local],
          events: peer.events.map((event) =>
            JSON.stringify(event).slice(0, 200),
          ),
          seen: peer.collaboration.presence?.peers(),
        })),
      })}`,
    );
    for (const viewer of peers)
      for (const owner of peers) {
        if (owner === viewer) continue;
        const held = caretOf(owner);
        const seen = viewer.collaboration.presence
          ?.peers()
          .find(({ actor }) => actor === owner.actor);
        expect(seen?.selection).toEqual({
          anchor: held.anchor,
          head: held.head,
        });
        expect(seen?.shown).toEqual({
          anchor: viewer.session.projection.positionAt(held.anchor).unwrap(),
          head: viewer.session.projection.positionAt(held.head).unwrap(),
        });
      }
  };

  return { snapshot, base, join, edit, expectConverged };
};
