import { panic } from "better-result";
import { expect, test } from "bun:test";

import { paragraphLogicalText } from "@stll/docx-core/ops";

import { paragraphs } from "./__tests__/canonicalPeerEdits";
import {
  createRoomPeers,
  emptyRoomCoverage,
  ROOM_BLOCK_SIZES,
  type RoomPeer,
  settle,
} from "./__tests__/canonicalRoomPeers";
import type { CanonicalPresenceChannel } from "./canonicalPresence";
import {
  type CanonicalRoomTransport,
  createCanonicalMemoryRoom,
} from "./canonicalRoomTransport";
import { toWireJson } from "./canonicalRoomWire";

/** Presence channels between peers of one process, each delivery a microtask. */
const createPresenceHub = () => {
  const listeners = new Map<
    string,
    Set<(peer: string, state: unknown) => void>
  >();
  const states = new Map<string, unknown>();
  return (self: string): CanonicalPresenceChannel => ({
    publish: (state) => {
      const sent = toWireJson(state);
      states.set(self, sent);
      for (const [peer, subscribers] of listeners)
        if (peer !== self)
          for (const listener of subscribers)
            queueMicrotask(() => listener(self, sent));
    },
    subscribe: (listener) => {
      const subscribers = listeners.get(self) ?? new Set();
      subscribers.add(listener);
      listeners.set(self, subscribers);
      queueMicrotask(() => {
        for (const [peer, state] of states)
          if (peer !== self) listener(peer, state);
      });
      return () => {
        subscribers.delete(listener);
      };
    },
  });
};

const withPresence = (
  transport: CanonicalRoomTransport,
  presence: CanonicalPresenceChannel,
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
  submit: (batch) => transport.submit(batch),
  subscribe: (afterRev, onSequenced) =>
    transport.subscribe(afterRev, onSequenced),
  onConnection: (listener) => transport.onConnection(listener),
  grantIdLease: (actor) => transport.grantIdLease(actor),
  renewIdLease: (actor, kind) => transport.renewIdLease(actor, kind),
  presence,
});

test("a peer's selection moves only once the edit that moves it is shown", async () => {
  const room = await createRoomPeers(emptyRoomCoverage());
  const memory = createCanonicalMemoryRoom({
    document: room.base,
    snapshot: room.snapshot,
    sequencer: {
      idLeases: { blockSizes: ROOM_BLOCK_SIZES, grants: [], minted: [] },
    },
  });
  const channel = createPresenceHub();
  const peers: RoomPeer[] = [];
  for (let index = 0; index < 3; index += 1)
    peers.push(
      await room.join({
        index,
        incarnation: 0,
        transport: withPresence(memory.connect(), channel(`tab-${index}`)),
      }),
    );
  for (let round = 0; round < 4; round += 1) await settle();
  const [owner, viewer, editor] = peers;
  const presence = viewer?.collaboration.presence;
  if (
    owner === undefined ||
    viewer === undefined ||
    editor === undefined ||
    presence === undefined
  )
    return panic("Three peers with presence joined.");
  const ownerParagraph = () =>
    paragraphs(viewer.session.document).find(
      ({ paraId }) => paraId === owner.home,
    );

  // What a host sees each time it is told the peers moved.
  const seen: { offset: number; prefixed: boolean }[] = [];
  presence.onChange(() => {
    const head = presence.peers().find(({ actor }) => actor === owner.actor)
      ?.selection?.head;
    if (head === undefined) return;
    const paragraph = ownerParagraph();
    seen.push({
      offset: head.offset,
      prefixed:
        paragraph !== undefined &&
        paragraphLogicalText(paragraph).startsWith(String(editor.index)),
    });
  });
  const before = presence.peers().find(({ actor }) => actor === owner.actor)
    ?.selection?.head.offset;
  expect(before).toBeDefined();

  // The third peer puts its digit at the start of the owner's paragraph,
  // ahead of the owner's caret.
  room.edit(editor, peers, "prefix", 0);
  for (let round = 0; round < 8; round += 1) await settle();

  const moved = seen.filter(({ offset }) => offset === (before ?? 0) + 1);
  expect(moved.length).toBeGreaterThan(0);
  // Each time the host heard of the moved caret, the edit was already shown.
  for (const { prefixed } of moved) expect(prefixed).toBe(true);
  const now = presence.peers().find(({ actor }) => actor === owner.actor);
  const held = owner.session.projection.selectionAt(owner.state).unwrap();
  expect(now?.selection?.head).toEqual(held.head);
  expect(now?.shown?.head).toBe(
    viewer.session.projection.positionAt(held.head).unwrap(),
  );
});
