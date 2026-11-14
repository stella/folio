import { openCanonicalIndexedDbJournal } from "@stll/folio-core/controller/canonicalIndexedDbJournal";
import {
  type CanonicalPresenceState,
  type CanonicalPresenceUser,
  parseCanonicalPresence,
} from "@stll/folio-core/controller/canonicalPresence";
import type {
  CanonicalRoomCollaboration,
  CanonicalRoomConnection,
} from "@stll/folio-core/controller/canonicalRoomEditor";
import { canonicalSnapshotId } from "@stll/folio-core/controller/canonicalRoomTransport";
import {
  type CanonicalTabRoom,
  openCanonicalTabRoom,
} from "@stll/folio-core/controller/canonicalTabRoom";

import type { CanonicalRoomBridgeStatus } from "./canonicalRoomBridge";

/*
 * The canonical room both playgrounds run with `?collaboration=1`: the tabs
 * of this origin that opened the same DOCX share one room, led by one of
 * them, its journal in IndexedDB. A reload keeps the tab's actor, and the
 * room survives every tab closing.
 */

const ACTOR_KEY = "folio-playground-room-actor";

/** This tab's actor: kept for the tab's life, reloads included. */
const tabActor = (): string => {
  try {
    const stored = sessionStorage.getItem(ACTOR_KEY);
    if (stored) return stored;
    const actor = crypto.randomUUID();
    sessionStorage.setItem(ACTOR_KEY, actor);
    return actor;
  } catch {
    return crypto.randomUUID();
  }
};

export type CanonicalRoomDemoPeer = CanonicalPresenceUser & {
  actor: string;
  isLocal: boolean;
};

export type CanonicalRoomDemoState = CanonicalRoomBridgeStatus & {
  /** This tab's user first, then every other peer's. */
  users: readonly CanonicalRoomDemoPeer[];
};

export type CanonicalRoomDemo = {
  /** What the editor joins the room with; stable for the demo's life. */
  collaboration: CanonicalRoomCollaboration;
  /** The room as it stands now, read live. */
  status: () => CanonicalRoomBridgeStatus;
  /** The room as of the last change, for a view to render. */
  getState: () => CanonicalRoomDemoState;
  /** Returns the unsubscribe. */
  subscribe: (listener: () => void) => () => void;
};

/** The room `room` of the DOCX `bytes`, joined as `user`. */
export const createCanonicalRoomDemo = async ({
  room,
  bytes,
  user,
}: {
  room: string;
  bytes: ArrayBuffer;
  user: CanonicalPresenceUser;
}): Promise<CanonicalRoomDemo> => {
  const snapshot = await canonicalSnapshotId(bytes);
  const actor = tabActor();
  const peers = new Map<string, CanonicalPresenceState>();
  const listeners = new Set<() => void>();
  let connection: CanonicalTabRoom | undefined;
  let notice: string | undefined;
  const status = (): CanonicalRoomBridgeStatus => ({
    connected: connection?.connected ?? false,
    leading: connection?.leading ?? false,
    peers: [...peers.values()].map((peer) => peer.actor),
    notice,
  });
  const read = (): CanonicalRoomDemoState => ({
    ...status(),
    users: [
      { ...user, actor, isLocal: true },
      ...[...peers.values()].map((peer) => ({
        ...peer.user,
        actor: peer.actor,
        isLocal: false,
      })),
    ],
  });
  let state = read();
  const changed = (): void => {
    state = read();
    for (const listener of listeners) listener();
  };

  const connect = ({
    document,
  }: Parameters<CanonicalRoomCollaboration["connect"]>[0]) => {
    const opened = openCanonicalTabRoom({
      room,
      snapshot,
      document,
      journal: openCanonicalIndexedDbJournal(`${room}:${snapshot}`),
      onError: (error) => {
        notice = error.message;
        changed();
      },
    });
    connection = opened;
    peers.clear();
    const offConnection = opened.onConnection(changed);
    const offPresence = opened.presence.subscribe((peer, value) => {
      const parsed = value === null ? undefined : parseCanonicalPresence(value);
      if (parsed === undefined) peers.delete(peer);
      else peers.set(peer, parsed);
      changed();
    });
    const close = opened.close;
    const closing: CanonicalRoomConnection = Object.assign(opened, {
      close: () => {
        offConnection();
        offPresence();
        close();
        if (connection === opened) connection = undefined;
        peers.clear();
        changed();
      },
    });
    changed();
    return closing;
  };

  return {
    collaboration: {
      snapshot,
      actor,
      user,
      connect,
      onEvent: (event) => {
        switch (event.type) {
          case "synced":
            return;
          case "stopped":
            notice = `Stopped following the room: ${event.error.message}`;
            break;
          case "localEditRefused":
            notice = "An edit could not merge and was taken back.";
            break;
          case "leaseRefused":
            notice = "No more ids of a kind are left in this room.";
            break;
        }
        changed();
      },
    },
    status,
    getState: () => state,
    subscribe: (listener) => {
      listeners.add(listener);
      return () => {
        listeners.delete(listener);
      };
    },
  };
};
