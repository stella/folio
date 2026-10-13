import { useCallback, useEffect, useMemo, useSyncExternalStore } from "react";
import * as Y from "yjs";
import { WebrtcProvider } from "y-webrtc";
import { yCursorPlugin, ySyncPlugin, yUndoPlugin } from "y-prosemirror";

import type { Comment } from "@stll/folio-core/types/content";
import type { DocxEditorCollaboration } from "@stll/folio-react";

export type CollaborativeUser = {
  clientId: number;
  name: string;
  color: string;
  isLocal: boolean;
};

export type CollaborationState = {
  collaboration: DocxEditorCollaboration | null;
  users: CollaborativeUser[];
  roomName: string;
  status: "connecting" | "connected" | "disconnected";
  comments: Comment[];
  setComments: (next: Comment[]) => void;
};

const SIGNALING_SERVERS = ["wss://signaling.yjs.dev", "wss://y-webrtc-signaling-eu.herokuapp.com"];

const createCollaborationResources = (roomName: string) => {
  const doc = new Y.Doc();
  const collabProvider = new WebrtcProvider(roomName, doc, { signaling: SIGNALING_SERVERS });
  const xmlFragment = doc.getXmlFragment("prosemirror");
  const collabPlugins = [
    ySyncPlugin(xmlFragment),
    yCursorPlugin(collabProvider.awareness),
    yUndoPlugin(),
  ];
  const commentsArray = doc.getArray<Comment>("comments");
  return {
    ydoc: doc,
    provider: collabProvider,
    plugins: collabPlugins,
    yComments: commentsArray,
    yXmlFragment: xmlFragment,
  };
};

type CollaborationResources = ReturnType<typeof createCollaborationResources>;

const syncYComments = (yComments: Y.Array<Comment>, next: Comment[]): void => {
  const nextIds = new Set(next.map((comment) => comment.id));

  for (let i = yComments.length - 1; i >= 0; i--) {
    if (!nextIds.has(yComments.get(i).id)) {
      yComments.delete(i, 1);
    }
  }

  const indexById = new Map(yComments.toArray().map((comment, index) => [comment.id, index]));

  for (const comment of next) {
    const index = indexById.get(comment.id);
    if (index === undefined) {
      yComments.push([comment]);
      continue;
    }
    const existing = yComments.get(index);
    if (JSON.stringify(existing) !== JSON.stringify(comment)) {
      yComments.delete(index, 1);
      yComments.insert(index, [comment]);
    }
  }
};

type CollaborationSnapshot = Pick<
  CollaborationState,
  "collaboration" | "users" | "status" | "comments"
>;

/** Connection resources exist only while React is subscribed, including StrictMode remounts. */
export const createCollaborationStore = (createResources: () => CollaborationResources) => {
  const disconnected: CollaborationSnapshot = {
    collaboration: null,
    users: [],
    status: "connecting",
    comments: [],
  };
  let snapshot = disconnected;
  let resources: CollaborationResources | null = null;
  const listeners = new Set<() => void>();
  let disconnect: (() => void) | null = null;
  const publish = (next: CollaborationSnapshot) => {
    snapshot = next;
    for (const listener of listeners) listener();
  };
  return {
    getSnapshot: () => snapshot,
    subscribe: (listener: () => void) => {
      listeners.add(listener);
      if (resources === null) {
        const connection = createResources();
        resources = connection;
        const { provider, yComments } = connection;
        const refreshUsers = () => {
          const users: CollaborativeUser[] = [];
          provider.awareness.getStates().forEach((state, clientId) => {
            const user = (state as { user?: { name: string; color: string } }).user;
            if (user)
              users.push({ clientId, ...user, isLocal: clientId === provider.awareness.clientID });
          });
          publish({ ...snapshot, users });
        };
        const handleStatus = (event: { connected: boolean }) => {
          publish({ ...snapshot, status: event.connected ? "connected" : "disconnected" });
        };
        const refreshComments = () => publish({ ...snapshot, comments: yComments.toArray() });
        provider.awareness.on("change", refreshUsers);
        provider.on("status", handleStatus);
        yComments.observeDeep(refreshComments);
        publish({
          collaboration: {
            yXmlFragment: connection.yXmlFragment,
            plugins: connection.plugins,
            awareness: provider.awareness,
            shouldSeed: true,
          },
          users: [],
          status: "connecting",
          comments: yComments.toArray(),
        });
        refreshUsers();
        disconnect = () => {
          provider.awareness.off("change", refreshUsers);
          provider.off("status", handleStatus);
          yComments.unobserveDeep(refreshComments);
          provider.destroy();
          connection.ydoc.destroy();
        };
      }
      return () => {
        listeners.delete(listener);
        if (listeners.size > 0) return;
        disconnect?.();
        disconnect = null;
        resources = null;
        snapshot = disconnected;
      };
    },
    setUser: (user: { name: string; color: string }) => {
      resources?.provider.awareness.setLocalStateField("user", user);
    },
    setComments: (next: Comment[]) => {
      const connection = resources;
      if (connection === null) return;
      connection.ydoc.transact(() => syncYComments(connection.yComments, next));
    },
  };
};

export const useCollaboration = (
  roomName: string,
  localUser: { name: string; color: string },
): CollaborationState => {
  const store = useMemo(
    () => createCollaborationStore(() => createCollaborationResources(roomName)),
    [roomName],
  );
  const snapshot = useSyncExternalStore(store.subscribe, store.getSnapshot);
  useEffect(() => {
    store.setUser({ name: localUser.name, color: localUser.color });
  }, [localUser.color, localUser.name, store]);
  const setComments = useCallback((next: Comment[]) => store.setComments(next), [store]);
  return { ...snapshot, roomName, setComments };
};
