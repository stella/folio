import { useCallback, useEffect, useMemo, useState } from "react";
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
  collaboration: DocxEditorCollaboration;
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

export const useCollaboration = (
  roomName: string,
  localUser: { name: string; color: string },
): CollaborationState => {
  const [{ ydoc, provider, plugins, yComments, yXmlFragment }] = useState(() =>
    createCollaborationResources(roomName),
  );

  const [users, setUsers] = useState<CollaborativeUser[]>([]);
  const [status, setStatus] = useState<CollaborationState["status"]>("connecting");
  const [comments, setCommentsState] = useState<Comment[]>(() => yComments.toArray());

  const collaboration = useMemo(
    (): DocxEditorCollaboration => ({
      yXmlFragment,
      plugins,
      awareness: provider.awareness,
      shouldSeed: true,
    }),
    [plugins, provider.awareness, yXmlFragment],
  );

  useEffect(() => {
    provider.awareness.setLocalStateField("user", localUser);
  }, [localUser.color, localUser.name, provider]);

  useEffect(() => {
    const refreshUsers = () => {
      const localId = provider.awareness.clientID;
      const all: CollaborativeUser[] = [];
      provider.awareness.getStates().forEach((state, clientId) => {
        const user = (state as { user?: { name: string; color: string } }).user;
        if (!user) {
          return;
        }
        all.push({
          clientId,
          name: user.name,
          color: user.color,
          isLocal: clientId === localId,
        });
      });
      setUsers(all);
    };
    const handleStatus = (event: { connected: boolean }) => {
      setStatus(event.connected ? "connected" : "disconnected");
    };

    refreshUsers();
    provider.awareness.on("change", refreshUsers);
    provider.on("status", handleStatus);

    return () => {
      provider.awareness.off("change", refreshUsers);
      provider.off("status", handleStatus);
    };
  }, [provider]);

  useEffect(() => {
    const sync = () => setCommentsState(yComments.toArray());
    sync();
    yComments.observeDeep(sync);
    return () => yComments.unobserveDeep(sync);
  }, [yComments]);

  const setComments = useCallback(
    (next: Comment[]) => {
      ydoc.transact(() => {
        syncYComments(yComments, next);
      });
    },
    [ydoc, yComments],
  );

  useEffect(
    () => () => {
      provider.destroy();
      ydoc.destroy();
    },
    [provider, ydoc],
  );

  return { collaboration, users, roomName, status, comments, setComments };
};
