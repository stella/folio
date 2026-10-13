import { expect, test } from "bun:test";
import * as Y from "yjs";
import { WebrtcProvider } from "y-webrtc";

import type { Comment } from "@stll/folio-core/types/content";
import { createCollaborationStore } from "./useCollaboration";

test("collaboration subscriptions own connections and publish external updates", () => {
  const createResources = () => {
    const ydoc = new Y.Doc();
    // No remote signaling: this exercises the real provider's events locally.
    const provider = new WebrtcProvider(`test-${ydoc.clientID}`, ydoc, { signaling: [] });
    provider.disconnect();
    const connection = {
      ydoc,
      provider,
      plugins: [],
      yComments: ydoc.getArray<Comment>("comments"),
      yXmlFragment: ydoc.getXmlFragment("prosemirror"),
    };
    return connection;
  };
  const connections: ReturnType<typeof createResources>[] = [];
  const store = createCollaborationStore(() => {
    const connection = createResources();
    connections.push(connection);
    return connection;
  });
  expect(connections).toHaveLength(0);
  expect(store.getSnapshot().collaboration).toBeNull();
  let notifications = 0;
  const unsubscribe = store.subscribe(() => notifications++);
  const secondUnsubscribe = store.subscribe(() => notifications++);
  const first = connections.at(0);
  if (!first) throw new Error("subscription did not create its connection");
  expect(connections).toHaveLength(1);
  store.setUser({ name: "Local", color: "blue" });
  expect(store.getSnapshot().users).toEqual([
    { clientId: first.provider.awareness.clientID, name: "Local", color: "blue", isLocal: true },
  ]);
  first.provider.emit("status", [{ connected: true }]);
  expect(store.getSnapshot().status).toBe("connected");
  const comment = { id: 1, author: "Local", content: [] } satisfies Comment;
  store.setComments([comment]);
  expect(store.getSnapshot().comments).toEqual([comment]);
  first.yComments.delete(0, 1);
  expect(store.getSnapshot().comments).toEqual([]);
  expect(notifications).toBeGreaterThan(0);
  unsubscribe();
  expect(first.ydoc.isDestroyed).toBe(false);
  secondUnsubscribe();
  expect(first.ydoc.isDestroyed).toBe(true);
  expect(store.getSnapshot()).toMatchObject({
    collaboration: null,
    users: [],
    status: "connecting",
    comments: [],
  });

  // StrictMode's unsubscribe/resubscribe cycle must create a fresh live connection.
  const unsubscribeAgain = store.subscribe(() => notifications++);
  expect(connections).toHaveLength(2);
  expect(store.getSnapshot().collaboration?.yXmlFragment.doc).not.toBe(first.ydoc);
  unsubscribeAgain();
  expect(connections.at(1)?.ydoc.isDestroyed).toBe(true);
});
