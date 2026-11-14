import { panic } from "better-result";
import { expect, test } from "bun:test";
import type { EditorState } from "prosemirror-state";

import {
  type CanonicalSession,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { parseDocx } from "../docx/parser";
import { readModel } from "../types/documentModel";
import type { RemoteSelection } from "../types/remote-selection";
import {
  documentText,
  homeState,
  paragraphs,
  peerSeed,
} from "./__tests__/canonicalPeerEdits";
import { settle } from "./__tests__/canonicalRoomPeers";
import type { CanonicalCollaborationEvent } from "./canonicalCollaboration";
import type { CanonicalPresenceChannel } from "./canonicalPresence";
import {
  type CanonicalRoomCollaboration,
  type CanonicalRoomEditor,
  openCanonicalRoomEditor,
  remoteSelectionColor,
} from "./canonicalRoomEditor";
import {
  type CanonicalMemoryRoom,
  canonicalSnapshotId,
  createCanonicalMemoryRoom,
} from "./canonicalRoomTransport";
import { toWireJson } from "./canonicalRoomWire";

/** Presence channels between the editors of one process. */
const createPresenceHub = () => {
  const subscribers = new Map<string, (peer: string, state: unknown) => void>();
  const states = new Map<string, unknown>();
  return (self: string): CanonicalPresenceChannel => ({
    publish: (state) => {
      const sent = toWireJson(state);
      states.set(self, sent);
      for (const [peer, listener] of subscribers)
        if (peer !== self) queueMicrotask(() => listener(self, sent));
    },
    subscribe: (listener) => {
      subscribers.set(self, listener);
      queueMicrotask(() => {
        for (const [peer, state] of states)
          if (peer !== self) listener(peer, state);
      });
      return () => {
        subscribers.delete(self);
      };
    },
  });
};

type Editor = {
  session: CanonicalSession;
  state: EditorState;
  room: CanonicalRoomEditor;
  selections: RemoteSelection[];
  events: CanonicalCollaborationEvent[];
  closed: boolean;
};

const openRoom = async () => {
  const bytes = await peerSeed();
  const snapshot = await canonicalSnapshotId(bytes);
  const presence = createPresenceHub();
  let room: CanonicalMemoryRoom | undefined;
  let seeded = 0;

  const open = async (
    index: number,
    color: string,
    editable: () => boolean = () => true,
  ): Promise<Editor> => {
    const session = createCanonicalSession(
      await parseDocx(bytes, { preloadFonts: false, detectVariables: false }),
    ).unwrap();
    const home =
      paragraphs(readModel(session.document)).at(index)?.paraId ??
      panic("The seed has a paragraph per editor.");
    const actor = `actor-${index}`;
    const collaboration: CanonicalRoomCollaboration = {
      snapshot,
      actor,
      user: { name: actor, color },
      onEvent: (event) => editor.events.push(event),
      connect: (base) => {
        expect(base.snapshot).toBe(snapshot);
        // The first editor seeds the room, as a tab room's leader would.
        if (room === undefined) {
          seeded += 1;
          room = createCanonicalMemoryRoom({
            document: base.document,
            snapshot,
          });
        }
        const connection = room.connect();
        return Object.assign(connection, {
          presence: presence(actor),
          close: () => {
            editor.closed = true;
            connection.disconnect();
          },
        });
      },
    };
    const editor: Editor = {
      session,
      state: homeState(session, home),
      selections: [],
      events: [],
      closed: false,
      room: openCanonicalRoomEditor({
        collaboration,
        session,
        getState: () => editor.state,
        publish: (commit) => {
          if (!editable()) return panic("The editor no longer shows the room.");
          return publishCanonicalProjection({
            state: editor.state,
            commit,
            session,
          }).map(({ state }) => {
            editor.state = state;
            editor.room.documentChanged();
            return undefined;
          });
        },
        onEditableChange: () => {},
        onRemoteSelections: (selections) => {
          editor.selections = selections;
        },
      }),
    };
    return editor;
  };

  /** Type `text` at the editor's caret, as its input would. */
  const type = (editor: Editor, text: string) => {
    const { from, to } = editor.state.selection;
    const prepared = editor.session
      .prepareReplace(editor.state, { from, to, text, semantic: "typing" })
      .unwrap();
    editor.state = publishCanonicalProjection({
      state: editor.state,
      commit: prepared,
      session: editor.session,
    }).unwrap().state;
    editor.room.selectionChanged();
  };

  return { open, type, seeded: () => seeded };
};

test("an editor takes local edits once the room granted its lease, and shows its peers' edits", async () => {
  const room = await openRoom();
  const first = await room.open(0, "#123456");
  expect(first.room.editable).toBe(false);
  await settle();
  expect(first.room.editable).toBe(true);
  expect(first.session.idLease).toBeDefined();

  const second = await room.open(1, "#654321");
  await settle();
  expect(second.room.editable).toBe(true);
  expect(room.seeded()).toBe(1);

  room.type(first, "[first]");
  room.type(second, "[second]");
  for (let round = 0; round < 10; round += 1) await settle();
  const text = documentText(readModel(first.session.document));
  expect(text).toBe(documentText(readModel(second.session.document)));
  expect(text).toContain("[first]");
  expect(text).toContain("[second]");
  expect(
    first.events.filter(({ type }) => type === "localEditRefused"),
  ).toEqual([]);
});

test("an editor shows its peers' selections at their own positions, in a colour a style may take", async () => {
  const room = await openRoom();
  const first = await room.open(0, "#123456");
  const second = await room.open(1, "url(https://example.invalid/beacon)");
  for (let round = 0; round < 10; round += 1) await settle();

  const shownToFirst = first.selections.map(({ peer, color }) => ({
    peer,
    color,
  }));
  expect(shownToFirst).toEqual([
    { peer: "actor-1", color: "var(--doc-image-selection)" },
  ]);
  const shownToSecond = second.selections.at(0);
  expect(shownToSecond?.peer).toBe("actor-0");
  expect(shownToSecond?.color).toBe("#123456");
  expect(shownToSecond?.head).toBe(first.state.selection.head);
});

test("leaving the room closes its connection, clears the peers' selections and takes no more of their edits", async () => {
  const room = await openRoom();
  let showing = true;
  const first = await room.open(0, "#123456", () => showing);
  const second = await room.open(1, "#654321");
  for (let round = 0; round < 10; round += 1) await settle();
  expect(first.selections.length).toBe(1);

  first.room.leave();
  showing = false;
  expect(first.closed).toBe(true);
  expect(first.selections).toEqual([]);
  expect(first.room.editable).toBe(false);
  const before = documentText(readModel(first.session.document));
  room.type(second, "[after]");
  for (let round = 0; round < 10; round += 1) await settle();
  expect(documentText(readModel(first.session.document))).toBe(before);
});

test("a session takes one lease, and only before its first edit", async () => {
  const bytes = await peerSeed();
  const session = createCanonicalSession(
    await parseDocx(bytes, { preloadFonts: false, detectVariables: false }),
  ).unwrap();
  const lease = { blocks: [] };
  expect(session.adoptIdLease(lease).isOk()).toBe(true);
  expect(session.adoptIdLease(lease).isErr()).toBe(true);
});

test("a peer's colour reaches a style only as a plain hex colour", () => {
  expect(remoteSelectionColor("#a1B2c3")).toBe("#a1B2c3");
  for (const color of ["red", "#abc", "url(x)", "#1234567", 7, undefined])
    expect(remoteSelectionColor(color)).toBe("var(--doc-image-selection)");
});
