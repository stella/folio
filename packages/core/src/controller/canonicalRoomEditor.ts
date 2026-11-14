import type { Result } from "better-result";
import { EditorState } from "prosemirror-state";

import type {
  CanonicalCommit,
  CanonicalSession,
  CanonicalSessionError,
} from "../canonical/canonicalSession";
import type { Document } from "../types/document";
import { readModel } from "../types/documentModel";
import type { RemoteSelection } from "../types/remote-selection";
import {
  type CanonicalCollaboration,
  type CanonicalCollaborationEvent,
  joinCanonicalRoom,
} from "./canonicalCollaboration";
import type { CanonicalPresenceUser } from "./canonicalPresence";
import type { CanonicalRoomTransport } from "./canonicalRoomTransport";

/*
 * An editor's place in a canonical room: the hidden editor's canonical
 * session joined to the room through the collaboration controller. The
 * editor shows the session as soon as it opens; it takes local edits once
 * the room granted the session its id lease, and takes none again once the
 * session stopped following the room.
 */

/** What a canonical room is seeded with: revision 0. */
export type CanonicalRoomBase = {
  /** The document the canonical session reads from the room's DOCX. */
  document: Document;
  /** `canonicalSnapshotId` of that DOCX. */
  snapshot: string;
};

/** A transport to one room, which the editor closes when it leaves. */
export type CanonicalRoomConnection = CanonicalRoomTransport & {
  close(): void;
};

/**
 * How an editor running the canonical session joins a room: the
 * `collaboration` a host passes. Keep the object stable while the editor
 * stays in the room; another object leaves the room and joins the one it
 * names.
 */
export type CanonicalRoomCollaboration = {
  /**
   * `canonicalSnapshotId` of the DOCX the editor opens. A room seeded from
   * other bytes refuses the session.
   */
  snapshot: string;
  /**
   * This peer's identity in the room: distinct for each open editor, and
   * the same when it opens again (a reload), so the room hands it back its
   * own id lease.
   */
  actor: string;
  /**
   * Reach the room once the editor has read its document. Called once per
   * document the editor opens; the editor closes the connection when it
   * opens another document, when `collaboration` changes and when it
   * unmounts. A room that runs its own authority (a tab room) is seeded
   * with `base`.
   */
  connect: (base: CanonicalRoomBase) => CanonicalRoomConnection;
  /**
   * Shown beside this peer's selection in the other peers' editors. Without
   * it the editor neither shares its selection nor shows the others'.
   */
  user?: CanonicalPresenceUser;
  /** Every event of the room: a local edit it refused, a lease, a stop. */
  onEvent?: (event: CanonicalCollaborationEvent) => void;
  /** A Yjs fragment binds the legacy engine; a canonical room has none. */
  yXmlFragment?: never;
};

/** Whether a `collaboration` value names a canonical room. */
export const isCanonicalRoomCollaboration = <Other extends object>(
  collaboration: Other | CanonicalRoomCollaboration,
): collaboration is CanonicalRoomCollaboration => "connect" in collaboration;

/**
 * A collaborator's colour as a selection overlay may paint it. It arrives
 * from another peer, so only a plain 6-digit hex colour passes: a crafted
 * value (a CSS `url(...)`, for one) never reaches a style.
 */
export const remoteSelectionColor = (color: unknown): string =>
  typeof color === "string" && /^#[0-9A-Fa-f]{6}$/u.test(color)
    ? color
    : "var(--doc-image-selection)";

export type CanonicalRoomEditorOptions = {
  collaboration: CanonicalRoomCollaboration;
  /** The editor's session, unedited. */
  session: CanonicalSession;
  /** The state the editor shows the session in, once its view exists. */
  getState: () => EditorState | null;
  /**
   * Show another peer's edit as a local commit is shown, whether or not the
   * editor takes local edits; refused once the editor no longer shows
   * `session`.
   */
  publish: (commit: CanonicalCommit) => Result<void, CanonicalSessionError>;
  /** Called when `editable` may have changed. */
  onEditableChange: () => void;
  /** The other peers' selections, in the editor's positions. */
  onRemoteSelections: (selections: RemoteSelection[]) => void;
};

export type CanonicalRoomEditor = {
  /** Whether local edits reach the room: joined, and still following it. */
  readonly editable: boolean;
  /** The editor's own selection moved. */
  selectionChanged(): void;
  /** The editor's document changed, so the peers' positions in it moved. */
  documentChanged(): void;
  /** A composition ended: apply the peers' edits that waited for it. */
  flush(): void;
  /** Leave the room and close its connection. */
  leave(): void;
};

/** Join `session`, which the editor already shows, to the room `collaboration` names. */
export const openCanonicalRoomEditor = ({
  collaboration,
  session,
  getState,
  publish,
  onEditableChange,
  onRemoteSelections,
}: CanonicalRoomEditorOptions): CanonicalRoomEditor => {
  const { snapshot, actor, user, onEvent } = collaboration;
  let status: "joining" | "joined" | "stopped" = "joining";
  let left = false;
  let joined: CanonicalCollaboration | undefined;
  let offPresence = (): void => {};
  const connection = collaboration.connect({
    document: readModel(session.document),
    snapshot,
  });

  const setStatus = (next: typeof status): void => {
    if (status === next) return;
    status = next;
    onEditableChange();
  };
  const report = (event: CanonicalCollaborationEvent): void => {
    if (left) return;
    if (event.type === "stopped") setStatus("stopped");
    onEvent?.(event);
  };

  const remoteSelections = (): RemoteSelection[] =>
    joined?.presence?.peers().flatMap((peer) =>
      peer.shown === undefined
        ? []
        : [
            {
              peer: peer.actor,
              name: peer.user.name,
              color: remoteSelectionColor(peer.user.color),
              anchor: peer.shown.anchor,
              head: peer.shown.head,
            },
          ],
    ) ?? [];
  const showRemoteSelections = (): void => {
    if (!left) onRemoteSelections(remoteSelections());
  };

  const join = async (): Promise<void> => {
    const result = await joinCanonicalRoom({
      transport: connection,
      actor,
      snapshot,
      ...(user === undefined ? {} : { presence: { user } }),
      onEvent: report,
      // An editor that left meanwhile refuses every publication, so the
      // session stops following the room at its first edit.
      open: (idLease) =>
        session.adoptIdLease(idLease).map(() => ({
          session,
          editor: {
            state: () =>
              getState() ?? EditorState.create({ doc: session.projection.doc }),
            publish,
          },
        })),
    });
    if (left) {
      if (result.isOk()) result.value.leave();
      return;
    }
    if (result.isErr()) {
      report({ type: "stopped", error: result.error });
      return;
    }
    joined = result.value;
    offPresence =
      joined.presence?.onChange(showRemoteSelections) ?? offPresence;
    // A room the session stopped following while it caught up stays read-only.
    if (!joined.stopped) setStatus("joined");
    joined.presence?.update();
    showRemoteSelections();
  };
  void join();

  return {
    get editable() {
      return status === "joined" && !left;
    },
    selectionChanged: () => {
      if (!left) joined?.presence?.update();
    },
    documentChanged: showRemoteSelections,
    flush: () => {
      if (!left) joined?.flush();
    },
    leave: () => {
      if (left) return;
      left = true;
      offPresence();
      joined?.leave();
      connection.close();
      onRemoteSelections([]);
    },
  };
};
