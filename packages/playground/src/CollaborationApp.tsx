import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { IntlProvider } from "use-intl";

import {
  DocxEditor,
  createEmptyDocument,
  createStellaStyleDocumentPreset,
} from "@stll/folio-react";
import { getFolioMessages } from "@stll/folio-react/messages";

import { AvatarStack } from "./collaboration/AvatarStack";
import { getOrCreateRoomFromUrl, loadOrCreateUser } from "./collaboration/identity";
import { useCollaboration } from "./collaboration/useCollaboration";

const collaborationStatusLabel = (
  status: "connecting" | "connected" | "disconnected",
  room: string,
): string => {
  if (status === "connected") {
    return `Live · ${room}`;
  }
  if (status === "connecting") {
    return "Connecting…";
  }
  return "Offline";
};

/**
 * Live Yjs + WebRTC collaboration demo.
 *
 * Open `/?collaboration=1` in two browser windows (or tabs) and share the
 * URL hash (`#room-…`) to sync document body and comment threads.
 */
export function CollaborationApp() {
  const [user] = useState(loadOrCreateUser);
  const [room] = useState(getOrCreateRoomFromUrl);
  const [shareCopied, setShareCopied] = useState(false);
  const shareCopiedTimeoutRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  useEffect(() => {
    if (!window.location.hash) {
      window.history.replaceState(null, "", `#${room}`);
    }
  }, [room]);

  useEffect(
    () => () => {
      if (shareCopiedTimeoutRef.current) {
        clearTimeout(shareCopiedTimeoutRef.current);
      }
    },
    [],
  );

  const { collaboration, users, status, comments, setComments } = useCollaboration(room, user);
  const seedDocument = useMemo(
    () => createEmptyDocument({ preset: createStellaStyleDocumentPreset() }),
    [],
  );

  const handleCopyShareLink = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(window.location.href);
      setShareCopied(true);
      if (shareCopiedTimeoutRef.current) {
        clearTimeout(shareCopiedTimeoutRef.current);
      }
      shareCopiedTimeoutRef.current = setTimeout(() => setShareCopied(false), 1500);
    } catch {
      setShareCopied(false);
    }
  }, []);

  return (
    <IntlProvider locale="en" messages={getFolioMessages("en")}>
      <div className="pg-shell pg-shell--collab">
        <header className="pg-collab-header">
          <div className="pg-collab-header__left">
            <strong>folio collaboration demo</strong>
            <span className="pg-collab-status" data-status={status}>
              {collaborationStatusLabel(status, room)}
            </span>
          </div>
          <div className="pg-collab-header__right">
            <AvatarStack users={users} />
            <button type="button" className="pg-button" onClick={() => void handleCopyShareLink()}>
              {shareCopied ? "Link copied!" : "Share link"}
            </button>
          </div>
        </header>
        <main className="pg-editor-area">
          {collaboration ? (
            <DocxEditor
              document={seedDocument}
              author={user.name}
              collaboration={collaboration}
              comments={comments}
              onCommentsChange={setComments}
              showToolbar={true}
              showRuler={true}
              initialZoom={1}
            />
          ) : (
            <div className="pg-collab-loading">Connecting...</div>
          )}
        </main>
      </div>
    </IntlProvider>
  );
}
