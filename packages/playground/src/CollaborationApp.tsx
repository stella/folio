import { useEffect, useRef, useState, useSyncExternalStore } from "react";

import { IntlProvider } from "use-intl";

import { DocxEditor, type DocxEditorRef } from "@stll/folio-react";
import { getFolioMessages } from "@stll/folio-react/messages";

import { buildCanonicalBridge } from "../../../tests/parity/canonicalBridge";
import { buildCanonicalRoomBridge } from "../../../tests/parity/canonicalRoomBridge";
import {
  type CanonicalRoomDemo,
  type CanonicalRoomDemoState,
  createCanonicalRoomDemo,
  readCanonicalRoomSeed,
} from "../../../tests/parity/canonicalRoomDemo";
import { AvatarStack } from "./collaboration/AvatarStack";
import {
  getOrCreateRoomFromUrl,
  loadOrCreateUser,
} from "./collaboration/identity";

const DEFAULT_FILE = "sample.docx";

type Opened =
  | { type: "loading" }
  | { type: "error"; message: string }
  | { type: "ready"; bytes: ArrayBuffer; demo: CanonicalRoomDemo };

/**
 * Canonical room demo: every tab of this origin that opens the same room
 * (the URL hash) and the same DOCX (`?file=`, `sample.docx` by default)
 * edits one document. One tab runs the room's authority, the journal lives
 * in IndexedDB, and comments travel in the journal with the text.
 */
export function CollaborationApp() {
  const [user] = useState(loadOrCreateUser);
  const [room] = useState(getOrCreateRoomFromUrl);
  const [opened, setOpened] = useState<Opened>({ type: "loading" });

  useEffect(() => {
    if (!window.location.hash) {
      window.history.replaceState(
        null,
        "",
        `${window.location.search}#${room}`,
      );
    }
  }, [room]);

  useEffect(() => {
    let cancelled = false;
    const file =
      new URLSearchParams(window.location.search).get("file") ?? DEFAULT_FILE;
    const open = async () => {
      const bytes = await readCanonicalRoomSeed(file);
      if (bytes === undefined) {
        setOpened({ type: "error", message: `Fixture not found: ${file}` });
        return;
      }
      const demo = await createCanonicalRoomDemo({ room, bytes, user });
      if (!cancelled) setOpened({ type: "ready", bytes, demo });
    };
    open().catch((error: unknown) => {
      if (!cancelled)
        setOpened({
          type: "error",
          message: error instanceof Error ? error.message : String(error),
        });
    });
    return () => {
      cancelled = true;
    };
  }, [room, user]);

  return (
    <IntlProvider locale="en" messages={getFolioMessages("en")}>
      <div className="pg-shell pg-shell--collab">
        {opened.type === "ready" ? (
          <RoomEditor bytes={opened.bytes} demo={opened.demo} room={room} />
        ) : (
          <>
            <RoomHeader room={room} state={undefined} />
            <main className="pg-editor-area">
              <div className="pg-collab-loading">
                {opened.type === "error" ? opened.message : "Opening..."}
              </div>
            </main>
          </>
        )}
      </div>
    </IntlProvider>
  );
}

const RoomEditor = ({
  bytes,
  demo,
  room,
}: {
  bytes: ArrayBuffer;
  demo: CanonicalRoomDemo;
  room: string;
}) => {
  const editorRef = useRef<DocxEditorRef>(null);
  const state = useSyncExternalStore(demo.subscribe, demo.getState);

  useEffect(() => {
    globalThis.__folioCanonical = buildCanonicalBridge(() => editorRef.current);
    globalThis.__folioCanonicalReady = true;
    globalThis.__folioCanonicalRoom = buildCanonicalRoomBridge(
      () => editorRef.current?.getEditor() ?? null,
      demo.status,
    );
    return () => {
      globalThis.__folioCanonicalReady = false;
      globalThis.__folioCanonical = undefined;
      globalThis.__folioCanonicalRoom = undefined;
    };
  }, [demo]);

  return (
    <>
      <RoomHeader room={room} state={state} />
      <main className="pg-editor-area">
        <DocxEditor
          ref={editorRef}
          accessibleName="Document editor"
          documentBuffer={bytes}
          experimentalSession="canonical"
          author={demo.collaboration.user?.name ?? "Folio User"}
          collaboration={demo.collaboration}
          showToolbar={true}
          showRuler={true}
          initialZoom={1}
        />
      </main>
    </>
  );
};

const RoomHeader = ({
  room,
  state,
}: {
  room: string;
  state: CanonicalRoomDemoState | undefined;
}) => {
  const connected = state?.connected === true;
  return (
    <header className="pg-collab-header">
      <div className="pg-collab-header__left">
        <strong>folio collaboration demo</strong>
        <span
          className="pg-collab-status"
          data-status={connected ? "connected" : "connecting"}
        >
          {connected ? `Live · ${room}` : "Connecting..."}
        </span>
        {state?.notice === undefined ? null : (
          <span className="pg-collab-status">{state.notice}</span>
        )}
      </div>
      <div className="pg-collab-header__right">
        <AvatarStack users={state?.users ?? []} />
        <ShareLinkButton />
      </div>
    </header>
  );
};

/** Copy the room's URL, which a second tab opens to join it. */
const ShareLinkButton = () => {
  const [copied, setCopied] = useState(false);
  useEffect(() => {
    if (!copied) return undefined;
    const reset = setTimeout(() => setCopied(false), 1500);
    return () => clearTimeout(reset);
  }, [copied]);
  const copy = () => {
    navigator.clipboard.writeText(window.location.href).then(
      () => setCopied(true),
      () => setCopied(false),
    );
  };
  return (
    <button type="button" className="pg-button" onClick={copy}>
      {copied ? "Link copied!" : "Share link"}
    </button>
  );
};
