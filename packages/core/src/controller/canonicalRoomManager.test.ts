import { panic } from "better-result";
import { expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import type { CanonicalHistoryDrop } from "../canonical/canonicalSession";
import { createEmptyDocument } from "../utils/createDocument";
import type { CanonicalRoomCollaboration } from "./canonicalRoomEditor";
import {
  type CanonicalMemoryRoom,
  createCanonicalMemoryRoom,
} from "./canonicalRoomTransport";
import { createHiddenEditorManager } from "./hiddenEditorManager";

/*
 * Two hidden editors joined to one canonical room in memory: the binding the
 * adapters drive, with real views over a DOM.
 */

const SNAPSHOT = "canonical-room-manager-test";

/** Let every promise continuation and hop run. */
const settle = async () => {
  for (let round = 0; round < 20; round += 1)
    await new Promise<void>((resolve) => setImmediate(resolve));
};

const openEditors = () => {
  let room: CanonicalMemoryRoom | undefined;
  const open = (actor: string) => {
    const host = document.createElement("div");
    document.body.append(host);
    const source = createEmptyDocument({ initialText: "Start" });
    const paragraph = source.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") return panic("Expected a paragraph.");
    paragraph.paraId = "12345678";
    const drops: CanonicalHistoryDrop[] = [];
    const collaboration: CanonicalRoomCollaboration = {
      snapshot: SNAPSHOT,
      actor,
      connect: (base) => {
        room ??= createCanonicalMemoryRoom({
          document: base.document,
          snapshot: base.snapshot,
        });
        const connection = room.connect();
        return Object.assign(connection, { close: connection.disconnect });
      },
    };
    const manager = createHiddenEditorManager({
      getHost: () => host,
      getDocument: () => source,
      getStyles: () => null,
      getExtensionManager: () => undefined,
      getExternalPlugins: () => [],
      getCollaboration: () => undefined,
      getCollaborationModules: () => null,
      getCanonicalRoom: () => collaboration,
      getPrecomputedInitialState: () => null,
      getReadOnly: () => false,
      getExperimentalSession: () => "canonical",
      onCanonicalHistoryDrop: (drop) => drops.push(drop),
      getDocumentIdentity: () => "0",
      getDocumentContext: () => source,
      onTransaction: () => {},
      onCanonicalCompositionEnd: () => {},
      onSelectionChange: () => {},
      onKeyDown: () => false,
      onReadOnlyEditAttempt: () => {},
      onEditorViewReady: () => {},
      onEditorViewDestroy: () => {},
      onRemoteSelectionsChange: () => {},
    });
    manager.ensureView();
    const view = manager.getView() ?? panic("The editor has no view.");
    const type = (text: string) =>
      view.dom.dispatchEvent(
        new InputEvent("beforeinput", {
          inputType: "insertText",
          data: text,
          cancelable: true,
          bubbles: true,
        }),
      );
    const close = () => {
      manager.destroyView();
      host.remove();
    };
    return { manager, view, type, drops, close };
  };
  return { open };
};

test("a peer's edit that removes a local step reaches the host as a history drop", async () => {
  const releaseDom = acquireHarnessDom();
  const editors = openEditors();
  const first = editors.open("first");
  const second = editors.open("second");
  try {
    // Neither editor takes local edits before the room granted its lease.
    expect(first.view.editable).toBe(false);
    await settle();
    expect(first.view.editable).toBe(true);
    expect(second.view.editable).toBe(true);

    first.manager.api.setSelection(6);
    first.type("!");
    await settle();
    expect(second.view.state.doc.textContent).toBe("Start!");
    expect(first.manager.api.canUndo()).toBe(true);
    expect(first.drops).toEqual([]);

    // The second editor replaces the paragraph's text, the first editor's
    // keystroke with it: its undo step no longer applies.
    second.manager.api.setSelection(1, 7);
    second.type("x");
    await settle();
    expect(first.view.state.doc.textContent).toBe("x");
    expect(second.view.state.doc.textContent).toBe("x");
    expect(first.drops).toEqual([{ undo: 1, redo: 0 }]);
    expect(first.manager.api.canUndo()).toBe(false);
    expect(second.drops).toEqual([]);
  } finally {
    first.close();
    second.close();
    releaseDom();
  }
});

test("a closed editor leaves the room and its view takes no peer's edit", async () => {
  const releaseDom = acquireHarnessDom();
  const editors = openEditors();
  const first = editors.open("first");
  const second = editors.open("second");
  try {
    await settle();
    first.close();
    second.manager.api.setSelection(6);
    second.type("?");
    await settle();
    expect(second.view.state.doc.textContent).toBe("Start?");
    expect(first.manager.getView()).toBeNull();
  } finally {
    second.close();
    releaseDom();
  }
});
