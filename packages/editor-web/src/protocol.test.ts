import { describe, expect, test } from "bun:test";

import { isEditorMessage, isHostMessage } from "./protocol";
import type { EditorMessage, HostMessage } from "./protocol";

const document = {
  bytes: new Uint8Array([0x50, 0x4b]),
  fileVersion: "1727344000000",
  fileName: "contract.docx",
};

/** One of every host message; `satisfies` keeps the list honest as the union grows. */
const HOST_MESSAGES = {
  load: { type: "load", document, author: "Ada", mode: "editing", locale: "cs" },
  reload: { type: "reload", document },
  serialize: { type: "serialize", requestId: 7 },
  undo: { type: "undo" },
  redo: { type: "redo" },
  setMode: { type: "setMode", mode: "suggesting" },
} as const satisfies { [Type in HostMessage["type"]]: Extract<HostMessage, { type: Type }> };

const EDITOR_MESSAGES = {
  ready: { type: "ready" },
  loaded: { type: "loaded", fileVersion: "1727344000000" },
  loadFailed: { type: "loadFailed", message: "Not a .docx file" },
  edit: { type: "edit" },
  dirty: { type: "dirty", dirty: true },
  modeChanged: { type: "modeChanged", mode: "viewing" },
  serialized: {
    type: "serialized",
    requestId: 7,
    bytes: new Uint8Array([0x50, 0x4b]),
    fileVersion: "1727344000000",
    strategy: { type: "full-repack", reason: "structuralChange" },
  },
  serializeFailed: { type: "serializeFailed", requestId: 7, message: "Failed" },
  error: { type: "error", message: "Font failed to load" },
} as const satisfies { [Type in EditorMessage["type"]]: Extract<EditorMessage, { type: Type }> };

describe("isHostMessage", () => {
  test("accepts every host message", () => {
    for (const message of Object.values(HOST_MESSAGES)) {
      expect(isHostMessage(message)).toBe(true);
    }
  });

  test("rejects what the extension never sends", () => {
    for (const message of [
      null,
      "undo",
      { type: "unknown" },
      { ...HOST_MESSAGES.load, mode: "track" },
      { ...HOST_MESSAGES.load, document: { ...document, bytes: [0x50, 0x4b] } },
      { ...HOST_MESSAGES.reload, document: { bytes: document.bytes } },
      { type: "serialize", requestId: -1 },
      { type: "serialize", requestId: 1.5 },
      { type: "serialize", requestId: "7" },
    ]) {
      expect(isHostMessage(message)).toBe(false);
    }
  });
});

describe("isEditorMessage", () => {
  test("accepts every editor message", () => {
    for (const message of Object.values(EDITOR_MESSAGES)) {
      expect(isEditorMessage(message)).toBe(true);
    }
    expect(
      isEditorMessage({ ...EDITOR_MESSAGES.serialized, strategy: { type: "selective-first" } }),
    ).toBe(true);
  });

  test("rejects what the webview never sends", () => {
    for (const message of [
      undefined,
      { type: "dirty", dirty: "yes" },
      { type: "modeChanged", mode: "edit" },
      { ...EDITOR_MESSAGES.serialized, bytes: new ArrayBuffer(2) },
      { ...EDITOR_MESSAGES.serialized, strategy: { type: "full-repack" } },
      { ...EDITOR_MESSAGES.serialized, strategy: { type: "full-repack", reason: "hunch" } },
      { type: "serializeFailed", requestId: 7 },
    ]) {
      expect(isEditorMessage(message)).toBe(false);
    }
  });
});
