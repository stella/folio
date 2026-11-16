import { describe, expect, test } from "bun:test";

import type { SelectiveSaveDecline } from "@stll/folio-core/types/selectiveSaveDecline";

import { isEditorMessage, isHostMessage } from "./protocol";
import type { EditorMessage, HostMessage } from "./protocol";

const document = {
  bytes: new Uint8Array([0x50, 0x4b]),
  fileVersion: "1727344000000",
  fileName: "contract.docx",
};

/** One of every host message; `satisfies` keeps the list honest as the union grows. */
const HOST_MESSAGES = {
  load: {
    type: "load",
    document,
    author: "Ada",
    mode: "editing",
    locale: "cs",
  },
  reload: { type: "reload", document },
  serialize: { type: "serialize", requestId: 7 },
  undo: { type: "undo" },
  redo: { type: "redo" },
  setMode: { type: "setMode", mode: "suggesting" },
} as const satisfies {
  [Type in HostMessage["type"]]: Extract<HostMessage, { type: Type }>;
};

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
    strategy: {
      type: "full-repack",
      decline: {
        scope: "part",
        part: "word/document.xml",
        reason: "structuralRangeSplit",
      },
    },
  },
  serializeFailed: { type: "serializeFailed", requestId: 7, message: "Failed" },
  error: { type: "error", message: "Font failed to load" },
} as const satisfies {
  [Type in EditorMessage["type"]]: Extract<EditorMessage, { type: Type }>;
};

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
      isEditorMessage({
        ...EDITOR_MESSAGES.serialized,
        strategy: { type: "selective" },
      }),
    ).toBe(true);
  });

  test("a full write without a selective attempt has no invented decline", () => {
    expect(
      isEditorMessage({
        ...EDITOR_MESSAGES.serialized,
        strategy: { type: "full-repack" },
      }),
    ).toBe(true);
  });

  test("rejects what the webview never sends", () => {
    for (const message of [
      undefined,
      { type: "dirty", dirty: "yes" },
      { type: "modeChanged", mode: "edit" },
      { ...EDITOR_MESSAGES.serialized, bytes: new ArrayBuffer(2) },
      { ...EDITOR_MESSAGES.serialized, strategy: { type: "selective-first" } },
      {
        ...EDITOR_MESSAGES.serialized,
        strategy: {
          type: "full-repack",
          decline: {
            scope: "part",
            part: "word/document.xml",
            reason: "hunch",
          },
        },
      },
      {
        ...EDITOR_MESSAGES.serialized,
        strategy: { type: "full-repack", decline: null },
      },
      { type: "serializeFailed", requestId: 7 },
    ]) {
      expect(isEditorMessage(message)).toBe(false);
    }
  });
});

test("selective refusal scope validates only its own payload", () => {
  for (const decline of [
    { scope: "package", reason: "tooLarge" },
    {
      scope: "part",
      part: "word/document.xml",
      reason: "structuralRangeSplit",
    },
    {
      scope: "story",
      story: { kind: "header", rId: "rId44" },
      reason: "unmaterializedStoryResource",
    },
  ] as const satisfies readonly SelectiveSaveDecline[])
    expect(
      isEditorMessage({
        ...EDITOR_MESSAGES.serialized,
        strategy: { type: "full-repack", decline },
      }),
    ).toBe(true);
  for (const decline of [
    { scope: "package", part: "word/document.xml", reason: "tooLarge" },
    { scope: "part", reason: "paragraphPatch" },
    { scope: "part", part: "", reason: "paragraphPatch" },
    { scope: "part", part: 44, reason: "paragraphPatch" },
    {
      scope: "story",
      story: { kind: "main", rId: "rId44" },
      reason: "unmaterializedStoryResource",
    },
    {
      scope: "story",
      story: { kind: "footer" },
      reason: "unmaterializedStoryResource",
    },
    {
      scope: "story",
      story: { kind: "footer", rId: "rId44" },
      part: "word/footer1.xml",
      reason: "unmaterializedStoryResource",
    },
    { scope: "unknown", reason: "tooLarge" },
  ])
    expect(
      isEditorMessage({
        ...EDITOR_MESSAGES.serialized,
        strategy: { type: "full-repack", decline },
      }),
    ).toBe(false);
});
