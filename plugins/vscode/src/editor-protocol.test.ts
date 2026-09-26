import { describe, expect, test } from "bun:test";

import { isEditorMessage, saveStrategyFlag } from "./editor-protocol";

describe("isEditorMessage", () => {
  test("accepts what the editor webview sends", () => {
    expect(isEditorMessage({ type: "ready" })).toBe(true);
    expect(isEditorMessage({ type: "edit" })).toBe(true);
    expect(isEditorMessage({ type: "modeChanged", mode: "suggesting" })).toBe(true);
    expect(
      isEditorMessage({
        type: "serialized",
        requestId: 3,
        bytes: new Uint8Array([1]),
        fileVersion: "load-1",
        strategy: { type: "full-repack", reason: "structuralChange" },
      }),
    ).toBe(true);
  });

  test("refuses a serialization without bytes or with an unknown strategy", () => {
    expect(
      isEditorMessage({
        type: "serialized",
        requestId: 3,
        bytes: [1],
        fileVersion: "load-1",
        strategy: { type: "selective-first" },
      }),
    ).toBe(false);
    expect(
      isEditorMessage({
        type: "serialized",
        requestId: 3,
        bytes: new Uint8Array([1]),
        fileVersion: "load-1",
        strategy: { type: "full-repack", reason: "because" },
      }),
    ).toBe(false);
    expect(isEditorMessage({ type: "modeChanged", mode: "drafting" })).toBe(false);
  });
});

describe("saveStrategyFlag", () => {
  test("maps the editor's strategy to folio save's", () => {
    expect(saveStrategyFlag({ type: "selective-first" })).toBe("selective");
    expect(saveStrategyFlag({ type: "full-repack", reason: "untrackedChange" })).toBe(
      "full-repack",
    );
  });
});
