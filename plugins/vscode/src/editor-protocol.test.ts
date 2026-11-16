import { describe, expect, test } from "bun:test";

import { isEditorMessage, saveStrategyFlag } from "./editor-protocol";

describe("isEditorMessage", () => {
  test("accepts what the editor webview sends", () => {
    expect(isEditorMessage({ type: "ready" })).toBe(true);
    expect(isEditorMessage({ type: "edit" })).toBe(true);
    expect(isEditorMessage({ type: "modeChanged", mode: "suggesting" })).toBe(
      true,
    );
    expect(
      isEditorMessage({
        type: "serialized",
        requestId: 3,
        bytes: new Uint8Array([1]),
        fileVersion: "load-1",
        strategy: {
          type: "full-repack",
          decline: {
            scope: "part",
            part: "word/document.xml",
            reason: "structuralRangeSplit",
          },
        },
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
        strategy: { type: "selective" },
      }),
    ).toBe(false);
    expect(
      isEditorMessage({
        type: "serialized",
        requestId: 3,
        bytes: new Uint8Array([1]),
        fileVersion: "load-1",
        strategy: {
          type: "full-repack",
          decline: {
            scope: "part",
            part: "word/document.xml",
            reason: "because",
          },
        },
      }),
    ).toBe(false);
    expect(isEditorMessage({ type: "modeChanged", mode: "drafting" })).toBe(
      false,
    );
  });
});

describe("saveStrategyFlag", () => {
  test("maps the editor's strategy to folio save's", () => {
    expect(saveStrategyFlag({ type: "selective" })).toBe("selective");
    expect(
      saveStrategyFlag({
        type: "full-repack",
        decline: {
          scope: "part",
          part: "word/document.xml",
          reason: "structuralRangeSplit",
        },
      }),
    ).toBe("full-repack");
  });
});
