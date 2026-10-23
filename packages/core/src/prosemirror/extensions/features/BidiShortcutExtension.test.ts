import { describe, expect, test } from "bun:test";

import { createBidiShortcutController } from "./bidiShortcutController";
import type { BidiKeyEvent } from "./bidiShortcutController";

const key = (over: Partial<BidiKeyEvent>): BidiKeyEvent => ({
  key: "Shift",
  code: "ShiftLeft",
  metaKey: false,
  ctrlKey: false,
  ...over,
});

describe("createBidiShortcutController", () => {
  test("Cmd/Ctrl + Left Shift applies LTR on Shift keyup", () => {
    const c = createBidiShortcutController();
    c.handleKeyDown(key({ code: "ShiftLeft", metaKey: true }));
    expect(c.handleKeyUp(key({ code: "ShiftLeft", metaKey: true }))).toBe(
      "ltr",
    );
  });

  test("Cmd/Ctrl + Right Shift applies RTL on Shift keyup", () => {
    const c = createBidiShortcutController();
    c.handleKeyDown(key({ code: "ShiftRight", ctrlKey: true }));
    expect(c.handleKeyUp(key({ code: "ShiftRight", ctrlKey: true }))).toBe(
      "rtl",
    );
  });

  test("a following key cancels the direction change (Ctrl/Cmd+Shift+Z passes through)", () => {
    // This is the redo chord: arming the direction change on the Shift keydown
    // and applying it immediately would wipe the redo stack before the "Z".
    const c = createBidiShortcutController();
    c.handleKeyDown(key({ code: "ShiftLeft", metaKey: true }));
    c.handleKeyDown(key({ key: "z", code: "KeyZ", metaKey: true })); // the "Z" of Cmd+Shift+Z
    expect(c.handleKeyUp(key({ code: "ShiftLeft", metaKey: true }))).toBeNull();
  });

  test("Shift without a Ctrl/Cmd modifier never arms a direction change", () => {
    const c = createBidiShortcutController();
    c.handleKeyDown(key({ code: "ShiftLeft", metaKey: false, ctrlKey: false }));
    expect(c.handleKeyUp(key({ code: "ShiftLeft" }))).toBeNull();
  });

  test("only the Shift keyup applies the change; other keyups are ignored", () => {
    const c = createBidiShortcutController();
    c.handleKeyDown(key({ code: "ShiftLeft", metaKey: true }));
    // A non-Shift keyup (e.g. releasing Cmd) neither fires nor disarms it.
    expect(
      c.handleKeyUp(key({ key: "Meta", code: "MetaLeft", metaKey: false })),
    ).toBeNull();
    expect(c.handleKeyUp(key({ code: "ShiftLeft" }))).toBe("ltr");
  });
});
