import { describe, expect, test } from "bun:test";

import { shouldClaimRightAlignmentShortcut } from "./formattingShortcutPolicy";

describe("right-alignment shortcut ownership", () => {
  test("leaves Cmd+R and Cmd+Ctrl+R to the browser", () => {
    expect(shouldClaimRightAlignmentShortcut({ ctrlKey: false, metaKey: true })).toBe(false);
    expect(shouldClaimRightAlignmentShortcut({ ctrlKey: true, metaKey: true })).toBe(false);
  });

  test("claims Ctrl+R only", () => {
    expect(shouldClaimRightAlignmentShortcut({ ctrlKey: true, metaKey: false })).toBe(true);
    expect(shouldClaimRightAlignmentShortcut({ ctrlKey: false, metaKey: false })).toBe(false);
  });
});
