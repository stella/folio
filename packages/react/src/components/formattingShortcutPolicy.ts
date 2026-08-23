type RightAlignmentShortcutEvent = Pick<KeyboardEvent, "ctrlKey" | "metaKey">;

/** Ctrl+R aligns right; any chord containing Meta remains browser-owned. */
export const shouldClaimRightAlignmentShortcut = ({
  ctrlKey,
  metaKey,
}: RightAlignmentShortcutEvent): boolean => ctrlKey && !metaKey;
