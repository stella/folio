import type { TextPosition } from "@stll/docx-core/ops";

/** Input meaning belongs to the producer, not the compiled operation shape. */
export type CanonicalTyping = {
  at: TextPosition;
  after: TextPosition;
  text: string;
};

const typingByBatch = new WeakMap<object, CanonicalTyping>();

export const recordCanonicalTyping = (
  batch: object,
  typing: CanonicalTyping,
): void => {
  typingByBatch.set(batch, typing);
};

export const canonicalTypingOf = (batch: object): CanonicalTyping | undefined =>
  typingByBatch.get(batch);
