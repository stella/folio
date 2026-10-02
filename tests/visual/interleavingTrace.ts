import fc from "fast-check";

export type InterleavingAction =
  | { kind: "suggest"; text: string }
  | { kind: "typing"; text: string }
  | { kind: "accept" | "reject"; target: number }
  | { kind: "undo" | "redo" };

const textArbitrary = fc
  .array(fc.constantFrom("alpha", "buyer", "café", "東京", "👩🏽‍⚖️"), {
    minLength: 1,
    maxLength: 3,
  })
  .map((words) => words.join(" "));
const suggestArbitrary = textArbitrary.map((text) => ({ kind: "suggest", text }) as const);
const typingArbitrary = textArbitrary.map((text) => ({ kind: "typing", text }) as const);
const remainingArbitrary = fc.oneof(
  suggestArbitrary,
  typingArbitrary,
  fc.record({ kind: fc.constantFrom("accept", "reject"), target: fc.nat(8) }),
  fc.record({ kind: fc.constantFrom("undo", "redo") }),
);

/** Shrinking preserves the AI/human overlap that makes this a distinct target. */
export const interleavingTraceArbitrary = fc
  .record({
    shape: fc.constantFrom("plain-markdown", "single-decimal-list", "rtl-cjk"),
    suggest: suggestArbitrary,
    typing: typingArbitrary,
    remaining: fc.array(remainingArbitrary, { minLength: 1, maxLength: 10 }),
  })
  .map(({ shape, suggest, typing, remaining }) => ({
    shape,
    actions: [suggest, typing, ...remaining],
  }));
