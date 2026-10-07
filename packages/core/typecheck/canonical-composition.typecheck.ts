import { createCanonicalInputBoundary } from "../src/controller/canonicalInput";

type Replacement = Parameters<Parameters<typeof createCanonicalInputBoundary>[0]["replace"]>[0];

const correction = {
  from: 1,
  to: 2,
  text: "契約",
  semantic: "composition",
  compositionPhase: "correction",
} as const satisfies Replacement;
const typingCorrection = {
  from: 1,
  to: 2,
  text: "契約",
  semantic: "typing",
  compositionPhase: "correction",
} as const;
// @ts-expect-error A composition correction cannot be a typing intent.
const rejected = typingCorrection satisfies Replacement;
void rejected;
void correction;
void typingCorrection;
