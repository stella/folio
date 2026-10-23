/**
 * Compile-time proof that a numbering cancellation declares no level.
 *
 * `numId` 0 cancels numbering whatever `ilvl` the package stated beside it.
 * Three committed tests used to disagree about what that state held, which
 * was possible only while the two `<w:numPr>` slots travelled together; the
 * `none` arm has no slot for a level, so the disputed state is
 * unconstructible. `numberingRoundTrip.property.test.ts` states the runtime
 * half; this is the half a package's `typecheck` can see, because it never
 * reads a test file.
 */

import {
  NO_NUMBERING_NUM_ID,
  type ParagraphNumberingOverride,
  paragraphNumberingReference,
} from "../../src/model/paragraphNumbering";

const CANCELLATION_WITH_A_LEVEL: ParagraphNumberingOverride = {
  kind: "none",
  // @ts-expect-error the cancellation arm declares no level
  ilvl: 0,
};

const CANCELLATION: ParagraphNumberingOverride = { kind: "none" };

// @ts-expect-error a reference is minted only after the reserved id is excluded
const RESERVED_REFERENCE: ParagraphNumberingOverride = {
  kind: "reference",
  numId: NO_NUMBERING_NUM_ID,
};

const RESERVED_CONSTRUCTOR = paragraphNumberingReference({
  // @ts-expect-error the constructor also rejects a literal reserved id
  numId: NO_NUMBERING_NUM_ID,
});

export type ParagraphNumberingProof = [
  typeof CANCELLATION_WITH_A_LEVEL,
  typeof CANCELLATION,
  typeof RESERVED_REFERENCE,
  typeof RESERVED_CONSTRUCTOR,
];
