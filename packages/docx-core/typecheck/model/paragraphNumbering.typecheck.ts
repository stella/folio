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

import type { ParagraphNumberingOverride } from "../../src/model/paragraphNumbering";

// @ts-expect-error the cancellation arm declares no level
const CANCELLATION_WITH_A_LEVEL: ParagraphNumberingOverride = { kind: "none", ilvl: 0 };

const CANCELLATION: ParagraphNumberingOverride = { kind: "none" };

export type ParagraphNumberingProof = [typeof CANCELLATION_WITH_A_LEVEL, typeof CANCELLATION];
