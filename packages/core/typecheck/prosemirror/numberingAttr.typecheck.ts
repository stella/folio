/**
 * Compile-time proof that a stored paragraph snapshot carries the numbering
 * union, not the two `<w:numPr>` slots a build before it wrote.
 *
 * `paragraphRejectOriginalFormatting` rebuilds `_originalFormatting` from a
 * stored snapshot. While the snapshot's `numPr` was the slot pair, the value
 * reached `sameStatedParagraphNumbering`, which switches on `kind`, and an
 * untouched save began materialising numbering as direct formatting. The
 * runtime half — that such a record panics rather than being copied — is
 * pinned in `numberingAttr.test.ts`; this is the half a runtime test cannot
 * reach, because a package's `typecheck` never reads a test file.
 */

import { paragraphRejectOriginalFormatting } from "../../src/prosemirror/commands/propertyChangeScope";
import {
  effectiveParagraphNumbering,
  effectiveParagraphNumberingReference,
  paragraphNumberingSourcesFromBoundary,
} from "../../src/prosemirror/numberingAttr";

const REJECTS_THE_PRE_UNION_SLOT_PAIR = paragraphRejectOriginalFormatting(
  {
    // @ts-expect-error the snapshot carries the minted attr, not the raw slots
    numPr: { numId: 1, ilvl: 0 },
  },
  null,
);

export type ParagraphSnapshotNumberingProof =
  typeof REJECTS_THE_PRE_UNION_SLOT_PAIR;

// Both tiers must be stated even when absent; null belongs only to the attrs boundary.
const COMPLETE_SOURCES = { numPr: undefined, numPrFromStyle: undefined };
effectiveParagraphNumbering(COMPLETE_SOURCES);
effectiveParagraphNumberingReference(COMPLETE_SOURCES);
effectiveParagraphNumbering(
  paragraphNumberingSourcesFromBoundary({
    numPr: null,
    numPrFromStyle: null,
  }),
);
// @ts-expect-error the authored source tier is required
effectiveParagraphNumbering({ numPrFromStyle: undefined });
// @ts-expect-error the style source tier is required
effectiveParagraphNumbering({ numPr: undefined });
// @ts-expect-error a reference consumer must also state both tiers
effectiveParagraphNumberingReference({ numPr: undefined });
// @ts-expect-error null cannot carry absence after the boundary
effectiveParagraphNumbering({ numPr: null, numPrFromStyle: undefined });
// @ts-expect-error null cannot carry style absence after the boundary
effectiveParagraphNumbering({ numPr: undefined, numPrFromStyle: null });
// @ts-expect-error the attrs boundary cannot silently omit the authored tier
paragraphNumberingSourcesFromBoundary({ numPrFromStyle: undefined });
// @ts-expect-error the attrs boundary cannot silently omit the style tier
paragraphNumberingSourcesFromBoundary({ numPr: undefined });
