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

const REJECTS_THE_PRE_UNION_SLOT_PAIR = paragraphRejectOriginalFormatting(
  {
    // @ts-expect-error the snapshot carries the minted attr, not the raw slots
    numPr: { numId: 1, ilvl: 0 },
  },
  null,
);

export type ParagraphSnapshotNumberingProof = typeof REJECTS_THE_PRE_UNION_SLOT_PAIR;
