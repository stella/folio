/**
 * How the survival law forces a part serializer folio's repack does not run.
 *
 * L2 forces the *element* serializers by stripping the verbatim captures the
 * model holds, because a replay hands the captured bytes back whatever the
 * serializers would have written. For a declaration part the replay is not a
 * capture inside the model, it is the part itself: `rezip.ts` carries the
 * original entry across unless something asked for it to be rewritten. So the
 * law has to force the **part** serializer, and this table is where it asks
 * whether folio has one.
 *
 * A rebuilder takes the parsed package — with its captures already stripped,
 * exactly as the document part's forcing hands it over — and returns the part
 * as folio would write it from the model alone. `undefined` means the model
 * holds no record for the part, which is itself a loss and is measured as one.
 *
 * `null` is a *stated* absence rather than a gap: folio does not rebuild the
 * part yet, so the census keeps every pair in it `unrepresentable` with the
 * reason {@link PART_REBUILD_ABSENCES} gives. The two tables are bound to each
 * other — the absence reasons are keyed by the roots that are `null` here — so
 * a root that gains a rebuilder cannot leave a stale reason behind and a root
 * that stays absent cannot leave the reason unwritten.
 *
 * Nothing here changes what a save writes. The repack keeps copying these
 * parts; a pair is `modelled` when the model *can* rebuild it, not when folio
 * does.
 */

import type { Document } from "@stll/folio-core/types/document";

import type { RebuiltPartRoot } from "./schemaSpace";

/**
 * Rebuild one part from the model.
 *
 * @returns the part as its serializer writes it, or `undefined` when the model
 *   holds no record for it at all.
 */
export type PartRebuilder = (document: Document) => string | undefined;

/**
 * Adoption order, by what a rebuild loses today: `fontTable` → `numbering` →
 * `styles` → notes → `settings`. Each part lands its sink in the same change as
 * its measurement, because measuring a part before it can keep what it loses
 * turns silent unknowns into recorded losses and nothing else.
 */
export const PART_REBUILDERS = {
  // The five parts folio already rebuilds on the save path reach L2 through
  // the repack, so the law never asks this table about them.
  document: null,
  comments: null,
  endnotes: null,
  footnotes: null,
  hdr: null,
  ftr: null,
  fonts: null,
  numbering: null,
  settings: null,
  styles: null,
  webSettings: null,
} as const satisfies Record<RebuiltPartRoot, PartRebuilder | null>;

/**
 * The roots {@link PART_REBUILDERS} states no rebuilder for.
 *
 * Derived from the table rather than listed again, so the reason set below is
 * total over the absences by construction: adding a rebuilder drops its root
 * from this union and a reason left behind stops compiling.
 */
type AbsentPartRoot = {
  [Root in RebuiltPartRoot]: (typeof PART_REBUILDERS)[Root] extends null ? Root : never;
}[RebuiltPartRoot];

/**
 * Why the census cannot measure a part, in the words of what is missing.
 *
 * The reason a reader needs is what folio lacks, not that a file was copied:
 * `webSettings` has no serializer at all, and the other four have one that no
 * save path reaches for a document somebody else authored.
 */
export const PART_REBUILD_ABSENCES = {
  document: "a repack rebuilds it, so the law never asks",
  comments: "a repack rebuilds it, so the law never asks",
  hdr: "a repack rebuilds it, so the law never asks",
  ftr: "a repack rebuilds it, so the law never asks",
  endnotes: "its serializer writes one note at a time and never the whole part",
  footnotes: "its serializer writes one note at a time and never the whole part",
  fonts: "no leg rebuilds it from the model",
  numbering: "folio splices it by id rather than rebuilding it from the model",
  settings: "folio copies it and rebuilds it from the model on no save path",
  styles: "folio splices style definitions into it rather than rebuilding it",
  webSettings: "folio has no serializer for it",
} as const satisfies Record<AbsentPartRoot, string>;

/**
 * What the law does with a part the repack copied through.
 *
 * A rebuilder means the part is measured; an absence means every pair in it
 * stays `unrepresentable` with the reason stated.
 */
export type PartRebuild =
  | { kind: "rebuilt"; rebuild: PartRebuilder }
  | { kind: "absent"; reason: string };

const statesNoRebuilder = (root: RebuiltPartRoot): root is AbsentPartRoot =>
  PART_REBUILDERS[root] === null;

export const partRebuildFor = (root: RebuiltPartRoot): PartRebuild =>
  statesNoRebuilder(root)
    ? { kind: "absent", reason: PART_REBUILD_ABSENCES[root] }
    : { kind: "rebuilt", rebuild: PART_REBUILDERS[root] };
