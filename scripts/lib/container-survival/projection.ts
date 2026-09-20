/**
 * The projection the survival law's L3 leg measures, with reuse declined.
 *
 * L3 asks whether the ProseMirror projection can carry a pair, and it answers
 * by comparing the fixture's markup against the part a
 * `toProseDoc` → `fromProseDoc` → save round trip writes. That question is only
 * answered while every record comes back out of ProseMirror. Once
 * `fromProseDoc` merges untouched records back from the base document by
 * reference, a pair the projection never carried returns through the base
 * object, the law reports it as surviving, and the contract ratchets a
 * projection survival in that the projection never performed — for every pair
 * the census currently records as `editorProjection`.
 *
 * So the law projects through here rather than calling `fromProseDoc` itself,
 * and `scripts/container-survival-projection.test.ts` binds the two sides: the
 * law may not reach the conversion past this module, and this module must
 * decline reuse as soon as the conversion offers the choice.
 *
 * `fromProseDoc` takes no reuse option on this branch — rebuilding every record
 * is its only behaviour, so the two-argument call already declines. The moment
 * it accepts one, this call passes `{ reuse: "none" }`.
 */

import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import type { Document } from "@stll/folio-core/types/document";

/** The ProseMirror document the conversion reads, named by the conversion. */
type ProjectedNode = Parameters<typeof fromProseDoc>[0];

/** Convert a ProseMirror document back to the model, rebuilding every record. */
export const projectWithoutReuse = (pmDoc: ProjectedNode, base: Document): Document =>
  fromProseDoc(pmDoc, base);
