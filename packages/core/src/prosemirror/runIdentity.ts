/**
 * The one place a `runIdentity` mark's attrs are built.
 *
 * `fromProseDoc` groups inline leaves into runs by `JSON.stringify` over their
 * mark attrs, so two leaves from the same authored `w:r` have to spell that
 * run the same way or one run is written out as two. Object key order is
 * insertion order in JavaScript, so the factory fixes it — `id`,
 * `preservedAttributes`, `preserved`, and each remainder entry as `namespace`,
 * `name`, `value` — and every construction site goes through the factory,
 * including the one that reads an id back out of the DOM.
 *
 * Absent stays absent: a field spelled as `undefined` would key differently
 * from one left out, and Yjs persists the attrs object as JSON.
 */

import type { PreservedAttribute, PreservedMarkup } from "../types/document";
import type { RunIdentityMarkAttrs } from "./schema/marks";

/** The DOM attribute that carries a run's identity through copy and paste. */
export const RUN_IDENTITY_ATTRIBUTE = "data-docx-run-identity";

/** The schema name of the mark these attrs belong to. */
export const RUN_IDENTITY_MARK_NAME = "runIdentity";

/** One remainder entry with its keys in canonical order. */
const preservedAttribute = ({ namespace, name, value }: PreservedAttribute): PreservedAttribute =>
  namespace === undefined ? { name, value } : { namespace, name, value };

/** The property-set sink with its keys in canonical order. */
const preservedMarkup = ({ children }: PreservedMarkup): PreservedMarkup => ({
  children: (children ?? []).map(({ index, xml }) => ({ index, xml })),
});

/**
 * What a run carries beside its id: absent fields left absent, empty
 * collections normalised to absent so a run with nothing to carry keys the
 * same however it was built.
 */
export type RunIdentityPayload = {
  preservedAttributes?: readonly PreservedAttribute[] | undefined;
  preserved?: PreservedMarkup | undefined;
};

/** A run identity's attrs in canonical form. */
export const runIdentityAttrs = (
  id: number,
  { preservedAttributes, preserved }: RunIdentityPayload = {},
): RunIdentityMarkAttrs => {
  const attrs: RunIdentityMarkAttrs = { id };
  if (preservedAttributes && preservedAttributes.length > 0) {
    attrs.preservedAttributes = preservedAttributes.map(preservedAttribute);
  }
  if (preserved && (preserved.children?.length ?? 0) > 0) {
    attrs.preserved = preservedMarkup(preserved);
  }
  return attrs;
};

/**
 * Whether a run holds markup the mark exists to carry.
 *
 * The mint condition, minus the page break: a run with neither remainder nor
 * sink and no break to rejoin its leaves across needs no identity, and 57.5%
 * of corpus files hold none.
 */
export const hasRunIdentityPayload = ({
  preservedAttributes,
  preserved,
}: RunIdentityPayload): boolean =>
  (preservedAttributes?.length ?? 0) > 0 || (preserved?.children?.length ?? 0) > 0;
