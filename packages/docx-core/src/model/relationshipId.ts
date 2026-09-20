import { panic } from "better-result";

/**
 * A relationship reference that names something.
 *
 * A reference in the body (`r:embed` on an `a:blip`, `r:id` on a
 * `v:imagedata`) is a key into the owning part's relationships. A drawing that
 * carries none is common: a chart, an OLE frame, an anchor with no
 * `a:graphic`, and every preview folio draws for markup it cannot project.
 * Absence has one spelling, `undefined`, so it can never reach a relationship
 * lookup as a key or be written back as `r:embed=""`.
 *
 * The brand is what makes that a property of the type rather than a rule
 * producers are asked to remember. The empty string is not a `RelationshipId`,
 * so a field typed as one cannot hold the second spelling of absence, and a
 * reference read off the source becomes one only through
 * {@link relationshipIdOf}.
 */
export type RelationshipId = string & { readonly __brand: "folio.relationshipId" };

/**
 * Whether a string names a relationship. The refinement is a type predicate
 * rather than a cast, so no site that mints an id needs an assertion.
 */
export const isRelationshipId = (value: string): value is RelationshipId => value.length > 0;

/**
 * The relationship a reference names, or `undefined` when it names none.
 *
 * Takes what a reader hands back (a missing attribute, an attribute present
 * and empty, a reference) and answers in the one spelling the model has for
 * absence.
 */
export const relationshipIdOf = (value: string | null | undefined): RelationshipId | undefined =>
  value === undefined || value === null || !isRelationshipId(value) ? undefined : value;

/**
 * The id a save gives a part it added, from the ordinal it allocated.
 *
 * A mint is not a read, so it has no absent case to report: `rId` followed by
 * a number names something by construction, and a build where it does not is
 * one where the refinement has stopped describing the type.
 */
export const mintRelationshipId = (ordinal: number): RelationshipId => {
  const id = `rId${ordinal}`;
  return isRelationshipId(id) ? id : panic("A minted relationship id names nothing.");
};
