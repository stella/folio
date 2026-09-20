/**
 * The one writer of a property set's children, in the order the schema
 * declares them.
 *
 * `w:rPr`, `w:tblPr` and `w:sectPr` each declare their children in one order,
 * and folio writes them in it. How binding that order is differs by container:
 * `CT_TblPrBase` is an `xsd:sequence` of distinct names and a validating
 * consumer refuses a `w:tblPr` written in any other order, while `EG_RPrBase`
 * is an `xsd:choice` referenced `maxOccurs="unbounded"`, so a `w:rPr` in any
 * order is valid. Writing the canonical order anyway is what keeps two
 * serializers of the same element from disagreeing, and it is what Word
 * writes.
 *
 * The order is read from the generated list rather than restated as the order
 * of a list of `if` statements — the restatement is what drifted, twice: once
 * inside folio-core, and once as a second `w:rPr` writer in this package that
 * put `w:sz` and `w:highlight` before `w:rFonts`.
 */

import type { PreservedMarkup } from "../model/preservedMarkup";
import {
  SEQUENCE_CHILDREN,
  type SequenceChild,
  type SequenceContainer,
} from "./sequenceChildren.gen";

export {
  SEQUENCE_CHILDREN,
  type SequenceChild,
  type SequenceContainer,
} from "./sequenceChildren.gen";

type SerializeSequenceChildrenOptions<Container extends SequenceContainer> = {
  container: Container;
  /**
   * The modelled children keyed by element name rather than pre-ordered, so
   * the caller cannot state an order of its own.
   *
   * An empty string means the caller wrote nothing for that child.
   */
  modelled: ReadonlyArray<readonly [name: SequenceChild<Container>, xml: string]>;
  /**
   * The container's unmodelled markup, each capture carrying the schema
   * ordinal it was read at, so it lands between the same two neighbours.
   *
   * A caller that builds a document from scratch has none.
   */
  preserved?: PreservedMarkup | undefined;
};

/**
 * A property set's children, modelled and captured, in schema order.
 *
 * Ties are only possible between a declared child and a capture sharing its
 * slot — an undeclared name takes the place of the last declared child before
 * it. The sort is stable and the modelled child leads.
 */
export const serializeSequenceChildren = <Container extends SequenceContainer>({
  container,
  modelled,
  preserved,
}: SerializeSequenceChildrenOptions<Container>): string[] => {
  const declared: readonly string[] = SEQUENCE_CHILDREN[container];
  const placed: Array<{ at: number; rank: number; xml: string }> = [];
  for (const [name, xml] of modelled) {
    if (xml.length > 0) {
      placed.push({ at: declared.indexOf(name), rank: 0, xml });
    }
  }
  for (const { index, xml } of preserved?.children ?? []) {
    placed.push({ at: index, rank: 1, xml });
  }
  return placed
    .sort((left, right) => left.at - right.at || left.rank - right.rank)
    .map(({ xml }) => xml);
};
