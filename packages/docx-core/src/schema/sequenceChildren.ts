/**
 * The one writer of a property set's children, in the order the schema
 * declares them.
 *
 * `CT_PPrBase` is an `xsd:sequence` of distinct optional names, so a
 * validating consumer refuses a `w:pPr` written in any other order. Four
 * writers produce one: a paragraph's own properties, a style's
 * `CT_PPrGeneral`, a numbering level's, and the `CT_PPrBase` snapshot inside
 * `w:pPrChange`. Each used to state the order as the order of its own list of
 * `if` statements, and a restated order drifts — the numbering level's wrote
 * `w:tabs` before `w:ind`, which is the reverse of the schema's.
 *
 * The order is read from the generated list instead, so the set the handler
 * map is total over and the order the serializer writes cannot disagree.
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
