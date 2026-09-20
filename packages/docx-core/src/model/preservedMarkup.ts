/**
 * The ordered verbatim sink: where a container's unmodelled markup goes.
 *
 * A parser that walks a container's children, models the ones it recognises
 * and lets the rest fall off the end of a `switch` loses them silently, and
 * verbatim replay hides the loss until somebody edits the document. The sink
 * is the other branch of that `switch`. It keeps the markup as bytes together
 * with **where it was**, so the serializer puts it back between the same
 * modelled siblings rather than at the end of the container, which is what the
 * schema's ordered content models require.
 *
 * The position is recorded as an ordinal, not a pointer. For a container that
 * models one kind of child, `index` is how many modelled children it had read
 * when the markup arrived: that survives the model being edited — inserting a
 * paragraph shifts what comes after it, which is what a reader would expect —
 * and it needs no identity on the modelled children. For a container whose
 * content model is one flat sequence — a property set such as `w:tblPr` or
 * `w:sectPr` — it is the child's own position in that sequence, because there
 * position is a property of the name; a count would mirror whichever
 * properties folio models today and would move under the capture the moment
 * one more of them was modelled.
 *
 * One field, `preserved`, carries both halves for every container that has a
 * sink, so a reader never has to know which of two names a given container
 * used.
 */

/** Markup the container's model does not hold, and where it sat. */
export type PreservedChild = {
  /**
   * Where the markup goes back. `0` puts it before the first modelled child;
   * the container's modelled count puts it after the last. In a sequence
   * container it is the schema ordinal instead, and the two halves merge by it.
   */
  index: number;
  /** Replayable markup for one child, as `captureVerbatimXml` wrote it. */
  xml: string;
};

/**
 * A container's unmodelled markup. Absent means the parser read every child
 * the source carried, which is the state a fully modelled container is in; an
 * empty record is never written.
 */
export type PreservedMarkup = {
  /** Ordered by `index`, then by source order within an index. */
  children?: PreservedChild[];
};

/**
 * One attribute an element carried that its parser has no field for.
 *
 * The remainder is a sibling of the child sink, not a member of it: an
 * attribute has no position among children to keep, so it rides the element's
 * own model record — `Paragraph.preservedAttributes` and its siblings — and
 * the serializer writes it back into the same start tag.
 *
 * The name is resolved, never the source's spelling. folio reads an attribute
 * by namespace URI plus local name, so a remainder that kept `"w:rsidR"`
 * textually would keep a second copy of the `w:rsidR` a source spelled
 * `altw:rsidR` under a second binding of the same URI. `namespace` is absent
 * for an unprefixed attribute, which has none; a namespace declaration is not
 * an attribute of the element in this sense and is never in the remainder,
 * because the rebuilt part binds its own prefixes.
 */
export type PreservedAttribute = {
  /** Resolved namespace URI, absent for an unprefixed attribute. */
  namespace?: string;
  /** Local name, without the prefix the source happened to bind. */
  name: string;
  value: string;
};
