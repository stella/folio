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
 * The position is recorded as a count, not a pointer: `index` is how many
 * modelled children the container had read when the markup arrived. That
 * survives the model being edited — inserting a paragraph shifts what comes
 * after it, which is what a reader would expect — and it needs no identity on
 * the modelled children.
 *
 * One field, `preserved`, carries both halves for every container that has a
 * sink, so a reader never has to know which of two names a given container
 * used.
 */

/** Markup the container's model does not hold, and where it sat. */
export type PreservedChild = {
  /**
   * Modelled children that preceded this markup in the source. `0` puts it
   * before the first modelled child; the container's modelled count puts it
   * after the last.
   */
  index: number;
  /** Replayable markup for one child, as `captureVerbatimXml` wrote it. */
  xml: string;
};

/** An attribute the container's model does not hold, kept as written. */
export type PreservedAttribute = {
  /** Qualified name exactly as the source spelled it, e.g. `w:rsidR`. */
  name: string;
  value: string;
};

/**
 * A container's unmodelled markup. Absent means the parser read every child
 * and every attribute the source carried, which is the state a fully modelled
 * container is in; an empty record is never written.
 */
export type PreservedMarkup = {
  /** Ordered by `index`, then by source order within an index. */
  children?: PreservedChild[];
  /** In source order. */
  attributes?: PreservedAttribute[];
};
