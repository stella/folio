/**
 * Which editor carrier holds a container's ordered verbatim sink.
 *
 * `preservedMarkup.ts` explains why a container whose model is a list of one
 * kind of child keeps its unmodelled markup in a `preserved` sink beside that
 * list rather than as a member of it. The sink closed the save leg. The editor
 * leg is a second question, and it was never asked per container: a record
 * whose sink nothing projected lost it on the first open, which is how a
 * `w:bookmarkEnd` written beside a table's rows came back gone and left its
 * `w:bookmarkStart` unpaired.
 *
 * The question is asked here once, for every record that declares a sink, and
 * the union is derived rather than listed. A model record that gains
 * `preserved` joins {@link SinkBearingBlock} and the map below stops being
 * total, so the next sink cannot reach the editor without somebody deciding
 * how it gets back.
 *
 * `Comment.preserved` is not here and does not need to be: a comment body is
 * not a node in the editor's document, it travels on the package the save
 * writes from, and the round trip carries it untouched.
 */

import type { BlockContent, PreservedMarkup, TableCell, TableRow } from "../../types/document";

/**
 * `Record_` when it declares `preserved`, `never` when it does not.
 *
 * A record without the field still satisfies `{ preserved?: PreservedMarkup }`,
 * so `Extract` would match everything. Inference tells the two apart: a record
 * that has no such field infers `unknown`, and one that does infers the sink's
 * own type.
 */
type DeclaresSink<Record_> = Record_ extends { preserved?: infer Sink }
  ? [unknown] extends [Sink]
    ? never
    : Record_
  : never;

/**
 * Every block-level record whose model holds an ordered verbatim sink.
 *
 * The universe is the whole block spine the editor walks — blocks, rows and
 * cells — so a sink added anywhere along it joins the union rather than
 * slipping past a hand-listed pair.
 */
export type SinkBearingBlock = DeclaresSink<BlockContent | TableRow | TableCell>;

/**
 * The editor carries a sink on the node's own attrs, by reference.
 *
 * Reference identity is what tells an authored record from a copy the editor
 * made, exactly as it does for the attribute remainder: two records that each
 * parsed their own children hold different arrays however equal their
 * contents, so the sink can follow the record it was authored on and no other.
 */
export const SINK_ON_NODE_ATTRS = "carried-on-node-attrs";

/** How the editor carries each sink-bearing record's markup. */
export const PRESERVED_SINK_CARRIERS = {
  table: SINK_ON_NODE_ATTRS,
  tableRow: SINK_ON_NODE_ATTRS,
} as const satisfies Record<SinkBearingBlock["type"], typeof SINK_ON_NODE_ATTRS>;

/** A sink worth carrying: an empty record is never written. */
export const hasSinkChildren = (
  preserved: PreservedMarkup | undefined,
): preserved is PreservedMarkup => preserved !== undefined && (preserved.children?.length ?? 0) > 0;
