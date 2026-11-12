/**
 * Where the blocks a batch puts beside an anchor land, and in what order.
 *
 * Every block insertion of a batch (paragraphs, tables, signature tables)
 * names a gap between blocks, read once as accepting every change leaves the
 * document (`insertionGap`), and lands in it in the order the batch names it
 * (`InsertionGapOrder`). Both hold in every mode: applied directly, as
 * tracked changes and as suggestions, so accepting leaves the blocks of one
 * gap in the order applying them directly does, whatever order the batch
 * runs them in and wherever a pending change sends the first one.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import { expectParagraphAttrs } from "./attrs";
import { paragraphGoes, tableGoes } from "./tableRunIn";

const isTable = (node: PMNode): boolean =>
  node.type.spec["tableRole"] === "table";

/** Whether accepting every change removes `block` whole. */
const goesWhole = (block: PMNode): boolean =>
  isTable(block) ? tableGoes(block, "accept") : paragraphGoes(block, "accept");

/**
 * The gap a block put at `boundary`, between blocks, lands in. Put right
 * after a paragraph (`afterParagraph`), it lands there, even after one
 * accepting removes: the batch named that place. Put anywhere else, it lands
 * before the blocks right before `boundary` that accepting removes whole,
 * tables and paragraphs alike: accepted, the two sides of those blocks are
 * one gap, so blocks put on either side of them share it, as they do once
 * those blocks are removed directly. A paragraph a pending join runs into,
 * from words before it whose break goes, is part of that join rather than
 * the gap (`pendingJoinEnd` reads where the join ends), so the gap stops
 * after it, past the tables before it alone.
 */
export const insertionGap = (
  doc: PMNode,
  boundary: number,
  afterParagraph: boolean,
): number => {
  if (afterParagraph) return boundary;
  const $boundary = doc.resolve(boundary);
  const parent = $boundary.parent;
  let first = $boundary.index();
  while (first > 0 && goesWhole(parent.child(first - 1))) first -= 1;
  const reached = parent.maybeChild(first - 1);
  // Accepting removes a break pending deletion, or a move's source.
  const reachedMark =
    reached?.type.name === "paragraph"
      ? expectParagraphAttrs(reached).pPrMark?.kind
      : undefined;
  const joinReaches = reachedMark === "del" || reachedMark === "moveFrom";
  let at = boundary;
  for (let index = $boundary.index(); index > first; index--) {
    const before = parent.child(index - 1);
    if (joinReaches && !isTable(before)) return at;
    at -= before.nodeSize;
  }
  return at;
};

type Landed = {
  /** Where in the batch it was named: its input index. */
  order: number;
  from: number;
  to: number;
  /** The transaction's step count `from` and `to` are read at. */
  steps: number;
};

/**
 * The blocks a batch has put into each gap so far, by the gap's position in
 * the document the batch read. A block lands next to the blocks of its gap
 * already in, by the order the batch names them: right after the last one
 * named before it, else right before the first one named after it. Only the
 * first block of a gap to run lands where its own placement puts it. The batch runs
 * its operations in an order of its own (right to left, a retraction before
 * a table it places), so no position read before it ran can say this.
 */
export class InsertionGapOrder {
  readonly #gaps = new Map<number, Landed[]>();

  /**
   * Where blocks the batch names at `orders` (ascending) put into `gap` go
   * now, beside `container`, a position in the container the gap is in.
   * Undefined while none of that gap is in: they go where they are placed.
   */
  landing(
    tr: Transaction,
    gap: number,
    orders: readonly number[],
    container: number,
  ): number | undefined {
    const first = orders[0];
    const last = orders.at(-1);
    if (first === undefined || last === undefined) return undefined;
    const $container = tr.doc.resolve(container);
    let before: { order: number; at: number } | undefined;
    let after: { order: number; at: number } | undefined;
    for (const landed of this.#gaps.get(gap) ?? []) {
      const mapping = tr.mapping.slice(landed.steps);
      const from = mapping.map(landed.from, 1);
      const to = mapping.map(landed.to, -1);
      // Removed since, or moved out of the gap's container.
      if (to <= from || !tr.doc.resolve(from).sameParent($container)) continue;
      if (landed.order < first && (!before || landed.order > before.order))
        before = { order: landed.order, at: to };
      if (landed.order > last && (!after || landed.order < after.order))
        after = { order: landed.order, at: from };
    }
    return before?.at ?? after?.at;
  }

  /** Note the blocks just put at `at`, in document order, `size` long each. */
  record(
    tr: Transaction,
    gap: number,
    landed: readonly { order: number; size: number }[],
    at: number,
  ): void {
    const entries = this.#gaps.get(gap) ?? [];
    let from = at;
    for (const { order, size } of landed) {
      entries.push({
        order,
        from,
        to: from + size,
        steps: tr.mapping.maps.length,
      });
      from += size;
    }
    this.#gaps.set(gap, entries);
  }
}
