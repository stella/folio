/**
 * The cut a kept range marker stands in (`KeptCut`): review state a
 * resolution records on a marker it kept from removed content and could not
 * place inside a piece of the record that content cut. It is never written.
 */

import type { BlockContent, KeptCut, RecordField } from "../model/document";
import { childNodes, type InlineNode, rebuildNode } from "./leaves";
import { isBoundary, isKeptBoundaryCarrier } from "./rangeBoundaries";
import { isTrackedWrapper } from "./review";

type KeptCutCarrier = Extract<InlineNode, { keptCut?: KeptCut }>;

const carriesKeptCut = (node: InlineNode): node is KeptCutCarrier =>
  isBoundary(node);

/** `node` marked as standing in `cut`, when it is a range marker. */
export const withKeptCut = (node: InlineNode, cut: KeptCut): InlineNode => {
  if (!carriesKeptCut(node)) return node;
  const known = node.keptCut;
  if (known?.depth === cut.depth && known.keptFrom === cut.keptFrom)
    return node;
  return Object.assign({}, node, { keptCut: cut });
};

/** `node` without a kept cut: placed inside a piece, it waits for nothing. */
export const withoutKeptCut = (node: InlineNode): InlineNode => {
  if (!carriesKeptCut(node) || node.keptCut === undefined) return node;
  const placed = { ...node };
  delete placed.keptCut;
  return placed;
};

/**
 * `node` with every range marker it stands for marked as standing in `cut`:
 * the marker itself, or those a pending deletion holding only kept ones
 * holds.
 */
export const withKeptCuts = (node: InlineNode, cut: KeptCut): InlineNode => {
  if (node.type !== "deletion") return withKeptCut(node, cut);
  const marked = node.content.map((child) => withKeptCuts(child, cut));
  return marked.every((child, index) => child === node.content[index])
    ? node
    : rebuildNode(node, marked);
};

/** `node` with no kept cut on the markers it stands for. */
export const withoutKeptCuts = (node: InlineNode): InlineNode => {
  if (node.type !== "deletion") return withoutKeptCut(node);
  const placed = node.content.map(withoutKeptCuts);
  return placed.every((child, index) => child === node.content[index])
    ? node
    : rebuildNode(node, placed);
};

const KEPT_CUT_FIELDS = {
  depth: true,
  keptFrom: true,
} as const satisfies Record<RecordField<KeptCut>, true>;

/**
 * Whether `cut` is a well-formed kept cut: exactly its fields, at least one
 * level deep, kept from a dated revision. The contract and the review
 * history reader share it.
 */
export const isKeptCut = (cut: unknown): cut is KeptCut =>
  typeof cut === "object" &&
  cut !== null &&
  !Array.isArray(cut) &&
  Object.keys(cut).every((key) => Object.hasOwn(KEPT_CUT_FIELDS, key)) &&
  "depth" in cut &&
  "keptFrom" in cut &&
  typeof cut.depth === "number" &&
  Number.isSafeInteger(cut.depth) &&
  cut.depth >= 1 &&
  typeof cut.keptFrom === "string";

const hasValidKeptCuts = (nodes: readonly InlineNode[]): boolean =>
  nodes.every((node) => {
    if (carriesKeptCut(node) && node.keptCut !== undefined)
      return isKeptCut(node.keptCut);
    return hasValidKeptCuts(childNodes(node) ?? []);
  });

/**
 * Whether every kept cut in `blocks` is well formed: inline markers only,
 * at least one level deep, kept from a dated revision.
 */
export const hasValidKeptCutsInBlocks = (
  blocks: readonly BlockContent[],
): boolean =>
  blocks.every((block) => {
    switch (block.type) {
      case "paragraph":
        return hasValidKeptCuts(block.content);
      case "table":
        return block.rows.every((row) =>
          row.cells.every((cell) => hasValidKeptCutsInBlocks(cell.content)),
        );
      case "blockSdt":
      case "blockCustomXml":
        return hasValidKeptCutsInBlocks(block.content);
      default:
        // A marker between blocks stands in no inline cut.
        return !("keptCut" in block);
    }
  });

/**
 * `nodes` with the kept cuts that wait on nothing cleared: a run of kept
 * markers records its cut only while a pending change stands beside it or
 * holds some of them, the change whose resolution can still let the pieces
 * meet. Whichever order resolved the rest, no cut outlives that change.
 */
export const settleKeptCuts = (
  nodes: readonly InlineNode[],
): readonly InlineNode[] => {
  let out: InlineNode[] | undefined;
  let start = 0;
  while (start < nodes.length) {
    const first = nodes[start];
    if (first === undefined || !isKeptBoundaryCarrier(first)) {
      start += 1;
      continue;
    }
    let end = start + 1;
    for (; end < nodes.length; end += 1) {
      const node = nodes[end];
      if (node === undefined || !isKeptBoundaryCarrier(node)) break;
    }
    const run = nodes.slice(start, end);
    const before = nodes[start - 1];
    const after = nodes[end];
    const waiting =
      run.some((node) => node.type === "deletion") ||
      (before !== undefined && isTrackedWrapper(before)) ||
      (after !== undefined && isTrackedWrapper(after));
    if (!waiting)
      for (const [offset, node] of run.entries()) {
        const settled = withoutKeptCuts(node);
        if (settled === node) continue;
        out ??= [...nodes];
        out[start + offset] = settled;
      }
    start = end;
  }
  return out ?? nodes;
};
