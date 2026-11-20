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

/**
 * `node` marked as standing in `cut`, when it is a range marker. A marker
 * already standing in a cut keeps the revision it was first kept from, and
 * the cut widens to the deeper of the two: whichever order kept it and then
 * floated it with another removal's seam, it records the same cut.
 */
export const withKeptCut = (node: InlineNode, cut: KeptCut): InlineNode => {
  if (!carriesKeptCut(node)) return node;
  const known = node.keptCut;
  const next =
    known === undefined
      ? cut
      : {
          depth: Math.max(known.depth, cut.depth),
          keptFrom: known.keptFrom,
        };
  if (known?.depth === next.depth && known.keptFrom === next.keptFrom)
    return node;
  return Object.assign({}, node, { keptCut: next });
};

/**
 * Whether `node` stands in a kept cut: a marked range marker, or a pending
 * deletion holding only marked ones.
 */
export const standsInKeptCut = (node: InlineNode): boolean =>
  node.type === "deletion"
    ? node.content.length > 0 && node.content.every(standsInKeptCut)
    : carriesKeptCut(node) && node.keptCut !== undefined;

/** Whether some range marker in `nodes`, at any depth, stands in a kept cut. */
export const holdsKeptCut = (nodes: readonly InlineNode[]): boolean =>
  nodes.some(
    (node) =>
      (carriesKeptCut(node) && node.keptCut !== undefined) ||
      holdsKeptCut(childNodes(node) ?? []),
  );

/** `node` without a kept cut: placed inside a piece, it waits for nothing. */
export const withoutKeptCut = (node: InlineNode): InlineNode => {
  if (!carriesKeptCut(node) || node.keptCut === undefined) return node;
  const placed = { ...node };
  delete placed.keptCut;
  return placed;
};

/**
 * `node` with `change` made to every range marker it stands for: the marker
 * itself, or those a pending deletion holding only kept ones holds.
 */
const changeStandingMarkers = (
  node: InlineNode,
  change: (marker: InlineNode) => InlineNode,
): InlineNode => {
  if (node.type !== "deletion") return change(node);
  const changed = node.content.map((child) =>
    changeStandingMarkers(child, change),
  );
  return changed.every((child, index) => child === node.content[index])
    ? node
    : rebuildNode(node, changed);
};

/**
 * `change`, keeping what the resolution recorded about a marker (`known`,
 * keyed by the marker itself) on the copy it makes.
 */
const keepingRecord =
  (
    change: (marker: InlineNode) => InlineNode,
    known: WeakMap<InlineNode, string>,
  ) =>
  (marker: InlineNode): InlineNode => {
    const next = change(marker);
    const date = known.get(marker);
    if (next !== marker && date !== undefined) known.set(next, date);
    return next;
  };

/**
 * `node` with every range marker it stands for marked as standing in `cut`:
 * the marker itself, or those a pending deletion holding only kept ones
 * holds.
 */
export const withKeptCuts = (node: InlineNode, cut: KeptCut): InlineNode =>
  changeStandingMarkers(node, (marker) => withKeptCut(marker, cut));

/** {@link withKeptCuts}, keeping what the resolution recorded per marker. */
export const withKeptCutsRecorded = (
  node: InlineNode,
  cut: KeptCut,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord((marker) => withKeptCut(marker, cut), known),
  );

/** `node` with no kept cut on the markers it stands for. */
export const withoutKeptCuts = (node: InlineNode): InlineNode =>
  changeStandingMarkers(node, withoutKeptCut);

/**
 * {@link withoutKeptCuts}, keeping what the resolution recorded about each
 * marker on its cleared copy.
 */
export const withoutKeptCutsRecorded = (
  node: InlineNode,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(node, keepingRecord(withoutKeptCut, known));

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
 * Whether a pending change stands at `node`'s `edge`: the node itself, or
 * one at that edge of the records it holds, whose resolution can still let
 * the pieces across from it meet.
 */
export const pendingAtEdge = (
  node: InlineNode | undefined,
  edge: "start" | "end",
): boolean => {
  if (node === undefined) return false;
  if (isTrackedWrapper(node)) return true;
  const children = childNodes(node);
  if (children === undefined) return false;
  return pendingAtEdge(
    edge === "start" ? children.at(0) : children.at(-1),
    edge,
  );
};

/** The deepest kept cut among the markers `node` stands for; 0 for none. */
const deepestKeptCut = (node: InlineNode): number => {
  if (node.type === "deletion")
    return Math.max(0, ...node.content.map(deepestKeptCut));
  if (!carriesKeptCut(node)) return 0;
  return node.keptCut?.depth ?? 0;
};

/**
 * `node` with every marked range marker it stands for widened to `depth`,
 * keeping what the resolution recorded per marker. An unmarked marker
 * stays unmarked: it was authored or accepted, not kept.
 */
const withKeptCutDepthRecorded = (
  node: InlineNode,
  depth: number,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord(
      (marker) =>
        carriesKeptCut(marker) && marker.keptCut !== undefined
          ? withKeptCut(marker, { depth, keptFrom: marker.keptCut.keptFrom })
          : marker,
      known,
    ),
  );

/** A paragraph join's cut, and when the paragraph mark it resolves was made. */
type JoinCut = { depth: number; after: string };

/**
 * `nodes` with the kept markers standing at its `edge` that were kept from
 * revisions made after the split widened to the join's cut: they stood in
 * its seam, as a join resolved earlier hands its depth to the removal that
 * keeps them. A marker kept from content the split cut keeps its own cut.
 */
export const withEdgeKeptCutsWidened = (
  nodes: readonly InlineNode[],
  edge: "start" | "end",
  { depth, after }: JoinCut,
): readonly InlineNode[] => {
  let out: InlineNode[] | undefined;
  for (let step = 0; step < nodes.length; step += 1) {
    const index = edge === "start" ? step : nodes.length - 1 - step;
    const node = nodes[index];
    if (node === undefined || !isKeptBoundaryCarrier(node)) break;
    const widened = changeStandingMarkers(node, (marker) =>
      carriesKeptCut(marker) &&
      marker.keptCut !== undefined &&
      marker.keptCut.keptFrom > after
        ? withKeptCut(marker, { depth, keptFrom: marker.keptCut.keptFrom })
        : marker,
    );
    if (widened === node) continue;
    out ??= [...nodes];
    out[index] = widened;
  }
  return out ?? nodes;
};

type SettleKeptCutsOptions = {
  /** What the resolution recorded about each marker, kept on its copies. */
  known?: WeakMap<InlineNode, string>;
  /** Told of every record rebuilt around a cleared marker. */
  rebuilt?: (from: InlineNode, to: InlineNode) => void;
  /** Whether a pending change meets the list from outside, at each end. */
  outside?: { start: boolean; end: boolean };
};

/**
 * `nodes` with the kept cuts that wait on nothing cleared, at every depth: a
 * run of kept markers records its cut only while a pending change stands
 * beside it, across the edges of the records around it, or holds some of
 * them: the change whose resolution can still let the pieces meet. A record
 * holding the run is no such change; resolving it keeps them again.
 * Whichever order resolved the rest, no cut outlives that change. A waiting
 * run stands in one cut, as deep as the deepest any of its markers records.
 */
export const settleKeptCuts = (
  nodes: readonly InlineNode[],
  {
    known = new WeakMap(),
    rebuilt,
    outside = { start: false, end: false },
  }: SettleKeptCutsOptions = {},
): readonly InlineNode[] => {
  let out: InlineNode[] | undefined;
  const pendingBefore = (index: number): boolean =>
    index === 0 ? outside.start : pendingAtEdge(nodes[index - 1], "end");
  const pendingAfter = (index: number): boolean =>
    index === nodes.length - 1
      ? outside.end
      : pendingAtEdge(nodes[index + 1], "start");
  let start = 0;
  while (start < nodes.length) {
    const first = nodes[start];
    if (first === undefined) break;
    if (!isKeptBoundaryCarrier(first)) {
      const children = childNodes(first);
      const settled =
        children === undefined
          ? children
          : settleKeptCuts(children, {
              known,
              ...(rebuilt === undefined ? {} : { rebuilt }),
              outside: {
                start: pendingBefore(start),
                end: pendingAfter(start),
              },
            });
      if (children !== undefined && settled !== children) {
        const next = rebuildNode(first, settled ?? []);
        rebuilt?.(first, next);
        out ??= [...nodes];
        out[start] = next;
      }
      start += 1;
      continue;
    }
    let end = start + 1;
    for (; end < nodes.length; end += 1) {
      const node = nodes[end];
      if (node === undefined || !isKeptBoundaryCarrier(node)) break;
    }
    const run = nodes.slice(start, end);
    const waiting =
      run.some((node) => node.type === "deletion") ||
      pendingBefore(start) ||
      pendingAfter(end - 1);
    // Markers standing together stand in one cut: the deepest any of them
    // records. Whether removals met at one seam at once or one after another,
    // each keeps that depth.
    const depth = Math.max(0, ...run.map(deepestKeptCut));
    for (const [offset, node] of run.entries()) {
      const settled = waiting
        ? withKeptCutDepthRecorded(node, depth, known)
        : withoutKeptCutsRecorded(node, known);
      if (settled === node) continue;
      out ??= [...nodes];
      out[start + offset] = settled;
    }
    start = end;
  }
  return out ?? nodes;
};
