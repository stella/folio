/**
 * The cut a kept range marker stands in (`KeptCut`): review state a
 * resolution records on a marker it kept from removed content and could not
 * place inside a piece of the record that content cut.
 */

import { MAX_REVISION_ID } from "../model/document";
import type {
  BlockContent,
  DocumentBody,
  KeptCut,
  KeptCutHolders,
  Paragraph,
  ParagraphContent,
  RecordField,
} from "../model/document";
import { blockListAt, storyParagraphs } from "./blocks";
import {
  asParagraphContent,
  childNodes,
  type InlineNode,
  rebuildNode,
} from "./leaves";
import {
  isBoundary,
  isKeptBoundaryCarrier,
  joinNeighbour,
} from "./rangeBoundaries";
import { isTrackedWrapper, type TrackedWrapper } from "./review";

type KeptCutCarrier = Extract<InlineNode, { keptCut?: KeptCut }>;

const carriesKeptCut = (node: InlineNode): node is KeptCutCarrier =>
  isBoundary(node);

/**
 * The side of a kept marker whose pending change holds a piece of its cut,
 * as its holders name it: `both` where changes either side hold one. Derived
 * from {@link KeptCutHolders}; never stored.
 */
export type KeptCutPiece = "before" | "after" | "both";

export const isKeptCutPiece = (value: unknown): value is KeptCutPiece =>
  value === "before" || value === "after" || value === "both";

export const keptCutHolders = ({
  before,
  after,
}: {
  before?: number | undefined;
  after?: number | undefined;
}): KeptCutHolders | undefined => {
  if (before === undefined) return after === undefined ? undefined : { after };
  return after === undefined ? { before } : { before, after };
};

/** The side `holders` name. */
const pieceOf = (
  holders: KeptCutHolders | undefined,
): KeptCutPiece | undefined => {
  if (holders === undefined) return undefined;
  if (
    holders.scope === "seam" ||
    (holders.before !== undefined && holders.after !== undefined)
  )
    return "both";
  return holders.before === undefined ? "after" : "before";
};

/** Retain the seam scope when its pending holders change. */
const withHolderScope = (
  holders: KeptCutHolders | undefined,
  scope: KeptCutHolders["scope"],
): KeptCutHolders | undefined =>
  holders === undefined || scope === undefined
    ? holders
    : { ...holders, scope };

/** `cut` with `holders`, or with none. */
const cutHeldBy = (
  cut: KeptCut,
  holders: KeptCutHolders | undefined,
): KeptCut =>
  holders === undefined
    ? { keptFrom: cut.keptFrom }
    : { keptFrom: cut.keptFrom, holders };

/**
 * `node` marked as standing in `cut`, when it is a range marker. A marker
 * already standing in a cut keeps the revision it was first kept from and
 * the holders first named for it: whichever order kept it and then floated
 * it with another removal's seam, it records the same cut.
 */
const withKeptCut = (node: InlineNode, cut: KeptCut): InlineNode => {
  if (!carriesKeptCut(node)) return node;
  const known = node.keptCut;
  if (known === undefined) return Object.assign({}, node, { keptCut: cut });
  if (known.holders !== undefined || cut.holders === undefined) return node;
  return Object.assign({}, node, {
    keptCut: cutHeldBy(known, cut.holders),
  });
};

/**
 * `node` with every marked range marker it stands for naming `holders` as
 * the pending changes beside it holding a piece of its cut, keeping what the
 * resolution recorded per marker. An unmarked marker stays unmarked, and one
 * already naming holders keeps them.
 */
export const withKeptCutHoldersRecorded = (
  node: InlineNode,
  holders: KeptCutHolders,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord(
      (marker) =>
        carriesKeptCut(marker) && marker.keptCut !== undefined
          ? withKeptCut(marker, cutHeldBy(marker.keptCut, holders))
          : marker,
      known,
    ),
  );

/** Whether all markers in this record name the revision on the given side. */
export const keptCutHeldBy = (
  node: InlineNode,
  { side, id }: { side: "before" | "after"; id: number },
): boolean => {
  if (node.type === "deletion")
    return (
      node.content.length > 0 &&
      node.content.every((child) => keptCutHeldBy(child, { side, id }))
    );
  return carriesKeptCut(node) && node.keptCut?.holders?.[side] === id;
};

/** Move a cut's holder to the revision taking its join. */
export const withKeptCutHolderHandedOver = (
  node: InlineNode,
  { side, from, to }: { side: "before" | "after"; from: number; to: number },
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord((marker) => {
      if (!carriesKeptCut(marker)) return marker;
      const cut = marker.keptCut;
      const holders = cut?.holders;
      if (cut === undefined || holders?.[side] !== from) return marker;
      const next =
        side === "before"
          ? { ...holders, before: to }
          : { ...holders, after: to };
      return Object.assign({}, marker, { keptCut: cutHeldBy(cut, next) });
    }, known),
  );

/** `node` with every marker it stands for naming no holder of its cut. */
export const withKeptCutPieceConsumed = (
  node: InlineNode,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord((marker) => {
      if (!carriesKeptCut(marker) || marker.keptCut?.holders === undefined)
        return marker;
      return Object.assign({}, marker, {
        keptCut: cutHeldBy(marker.keptCut, undefined),
      });
    }, known),
  );

/**
 * `node` with the holder ids of the markers it stands for redirected by
 * `successor`: a holder merged into another change, or given a fresh id,
 * still holds the piece under its new id.
 */
export const withKeptCutHoldersRenamed = (
  node: InlineNode,
  successor: (id: number) => number,
): InlineNode =>
  changeStandingMarkers(node, (marker) => {
    const holders = carriesKeptCut(marker)
      ? marker.keptCut?.holders
      : undefined;
    if (
      !carriesKeptCut(marker) ||
      marker.keptCut === undefined ||
      holders === undefined
    )
      return marker;
    const before =
      holders.before === undefined ? undefined : successor(holders.before);
    const after =
      holders.after === undefined ? undefined : successor(holders.after);
    if (before === holders.before && after === holders.after) return marker;
    const renamed = withHolderScope(
      keptCutHolders({ before, after }),
      holders.scope,
    );
    return Object.assign({}, marker, {
      keptCut: cutHeldBy(marker.keptCut, renamed),
    });
  });

/** The earliest revision a marker `node` stands for was kept from. */
export const keptCutFrom = (node: InlineNode): string | undefined => {
  if (node.type === "deletion")
    return node.content
      .flatMap((child) => keptCutFrom(child) ?? [])
      .toSorted()
      .at(0);
  return carriesKeptCut(node) ? node.keptCut?.keptFrom : undefined;
};

/**
 * The piece the markers `node` stands for name, when every marked one names
 * the same side.
 */
export const keptCutPiece = (node: InlineNode): KeptCutPiece | undefined => {
  if (node.type === "deletion") {
    const pieces = new Set(node.content.map(keptCutPiece));
    return pieces.size === 1 ? pieces.values().next().value : undefined;
  }
  return carriesKeptCut(node) ? pieceOf(node.keptCut?.holders) : undefined;
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

/** `node` without a kept cut: it waits for nothing. */
const withoutKeptCut = (node: InlineNode): InlineNode => {
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
const withKeptCuts = (node: InlineNode, cut: KeptCut): InlineNode =>
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

/**
 * `node` with no kept cut on the markers it stands for, keeping what the
 * resolution recorded about each marker on its cleared copy.
 */
const withoutKeptCutsRecorded = (
  node: InlineNode,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(node, keepingRecord(withoutKeptCut, known));

type SeamJoins = NonNullable<TrackedWrapper["resolutionJoins"]>;

/**
 * Kept markers a change's joins name in review history version 1: the
 * revision they were kept from and how many boundaries they are, nearest
 * the seam (all of those standing there when no count is given).
 */
export type VersionOneKeptMarkers = { keptFrom: string; keptCount?: number };

/**
 * What review history version 1 wrote on a change about kept markers: those
 * its own joins name (`change`), and those each of its deferred seam groups
 * names, by the group's place in `deferredRemove`.
 */
export type VersionOneKept = {
  change?: VersionOneKeptMarkers;
  groups: readonly (VersionOneKeptMarkers | undefined)[];
};

/**
 * `nodes` with the kept markers that pending changes' joins name, as review
 * history version 1 wrote them (on the change or on a deferred seam's group,
 * no cut on the marker), marked as the version that follows marks them: the
 * markers standing nearest each side the change joins, or beside the group's
 * blockers, as many boundaries as it counts, stand in that cut, naming the
 * change's side (the blockers', or both) as its piece.
 * `writtenOnChange` gives what version 1 wrote on a change's joins.
 * Markers already standing in a cut keep it.
 */
const withKeptCutsFromChanges = (
  nodes: readonly InlineNode[],
  writtenOnChange: (joins: SeamJoins) => VersionOneKept | undefined,
): readonly InlineNode[] => {
  const out = nodes.map((node) => {
    const children = childNodes(node);
    if (children === undefined) return node;
    const marked = withKeptCutsFromChanges(children, writtenOnChange);
    return marked === children ? node : rebuildNode(node, [...marked]);
  });
  const mark = (
    from: number,
    step: 1 | -1,
    { keptFrom, count }: { keptFrom: string; count: number },
    holders: KeptCutHolders,
  ): void => {
    let boundaries = 0;
    for (let index = from; ; index += step) {
      const node = out[index];
      if (
        node === undefined ||
        !isKeptBoundaryCarrier(node) ||
        holdsKeptCut([node])
      )
        return;
      boundaries += keptBoundaryCount(node);
      if (boundaries > count) return;
      out[index] = withKeptCuts(node, { keptFrom, holders });
    }
  };
  for (const [index, node] of nodes.entries()) {
    const joins = isTrackedWrapper(node) ? node.resolutionJoins : undefined;
    const kept =
      joins === undefined ? undefined : writtenOnChange(joins)?.change;
    if (joins === undefined || kept === undefined || !isTrackedWrapper(node))
      continue;
    const count = kept.keptCount ?? Number.POSITIVE_INFINITY;
    const named = { keptFrom: kept.keptFrom, count };
    if (joins.before > 0) mark(index - 1, -1, named, { after: node.info.id });
    if (joins.after > 0) mark(index + 1, 1, named, { before: node.info.id });
  }
  // A deferred seam's group dates the boundaries kept standing at the seam:
  // the run of unmarked markers its blockers stand beside, naming the
  // blockers' side, or both, as its piece.
  type KeptGroup = VersionOneKeptMarkers & {
    depth: number;
    blockers: readonly number[];
  };
  const keptGroups = (node: InlineNode | undefined): readonly KeptGroup[] => {
    const joins =
      node !== undefined && isTrackedWrapper(node)
        ? node.resolutionJoins
        : undefined;
    const written = joins === undefined ? undefined : writtenOnChange(joins);
    if (joins === undefined || written === undefined) return [];
    return (joins.deferredRemove ?? []).flatMap(({ depth, blockers }, at) => {
      const kept = written.groups[at];
      return kept === undefined ? [] : [{ ...kept, depth, blockers }];
    });
  };
  const seamOf = ({ depth, blockers, keptFrom }: KeptGroup): string =>
    `${depth}:${blockers.join(",")}:${keptFrom}`;
  const unmarked = (node: InlineNode | undefined): boolean =>
    node !== undefined && isKeptBoundaryCarrier(node) && !holdsKeptCut([node]);
  let start = 0;
  while (start < out.length) {
    if (!unmarked(out[start])) {
      start += 1;
      continue;
    }
    let end = start + 1;
    while (end < out.length && unmarked(out[end])) end += 1;
    const before = keptGroups(start === 0 ? undefined : out[start - 1]);
    const after = keptGroups(out[end]);
    const afterSeams = new Set(after.map(seamOf));
    const across = before.find((group) => afterSeams.has(seamOf(group)));
    const group = across ?? after.at(0) ?? before.at(0);
    // The groups' changes stand either side of the run: they hold its piece.
    const beforeId = changeAtEdge(
      start === 0 ? undefined : out[start - 1],
      "end",
    );
    const afterId = changeAtEdge(out[end], "start");
    if (group !== undefined) {
      const cut = {
        keptFrom: group.keptFrom,
        count: group.keptCount ?? Number.POSITIVE_INFINITY,
      };
      if (
        across !== undefined &&
        beforeId !== undefined &&
        afterId !== undefined
      )
        mark(start, 1, cut, { before: beforeId, after: afterId });
      else if (after.includes(group) && afterId !== undefined)
        mark(end - 1, -1, cut, { after: afterId });
      else if (beforeId !== undefined)
        mark(start, 1, cut, { before: beforeId });
    }
    start = end;
  }
  return out.every((node, index) => node === nodes[index]) ? nodes : out;
};

/** {@link withKeptCutsFromChanges} over a paragraph's content. */
export const paragraphKeptCutsFromChanges = (
  content: readonly ParagraphContent[],
  writtenOnChange: (joins: SeamJoins) => VersionOneKept | undefined,
  legacyPiece?: LegacyKeptCutPiece,
): ParagraphContent[] => {
  const cuts = withKeptCutsFromChanges(content, writtenOnChange);
  return asParagraphContent(
    legacyPiece === undefined ? cuts : settleKeptCuts(cuts, { legacyPiece }),
  );
};

/** How many range boundaries a kept marker `node` stands for. */
const keptBoundaryCount = (node: InlineNode): number =>
  node.type === "deletion"
    ? node.content.reduce((sum, child) => sum + keptBoundaryCount(child), 0)
    : 1;

const KEPT_CUT_HOLDER_FIELDS = {
  before: true,
  after: true,
  scope: true,
} as const satisfies Record<RecordField<KeptCutHolders>, true>;

const KEPT_CUT_FIELDS = {
  keptFrom: true,
  holders: true,
} as const satisfies Record<RecordField<KeptCut>, true>;

const isRevisionId = (value: unknown): boolean =>
  typeof value === "number" &&
  Number.isInteger(value) &&
  value >= 0 &&
  value <= MAX_REVISION_ID;

/** Whether `holders` names a revision on one side of a marker, or both. */
const isKeptCutHolders = (holders: unknown): holders is KeptCutHolders => {
  if (typeof holders !== "object" || holders === null || Array.isArray(holders))
    return false;
  const keys = Object.keys(holders);
  return (
    (Object.hasOwn(holders, "before") || Object.hasOwn(holders, "after")) &&
    keys.every((key) => {
      if (!Object.hasOwn(KEPT_CUT_HOLDER_FIELDS, key)) return false;
      const value = Reflect.get(holders, key);
      return key === "scope" ? value === "seam" : isRevisionId(value);
    })
  );
};

/**
 * Whether `cut` is a well-formed kept cut: exactly its fields, kept from a
 * dated revision, naming the revisions holding its piece if any. The
 * contract and the review history reader share it.
 */
export const isKeptCut = (cut: unknown): cut is KeptCut =>
  typeof cut === "object" &&
  cut !== null &&
  !Array.isArray(cut) &&
  Object.keys(cut).every((key) => Object.hasOwn(KEPT_CUT_FIELDS, key)) &&
  "keptFrom" in cut &&
  typeof cut.keptFrom === "string" &&
  (!("holders" in cut) || isKeptCutHolders(cut.holders));

const hasValidKeptCuts = (nodes: readonly InlineNode[]): boolean =>
  nodes.every((node) => {
    if (carriesKeptCut(node) && node.keptCut !== undefined)
      return isKeptCut(node.keptCut);
    return hasValidKeptCuts(childNodes(node) ?? []);
  });

/**
 * Whether every kept cut in `blocks` is well formed: inline markers only,
 * kept from a dated revision.
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

/** The id of the pending change at a record edge. */
export const changeAtEdge = (
  node: InlineNode | undefined,
  edge: "start" | "end",
): number | undefined => {
  if (node === undefined) return undefined;
  if (isTrackedWrapper(node)) return node.info.id;
  const children = childNodes(node);
  return changeAtEdge(
    edge === "start" ? children?.at(0) : children?.at(-1),
    edge,
  );
};

/** Keep each holder only while that revision is still at the facing edge. */
const withFacingHolders = (
  node: InlineNode,
  facing: { before: KeptCutEdge; after: KeptCutEdge },
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord((marker) => {
      if (!carriesKeptCut(marker)) return marker;
      const cut = marker.keptCut;
      const holders = cut?.holders;
      if (cut === undefined || holders === undefined) return marker;
      const keep = (side: "before" | "after") =>
        facing[side] === KEPT_CUT_EDGES.UNSETTLED ||
        facing[side] === holders[side];
      const before = keep("before") ? holders.before : undefined;
      const after = keep("after") ? holders.after : undefined;
      if (before === holders.before && after === holders.after) return marker;
      const facingHolders = withHolderScope(
        keptCutHolders({ before, after }),
        holders.scope,
      );
      return Object.assign({}, marker, {
        keptCut: cutHeldBy(cut, facingHolders),
      });
    }, known),
  );

/** The sides of a run of kept markers whose records a resolution removes. */
export type KeptRunSides = { before: boolean; after: boolean };

/** Whether the piece a marker names goes with the records on the sides `gone`. */
const pieceGoes = (piece: KeptCutPiece, gone: KeptRunSides): boolean => {
  switch (piece) {
    case "before":
      return gone.before;
    case "after":
      return gone.after;
    case "both":
      return gone.before && gone.after;
    default: {
      const unhandled: never = piece;
      return unhandled;
    }
  }
};

type KeptRunBesideRemovedOptions = {
  gone: KeptRunSides;
  /** What the resolution recorded about each marker, kept on its copies. */
  known: WeakMap<InlineNode, string>;
};

/**
 * `node`, standing in a run of kept markers, once the records beside the
 * run on the sides `gone` names go with a resolution: a piece named on a
 * side that goes is consumed with the change that took the seam, as
 * resolving those records with the revision the markers were kept from
 * leaves them.
 */
export const withKeptCutBesideRemoved = (
  node: InlineNode,
  { gone, known }: KeptRunBesideRemovedOptions,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord((marker) => {
      if (!carriesKeptCut(marker) || marker.keptCut === undefined)
        return marker;
      const piece = pieceOf(marker.keptCut.holders);
      return piece !== undefined && pieceGoes(piece, gone)
        ? Object.assign({}, marker, {
            keptCut: cutHeldBy(marker.keptCut, undefined),
          })
        : marker;
    }, known),
  );

/** What meets a list's edge from outside, as kept cuts at that edge read it. */
export const KEPT_CUT_EDGES: Readonly<{
  readonly NOTHING: "nothing";
  readonly PENDING_MARK: "pendingMark";
  readonly UNSETTLED: "unsettled";
}> = Object.freeze({
  /** Nothing pending: a cut there waits on nothing. */
  NOTHING: "nothing",
  /**
   * A pending change of a paragraph mark: resolving it can join the
   * paragraph with its neighbour, so a cut waits on it, but it took no
   * seam, so it holds no piece a marker names.
   */
  PENDING_MARK: "pendingMark",
  /**
   * Not known yet: the edge of a nested list, which the paragraph's own
   * settlement reads once the joins stand. Its markers keep what they have.
   */
  UNSETTLED: "unsettled",
} as const);

type KeptCutEdge =
  number | (typeof KEPT_CUT_EDGES)[keyof typeof KEPT_CUT_EDGES];

type LegacyKeptCutPiece = (cut: KeptCut) => KeptCutPiece | undefined;

type SettleKeptCutsOptions = {
  /** When reading positional history, bind it to the revisions at the seam. */
  legacyPiece?: LegacyKeptCutPiece;
  /** What the resolution recorded about each marker, kept on its copies. */
  known?: WeakMap<InlineNode, string>;
  /** Told of every record rebuilt around a cleared marker. */
  rebuilt?: (from: InlineNode, to: InlineNode) => void;
  /** What meets the list from outside, at each end. */
  outside?: { start: KeptCutEdge; end: KeptCutEdge };
};

/**
 * `nodes` with the kept cuts that wait on nothing cleared, at every depth: a
 * run of kept markers records its cut only while a pending change stands
 * beside it, across the edges of the records around it, or holds some of
 * them: the change whose resolution can still let the pieces meet. A record
 * holding the run is no such change; resolving it keeps them again.
 * Whichever order resolved the rest, no cut outlives that change.
 */
export const settleKeptCuts = (
  nodes: readonly InlineNode[],
  {
    known = new WeakMap(),
    legacyPiece,
    rebuilt,
    outside = { start: KEPT_CUT_EDGES.NOTHING, end: KEPT_CUT_EDGES.NOTHING },
  }: SettleKeptCutsOptions = {},
): readonly InlineNode[] => {
  let out: InlineNode[] | undefined;
  const edgeOf = (node: InlineNode | undefined, edge: "start" | "end") =>
    changeAtEdge(node, edge) ?? KEPT_CUT_EDGES.NOTHING;
  const pendingBefore = (index: number): KeptCutEdge =>
    index === 0 ? outside.start : edgeOf(nodes[index - 1], "end");
  const pendingAfter = (index: number): KeptCutEdge =>
    index === nodes.length - 1
      ? outside.end
      : edgeOf(nodes[index + 1], "start");
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
              ...(legacyPiece === undefined ? {} : { legacyPiece }),
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
    const before = pendingBefore(start);
    const after = pendingAfter(end - 1);
    const waiting =
      run.some((node) => node.type === "deletion") ||
      before !== KEPT_CUT_EDGES.NOTHING ||
      after !== KEPT_CUT_EDGES.NOTHING;
    for (const [offset, node] of run.entries()) {
      let settled;
      if (legacyPiece === undefined) {
        settled = waiting
          ? withFacingHolders(node, { before, after }, known)
          : withoutKeptCutsRecorded(node, known);
      } else {
        settled = changeStandingMarkers(node, (marker) => {
          if (!carriesKeptCut(marker) || marker.keptCut === undefined)
            return marker;
          const piece = legacyPiece(marker.keptCut);
          if (piece === undefined) return marker;
          const holders = keptCutHolders({
            before:
              piece !== "after" && typeof before === "number"
                ? before
                : undefined,
            after:
              piece !== "before" && typeof after === "number"
                ? after
                : undefined,
          });
          return withKeptCut(
            marker,
            cutHeldBy(
              marker.keptCut,
              withHolderScope(holders, piece === "both" ? "seam" : undefined),
            ),
          );
        });
      }
      if (settled === node) continue;
      out ??= [...nodes];
      out[start + offset] = settled;
    }
    start = end;
  }
  return out ?? nodes;
};

/** A paragraph whose kept cuts settle, with its settled content. */
type SettledParagraph = {
  paragraph: Paragraph;
  content: readonly InlineNode[];
};

/**
 * The paragraphs of `body` whose kept cuts settle, each with its settled
 * content. A pending change of the mark between two paragraphs, past any
 * range boundaries between them, stands at both edges it faces: resolving
 * it can join them (see {@link joinNeighbour}), so the markers at those
 * edges wait on it, as they wait on whatever stands beside them once the
 * paragraphs are joined first. It took no seam: a piece a marker there
 * names goes with the change that took it.
 */
export const settledParagraphKeptCuts = (
  body: DocumentBody,
): SettledParagraph[] => {
  const out: SettledParagraph[] = [];
  for (const { list, index, paragraph } of storyParagraphs(body)) {
    const siblings = blockListAt(body.content, list);
    const markEdge = (pending: boolean): KeptCutEdge =>
      pending ? KEPT_CUT_EDGES.PENDING_MARK : KEPT_CUT_EDGES.NOTHING;
    const content = settleKeptCuts(paragraph.content, {
      outside: {
        start: markEdge(
          joinNeighbour(siblings, { index, side: "before" })?.pPrMark !==
            undefined,
        ),
        end: markEdge(
          paragraph.pPrMark !== undefined &&
            joinNeighbour(siblings, { index, side: "after" }) !== undefined,
        ),
      },
    });
    if (content !== paragraph.content) out.push({ paragraph, content });
  }
  return out;
};
