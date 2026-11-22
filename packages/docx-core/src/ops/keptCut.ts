/**
 * The cut a kept range marker stands in (`KeptCut`): review state a
 * resolution records on a marker it kept from removed content and could not
 * place inside a piece of the record that content cut. It is never written.
 */

import type {
  BlockContent,
  DocumentBody,
  KeptCut,
  KeptCutPiece,
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
 * `node` marked as standing in `cut`, when it is a range marker. A marker
 * already standing in a cut keeps the revision it was first kept from and
 * the piece first named for it: whichever order kept it and then floated it
 * with another removal's seam, it records the same cut.
 */
const withKeptCut = (node: InlineNode, cut: KeptCut): InlineNode => {
  if (!carriesKeptCut(node)) return node;
  const known = node.keptCut;
  if (known === undefined) return Object.assign({}, node, { keptCut: cut });
  if (known.piece !== undefined || cut.piece === undefined) return node;
  return Object.assign({}, node, {
    keptCut: { ...known, piece: cut.piece },
  });
};

/**
 * `node` with every marked range marker it stands for naming `piece` as the
 * side whose record is a piece of its cut, keeping what the resolution
 * recorded per marker. An unmarked marker stays unmarked, and one already
 * naming a piece keeps it.
 */
export const withKeptCutPieceRecorded = (
  node: InlineNode,
  piece: KeptCutPiece,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord(
      (marker) =>
        carriesKeptCut(marker) && marker.keptCut !== undefined
          ? withKeptCut(marker, { ...marker.keptCut, piece })
          : marker,
      known,
    ),
  );

/** `node` with every marker it stands for naming no side as its cut's piece. */
export const withKeptCutPieceConsumed = (
  node: InlineNode,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord((marker) => {
      if (!carriesKeptCut(marker) || marker.keptCut?.piece === undefined)
        return marker;
      const { piece: _consumed, ...cut } = marker.keptCut;
      return Object.assign({}, marker, { keptCut: cut });
    }, known),
  );

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
  return carriesKeptCut(node) ? node.keptCut?.piece : undefined;
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
    piece: KeptCutPiece,
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
      out[index] = withKeptCuts(node, { keptFrom, piece });
    }
  };
  for (const [index, node] of nodes.entries()) {
    const joins = isTrackedWrapper(node) ? node.resolutionJoins : undefined;
    const kept =
      joins === undefined ? undefined : writtenOnChange(joins)?.change;
    if (joins === undefined || kept === undefined) continue;
    const count = kept.keptCount ?? Number.POSITIVE_INFINITY;
    const named = { keptFrom: kept.keptFrom, count };
    if (joins.before > 0) mark(index - 1, -1, named, "after");
    if (joins.after > 0) mark(index + 1, 1, named, "before");
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
    if (group !== undefined) {
      let piece: KeptCutPiece = "before";
      if (across !== undefined) piece = "both";
      else if (after.includes(group)) piece = "after";
      const cut = {
        keptFrom: group.keptFrom,
        count: group.keptCount ?? Number.POSITIVE_INFINITY,
      };
      if (piece === "after") mark(end - 1, -1, cut, piece);
      else mark(start, 1, cut, piece);
    }
    start = end;
  }
  return out.every((node, index) => node === nodes[index]) ? nodes : out;
};

/** {@link withKeptCutsFromChanges} over a paragraph's content. */
export const paragraphKeptCutsFromChanges = (
  content: readonly ParagraphContent[],
  writtenOnChange: (joins: SeamJoins) => VersionOneKept | undefined,
): ParagraphContent[] =>
  asParagraphContent(withKeptCutsFromChanges(content, writtenOnChange));

/** How many range boundaries a kept marker `node` stands for. */
const keptBoundaryCount = (node: InlineNode): number =>
  node.type === "deletion"
    ? node.content.reduce((sum, child) => sum + keptBoundaryCount(child), 0)
    : 1;

const KEPT_CUT_PIECES = {
  before: true,
  after: true,
  both: true,
} as const satisfies Record<KeptCutPiece, true>;

const KEPT_CUT_FIELDS = {
  keptFrom: true,
  piece: true,
} as const satisfies Record<RecordField<KeptCut>, true>;

/**
 * Whether `cut` is a well-formed kept cut: exactly its fields, kept from a
 * dated revision, naming one side as its piece if any. The contract and the
 * review history reader share it.
 */
export const isKeptCut = (cut: unknown): cut is KeptCut =>
  typeof cut === "object" &&
  cut !== null &&
  !Array.isArray(cut) &&
  Object.keys(cut).every((key) => Object.hasOwn(KEPT_CUT_FIELDS, key)) &&
  "keptFrom" in cut &&
  typeof cut.keptFrom === "string" &&
  (!("piece" in cut) ||
    (typeof cut.piece === "string" &&
      Object.hasOwn(KEPT_CUT_PIECES, cut.piece)));

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

/**
 * `node` with every marker it stands for naming a side as its cut's piece
 * keeping that name only where `facing` says the change that took its seam
 * still stands, keeping what the resolution recorded per marker.
 */
const withKeptCutPiecesFacing = (
  node: InlineNode,
  facing: Readonly<Record<KeptCutPiece, boolean>>,
  known: WeakMap<InlineNode, string>,
): InlineNode =>
  changeStandingMarkers(
    node,
    keepingRecord((marker) => {
      if (!carriesKeptCut(marker) || marker.keptCut?.piece === undefined)
        return marker;
      if (facing[marker.keptCut.piece]) return marker;
      const { piece: _consumed, ...cut } = marker.keptCut;
      return Object.assign({}, marker, { keptCut: cut });
    }, known),
  );

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
      if (!carriesKeptCut(marker) || marker.keptCut?.piece === undefined)
        return marker;
      const { piece, ...cut } = marker.keptCut;
      return pieceGoes(piece, gone)
        ? Object.assign({}, marker, { keptCut: cut })
        : marker;
    }, known),
  );

/** What meets a list's edge from outside, as kept cuts at that edge read it. */
export const KEPT_CUT_EDGES: Readonly<{
  readonly NOTHING: "nothing";
  readonly PENDING_MARK: "pendingMark";
  readonly PENDING_CHANGE: "pendingChange";
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
  /** A pending tracked change, which can hold a piece a marker names. */
  PENDING_CHANGE: "pendingChange",
  /**
   * Not known yet: the edge of a nested list, which the paragraph's own
   * settlement reads once the joins stand. Its markers keep what they have.
   */
  UNSETTLED: "unsettled",
} as const);

type KeptCutEdge = (typeof KEPT_CUT_EDGES)[keyof typeof KEPT_CUT_EDGES];

/** Whether a cut at an edge waits there, and whether a piece named there stands. */
const EDGE_READING = {
  nothing: { waits: false, holdsPiece: false },
  pendingMark: { waits: true, holdsPiece: false },
  pendingChange: { waits: true, holdsPiece: true },
  unsettled: { waits: true, holdsPiece: true },
} as const satisfies Record<
  KeptCutEdge,
  { waits: boolean; holdsPiece: boolean }
>;

type SettleKeptCutsOptions = {
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
    rebuilt,
    outside = { start: KEPT_CUT_EDGES.NOTHING, end: KEPT_CUT_EDGES.NOTHING },
  }: SettleKeptCutsOptions = {},
): readonly InlineNode[] => {
  let out: InlineNode[] | undefined;
  const edgeOf = (node: InlineNode | undefined, edge: "start" | "end") =>
    pendingAtEdge(node, edge)
      ? KEPT_CUT_EDGES.PENDING_CHANGE
      : KEPT_CUT_EDGES.NOTHING;
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
    const before = EDGE_READING[pendingBefore(start)];
    const after = EDGE_READING[pendingAfter(end - 1)];
    const waiting =
      run.some((node) => node.type === "deletion") ||
      before.waits ||
      after.waits;
    // A marker names a side as its cut's piece only while the pending change
    // that took its seam stands there: resolving that change consumes it.
    const facing = {
      before: before.holdsPiece,
      after: after.holdsPiece,
      both: before.holdsPiece || after.holdsPiece,
    } as const satisfies Record<KeptCutPiece, boolean>;
    for (const [offset, node] of run.entries()) {
      const settled = waiting
        ? withKeptCutPiecesFacing(node, facing, known)
        : withoutKeptCutsRecorded(node, known);
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
