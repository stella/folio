/**
 * One range per comment id: what an insertion does with the comment anchors
 * its slice carries.
 *
 * A slice cut from the document it goes back into repeats the anchors of the
 * comments it covers. Inserted as they are, the comment would own two
 * ranges, which a file cannot say and which resolve differently depending
 * on which of them goes first. So an anchor whose kind its comment already
 * has standing in the document is a copy: when the operation carries that
 * comment's thread and the slice holds the comment's range whole, the
 * anchors take a fresh comment id and the thread is copied under it, as a
 * paste copies the comment; otherwise those anchors are left out. An anchor
 * completing a comment's range, or one of a comment the document does not
 * hold, goes in as it is, except where the slice repeats it: the slice's
 * first range is the comment's.
 */
import { Result } from "better-result";

import type {
  BlockContent,
  Comment,
  Document,
  ParagraphContent,
  PreservedChild,
  PreservedMarkup,
  Table,
  TableRow,
} from "../model/document";
import { MAX_REVISION_ID, preservedChildGap } from "../model/document";
import {
  preservedCommentRangeMarker,
  withPreservedCommentRangeId,
} from "../xml/preservedCommentRangeMarkers";
import { commentAnchorsIn } from "./comments";
import { IDENTITY_SPACES, isParaId, packageIdentityKeys } from "./ids";
import {
  childNodes,
  type InlineNode,
  isCommentAnchor,
  isEmptyRecord,
  isParagraphContent,
  rebuildNode,
} from "./leaves";
import { cloneModel } from "./modelClone";
import { paragraphIdsCollide } from "./paragraphPrimitive";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
import { DOCUMENT_OP_TYPES, type InsertContentOp } from "./types";

type AnchorKind = "commentRangeStart" | "commentRangeEnd" | "commentReference";

const MOVE_BOUNDARY_KINDS = {
  moveFromRangeStart: true,
  moveFromRangeEnd: true,
  moveToRangeStart: true,
  moveToRangeEnd: true,
} as const;

export type MoveBoundaryKind = keyof typeof MOVE_BOUNDARY_KINDS;

export const isMoveBoundaryKind = (kind: string): kind is MoveBoundaryKind =>
  Object.hasOwn(MOVE_BOUNDARY_KINDS, kind);

type MoveBoundary = Extract<InlineNode, { type: MoveBoundaryKind }>;

const isMoveBoundary = (node: InlineNode): node is MoveBoundary =>
  isMoveBoundaryKind(node.type);

/** The slice's content with its comment anchors settled, and the records it adds. */
export type PastedComments = {
  content: ParagraphContent[];
  /** Copied or carried comment records the insertion adds, in carried order. */
  added: Comment[];
};

const refuse = (
  reason: DocumentOpRefusalReason,
  message: string,
): Result<never, DocumentOpRefusal> =>
  Result.err(
    new DocumentOpRefusal({
      reason,
      message,
      opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    }),
  );

const validCommentId = (id: number): boolean =>
  Number.isInteger(id) && id >= 0 && id <= MAX_REVISION_ID;

/**
 * A comment anchor of the slice, and which of the slice's stories it stands
 * in: the slice's own content is 0, and each text box in it one of its own.
 */
type SliceAnchor = { id: number; kind: AnchorKind; story: number };

/** The id an anchor goes on under, or `null` to leave it out. */
type AnchorVisitor = (anchor: SliceAnchor) => number | null;

/** A comment anchor or move range boundary of the slice. */
type SliceMarker =
  SliceAnchor | { id: number; kind: MoveBoundaryKind; story: number };

type MarkerVisitor = (marker: SliceMarker) => number | null;

type SliceWalk = { visit: MarkerVisitor; stories: number };

type StoryWalk = { walk: SliceWalk; story: number };

/** Whether a rewrite kept every record it was given. */
const unchanged = <T>(before: readonly T[], after: readonly T[]): boolean =>
  after.length === before.length &&
  after.every((record, index) => record === before[index]);

/**
 * The slice with `visit` deciding the id of each comment anchor and move range
 * boundary, in leaf order and into the text boxes its shapes hold; records
 * left empty go too. Reading the markers and settling them are this one walk,
 * so they agree on every marker.
 */
export const rewriteSliceMarkers = (
  nodes: readonly ParagraphContent[],
  visit: MarkerVisitor,
): ParagraphContent[] =>
  rewriteInline(nodes, {
    walk: { visit, stories: 0 },
    story: 0,
  }).filter(isParagraphContent);

/** {@link rewriteSliceMarkers} of the comment anchors alone. */
const rewriteAnchors = (
  nodes: readonly ParagraphContent[],
  visit: AnchorVisitor,
): ParagraphContent[] =>
  rewriteSliceMarkers(nodes, (marker) =>
    isMoveBoundaryKind(marker.kind)
      ? marker.id
      : visit({ ...marker, kind: marker.kind }),
  );

const rewriteInline = (
  nodes: readonly InlineNode[],
  at: StoryWalk,
): InlineNode[] => {
  const out: InlineNode[] = [];
  for (const node of nodes) {
    if (isCommentAnchor(node) || isMoveBoundary(node)) {
      const id = at.walk.visit({
        id: node.id,
        kind: node.type,
        story: at.story,
      });
      if (id !== null) out.push(id === node.id ? node : { ...node, id });
      continue;
    }
    if (node.type === "shape" && node.shape.textBody !== undefined) {
      const { textBody } = node.shape;
      at.walk.stories += 1;
      const content = rewriteBlocks(textBody.content, {
        walk: at.walk,
        story: at.walk.stories,
      });
      out.push(
        unchanged(textBody.content, content)
          ? node
          : {
              ...node,
              shape: { ...node.shape, textBody: { ...textBody, content } },
            },
      );
      continue;
    }
    const children = childNodes(node);
    if (children === undefined) {
      out.push(node);
      continue;
    }
    const next = rewriteInline(children, at);
    if (unchanged(children, next)) {
      out.push(node);
      continue;
    }
    const rebuilt = rebuildNode(node, next);
    // A record holding only left-out anchors holds nothing now.
    if (!isEmptyRecord(rebuilt)) out.push(rebuilt);
  }
  return out;
};

const rewriteBlocks = (
  blocks: readonly BlockContent[],
  at: StoryWalk,
): BlockContent[] =>
  blocks.flatMap((block): BlockContent[] => {
    switch (block.type) {
      case "paragraph": {
        const content = rewriteInline(block.content, at);
        return [
          unchanged<InlineNode>(block.content, content)
            ? block
            : { ...block, content: content.filter(isParagraphContent) },
        ];
      }
      case "table":
        return [rewriteTable(block, at)];
      case "blockSdt":
      case "blockCustomXml": {
        const content = rewriteBlocks(block.content, at);
        return [
          unchanged(block.content, content) ? block : { ...block, content },
        ];
      }
      case "commentRangeStart":
      case "commentRangeEnd":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd": {
        const id = at.walk.visit({
          id: block.id,
          kind: block.type,
          story: at.story,
        });
        if (id === null) return [];
        return [id === block.id ? block : { ...block, id }];
      }
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        return [block];
      default: {
        const unhandled: never = block;
        return unhandled;
      }
    }
  });

type PositionedRewrite = {
  /** Settle the boundaries standing at a gap between rows or cells. */
  settleGap: (gap: number) => void;
  /** The markup once every gap is settled: the same one when nothing changed. */
  markup: () => PreservedMarkup | undefined;
};

/** Range boundaries a table or row keeps as positioned markup, gap by gap. */
const positionedRewrite = (
  preserved: PreservedMarkup | undefined,
  { childCount, at }: { childCount: number; at: StoryWalk },
): PositionedRewrite => {
  const settled = new Map<PreservedChild, PreservedChild | null>();
  return {
    settleGap: (gap) => {
      for (const child of preserved?.children ?? []) {
        if (preservedChildGap(child, childCount) !== gap) continue;
        const marker = preservedCommentRangeMarker(child);
        if (marker === undefined) continue;
        const id = at.walk.visit({
          id: marker.id,
          kind: marker.type,
          story: at.story,
        });
        settled.set(
          child,
          id === null ? null : withPreservedCommentRangeId(child, id),
        );
      }
    },
    markup: () => {
      const before = preserved?.children ?? [];
      const children = before.flatMap((child) => {
        const next = settled.get(child);
        return next === null ? [] : [next ?? child];
      });
      if (unchanged(before, children)) return preserved;
      if (children.length > 0) return { ...preserved, children };
      // An empty record is never written.
      const { children: _left, ...rest } = preserved ?? {};
      return Reflect.ownKeys(rest).length === 0 ? undefined : rest;
    },
  };
};

const rewriteTable = (table: Table, at: StoryWalk): Table => {
  const between = positionedRewrite(table.preserved, {
    childCount: table.rows.length,
    at,
  });
  const rows = table.rows.map((row, rowIndex): TableRow => {
    between.settleGap(rowIndex);
    const inRow = positionedRewrite(row.preserved, {
      childCount: row.cells.length,
      at,
    });
    const cells = row.cells.map((cell, cellIndex) => {
      inRow.settleGap(cellIndex);
      const content = rewriteBlocks(cell.content, at);
      return unchanged(cell.content, content) ? cell : { ...cell, content };
    });
    inRow.settleGap(row.cells.length);
    const markup = inRow.markup();
    if (markup === row.preserved && unchanged(row.cells, cells)) return row;
    const { preserved: _row, ...rest } = row;
    return markup === undefined
      ? { ...rest, cells }
      : { ...rest, cells, preserved: markup };
  });
  between.settleGap(table.rows.length);
  const markup = between.markup();
  if (markup === table.preserved && unchanged(table.rows, rows)) return table;
  const { preserved: _table, ...rest } = table;
  return markup === undefined
    ? { ...rest, rows }
    : { ...rest, rows, preserved: markup };
};

/** Each comment id's anchors in the slice, in leaf order. */
const sliceAnchorsIn = (
  content: readonly ParagraphContent[],
): Map<number, SliceAnchor[]> => {
  const out = new Map<number, SliceAnchor[]>();
  rewriteAnchors(content, (anchor) => {
    const own = out.get(anchor.id);
    if (own === undefined) out.set(anchor.id, [anchor]);
    else own.push(anchor);
    return anchor.id;
  });
  return out;
};

type HoldsWholeOptions = {
  /** The comment's anchors in the slice, in leaf order. */
  anchors: readonly SliceAnchor[] | undefined;
  /** Whether the comment is on a point, its reference the only anchor. */
  point: boolean;
};

/** Whether a comment's anchors in the slice make its one range, or its point. */
const holdsWhole = ({ anchors, point }: HoldsWholeOptions): boolean => {
  if (anchors === undefined) return true;
  const of = (kind: AnchorKind) =>
    anchors.filter((anchor) => anchor.kind === kind);
  const [start, ...starts] = of("commentRangeStart");
  const [end, ...ends] = of("commentRangeEnd");
  const references = of("commentReference").length;
  if (start === undefined && end === undefined)
    return point && references === 1;
  return (
    start !== undefined &&
    end !== undefined &&
    starts.length === 0 &&
    ends.length === 0 &&
    references <= 1 &&
    start.story === end.story &&
    anchors.indexOf(start) < anchors.indexOf(end)
  );
};

/**
 * Which of a comment's anchors in the slice go in as they are: one of each
 * kind it has none of standing, and an end after its start, in the story its
 * start stands in, where there is one.
 */
const keptAnchors = (
  anchors: readonly SliceAnchor[],
  standing: ReadonlySet<string> | undefined,
): boolean[] => {
  const first = (
    kind: AnchorKind,
    matches: (anchor: SliceAnchor, index: number) => boolean = () => true,
  ): number =>
    standing?.has(kind) === true
      ? -1
      : anchors.findIndex(
          (anchor, index) => anchor.kind === kind && matches(anchor, index),
        );
  const start = first("commentRangeStart");
  const following = first(
    "commentRangeEnd",
    ({ story }, index) =>
      index > start && (start < 0 || story === anchors[start]?.story),
  );
  const end = following < 0 ? first("commentRangeEnd") : following;
  const kept = new Set([start, end, first("commentReference")]);
  return anchors.map((_, index) => kept.has(index));
};

/** A carried thread: its root and the replies under it, in carried order. */
type CarriedThread = { root: number; members: Comment[] };

const carriedThreads = (
  carried: readonly Comment[],
): Result<CarriedThread[], DocumentOpRefusal> => {
  const byId = new Map<number, Comment>();
  for (const comment of carried) {
    if (!validCommentId(comment.id) || byId.has(comment.id))
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
        "Carried comment ids must be unique bounded integers.",
      );
    byId.set(comment.id, comment);
  }
  const rootOf = (comment: Comment): number | undefined => {
    const seen = new Set<number>();
    let current = comment;
    while (current.parentId !== undefined) {
      if (seen.has(current.id)) return undefined;
      seen.add(current.id);
      const parent = byId.get(current.parentId);
      if (parent === undefined) return undefined;
      current = parent;
    }
    return current.id;
  };
  const threads = new Map<number, CarriedThread>();
  for (const comment of carried) {
    const root = rootOf(comment);
    if (root === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "A carried reply must reply to a carried comment.",
      );
    const thread = threads.get(root) ?? { root, members: [] };
    thread.members.push(comment);
    threads.set(root, thread);
  }
  return Result.ok([...threads.values()]);
};

type SettleAnchorsOptions = {
  content: readonly ParagraphContent[];
  renamed: ReadonlyMap<number, number>;
  /** Per comment id, which of its anchors stay, in leaf order; all when absent. */
  kept: ReadonlyMap<number, readonly boolean[]>;
};

/** Rewrite the slice's anchors: renamed ids, left-out ones. */
const settleAnchors = ({
  content,
  renamed,
  kept,
}: SettleAnchorsOptions): ParagraphContent[] => {
  // How many of each comment's anchors came before.
  const passed = new Map<number, number>();
  return rewriteAnchors(content, ({ id }) => {
    const index = passed.get(id) ?? 0;
    passed.set(id, index + 1);
    if (kept.get(id)?.[index] === false) return null;
    return renamed.get(id) ?? id;
  });
};

/** The ids a package already gives meaning to: comments, their anchors, revisions. */
const takenCommentIds = (document: Document): Set<number> => {
  const taken = new Set<number>();
  for (const comment of document.package.document.comments ?? []) {
    taken.add(comment.id);
    if (comment.parentId !== undefined) taken.add(comment.parentId);
  }
  for (const { node } of commentAnchorsIn(document))
    if (isCommentAnchor(node)) taken.add(node.id);
  const revision = `${IDENTITY_SPACES.REVISION}:`;
  for (const key of packageIdentityKeys(document.package))
    if (key.startsWith(revision)) taken.add(Number(key.slice(revision.length)));
  return taken;
};

/** Each comment id's anchor kinds standing in the document. */
const standingAnchorKinds = (document: Document): Map<number, Set<string>> => {
  const standing = new Map<number, Set<string>>();
  for (const { node } of commentAnchorsIn(document)) {
    if (!isCommentAnchor(node)) continue;
    const kinds = standing.get(node.id) ?? new Set<string>();
    kinds.add(node.type);
    standing.set(node.id, kinds);
  }
  return standing;
};

/**
 * Whether `content` holds a comment anchor whose kind its comment has
 * standing in `document`, or repeats one, so that inserting it there settles
 * the anchor instead of putting it back as it was.
 */
export const repeatsCommentAnchor = (
  document: Document,
  content: readonly ParagraphContent[],
): boolean => {
  const sliceAnchors = sliceAnchorsIn(content);
  if (sliceAnchors.size === 0) return false;
  const standing = standingAnchorKinds(document);
  for (const [id, anchors] of sliceAnchors)
    if (keptAnchors(anchors, standing.get(id)).includes(false)) return true;
  return false;
};

/**
 * Settle the comment anchors `content` (the slice of `op`) carries against
 * the document it goes into; see the module comment.
 */
export const pastedComments = (
  document: Document,
  op: InsertContentOp,
  content: readonly ParagraphContent[],
): Result<PastedComments, DocumentOpRefusal> => {
  const sliceAnchors = sliceAnchorsIn(content);
  const carried = op.comments ?? [];
  if (sliceAnchors.size === 0 && carried.length === 0)
    return Result.ok({ content: [...content], added: [] });
  const threads = carriedThreads(carried);
  if (threads.isErr()) return Result.err(threads.error);
  const standing = standingAnchorKinds(document);
  const duplicated = (id: number): SliceAnchor[] =>
    (sliceAnchors.get(id) ?? []).filter(
      ({ kind }) => standing.get(id)?.has(kind) === true,
    );
  const records = new Set(
    (document.package.document.comments ?? []).map(({ id }) => id),
  );
  const renamed = new Map<number, number>();
  const added: Comment[] = [];
  const fresh = [...(op.newCommentIds ?? [])];
  const taken = takenCommentIds(document);
  for (const { id } of carried) taken.add(id);
  for (const id of sliceAnchors.keys()) taken.add(id);
  const copied = new Set<Comment>();
  const joining = new Set<Comment>();
  for (const { root, members } of threads.value) {
    if (!sliceAnchors.has(root))
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        "A carried comment names no anchor of the slice.",
      );
    const copies = members.some(({ id }) => duplicated(id).length > 0);
    if (!copies) {
      // Back where it was cut from, or new to the package: the anchors keep
      // their ids, and the thread joins the records unless they hold it.
      if (records.has(root)) continue;
      if (members.some(({ id }) => records.has(id)))
        return refuse(
          DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
          "A carried reply's id is already used.",
        );
      for (const comment of members) joining.add(comment);
      continue;
    }
    // A copy of part of a range would be a range of its own half.
    if (
      members.every(({ id, anchorKind }) =>
        holdsWhole({
          anchors: sliceAnchors.get(id),
          point:
            anchorKind !== "range" &&
            standing.get(id)?.has("commentRangeStart") !== true &&
            standing.get(id)?.has("commentRangeEnd") !== true,
        }),
      )
    )
      for (const comment of members) copied.add(comment);
  }
  // Threads may be carried interleaved: the fresh ids go to the copied
  // comments in carried order, which the records keep.
  const ids = new Map<number, number>();
  for (const comment of carried) {
    if (!copied.has(comment)) continue;
    const next = fresh.shift();
    if (next === undefined)
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        "A pasted copy of a standing comment needs a fresh comment id.",
      );
    if (!validCommentId(next) || taken.has(next))
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        "A fresh comment id is invalid or already used.",
      );
    taken.add(next);
    ids.set(comment.id, next);
    if (sliceAnchors.has(comment.id)) renamed.set(comment.id, next);
  }
  const suggestionId = op.revision?.suggestionId;
  for (const comment of carried) {
    if (!copied.has(comment) && !joining.has(comment)) continue;
    const record = cloneModel(comment);
    record.id = ids.get(comment.id) ?? comment.id;
    if (comment.parentId !== undefined)
      record.parentId = ids.get(comment.parentId) ?? comment.parentId;
    // What proposed the comment elsewhere does not propose it here.
    if (suggestionId === undefined) delete record.suggestionId;
    else record.suggestionId = suggestionId;
    added.push(record);
  }
  const kept = new Map<number, readonly boolean[]>();
  for (const [id, anchors] of sliceAnchors) {
    if (renamed.has(id)) continue;
    const keeps = keptAnchors(anchors, standing.get(id));
    if (keeps.includes(false)) kept.set(id, keeps);
  }
  for (const comment of added) {
    if (
      comment.content.length === 0 ||
      comment.content.some(
        ({ paraId }) => paraId === undefined || !isParaId(paraId),
      )
    )
      return refuse(
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        "Carried comment paragraphs need explicit package paragraph identities.",
      );
  }
  const incoming = added.flatMap(({ content: body }) =>
    body.flatMap(({ paraId }) => (paraId === undefined ? [] : [paraId])),
  );
  if (paragraphIdsCollide({ document, incoming }))
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
      "A carried comment holds a paragraph id the package already uses.",
    );
  const settled =
    renamed.size === 0 && kept.size === 0
      ? [...content]
      : settleAnchors({ content, renamed, kept });
  return Result.ok({ content: settled, added });
};
