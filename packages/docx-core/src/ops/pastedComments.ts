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

import type { Comment, Document, ParagraphContent } from "../model/document";
import { MAX_REVISION_ID } from "../model/document";
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

/** Each comment id's anchor kinds, in leaf order. */
const anchorKindsIn = (
  nodes: readonly InlineNode[],
  out = new Map<number, AnchorKind[]>(),
): Map<number, AnchorKind[]> => {
  for (const node of nodes) {
    if (isCommentAnchor(node)) {
      const kinds = out.get(node.id) ?? [];
      kinds.push(node.type);
      out.set(node.id, kinds);
      continue;
    }
    anchorKindsIn(childNodes(node) ?? [], out);
  }
  return out;
};

type HoldsWholeOptions = {
  /** The comment's anchors in the slice, in leaf order. */
  kinds: readonly AnchorKind[] | undefined;
  /** Whether the comment is on a point, its reference the only anchor. */
  point: boolean;
};

/** Whether a comment's anchors in the slice make its one range, or its point. */
const holdsWhole = ({ kinds, point }: HoldsWholeOptions): boolean => {
  if (kinds === undefined) return true;
  const starts = kinds.filter((kind) => kind === "commentRangeStart").length;
  const ends = kinds.filter((kind) => kind === "commentRangeEnd").length;
  const references = kinds.filter((kind) => kind === "commentReference").length;
  if (starts === 0 && ends === 0) return point && references === 1;
  return (
    starts === 1 &&
    ends === 1 &&
    references <= 1 &&
    kinds.indexOf("commentRangeStart") < kinds.indexOf("commentRangeEnd")
  );
};

/**
 * Which of a comment's anchors in the slice go in as they are: one of each
 * kind it has none of standing, and an end after its start where there is one.
 */
const keptAnchors = (
  kinds: readonly AnchorKind[],
  standing: ReadonlySet<string> | undefined,
): boolean[] => {
  const first = (kind: AnchorKind, from = 0): number =>
    standing?.has(kind) === true ? -1 : kinds.indexOf(kind, from);
  const start = first("commentRangeStart");
  const following = first("commentRangeEnd", start + 1);
  const end = following < 0 ? first("commentRangeEnd") : following;
  const kept = new Set([start, end, first("commentReference")]);
  return kinds.map((_, index) => kept.has(index));
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
  nodes: readonly InlineNode[];
  renamed: ReadonlyMap<number, number>;
  /** Per comment id, which of its anchors stay, in leaf order; all when absent. */
  kept: ReadonlyMap<number, readonly boolean[]>;
  /** How many of each comment's anchors came before `nodes`. */
  passed?: Map<number, number>;
};

/** Rewrite the slice's anchors: renamed ids, left-out ones; records they empty go too. */
const settleAnchors = ({
  nodes,
  renamed,
  kept,
  passed = new Map(),
}: SettleAnchorsOptions): InlineNode[] => {
  const out: InlineNode[] = [];
  for (const node of nodes) {
    if (isCommentAnchor(node)) {
      const index = passed.get(node.id) ?? 0;
      passed.set(node.id, index + 1);
      if (kept.get(node.id)?.[index] === false) continue;
      const id = renamed.get(node.id);
      out.push(id === undefined ? node : { ...node, id });
      continue;
    }
    const children = childNodes(node);
    if (children === undefined) {
      out.push(node);
      continue;
    }
    const next = settleAnchors({ nodes: children, renamed, kept, passed });
    if (
      next.length === children.length &&
      next.every((child, index) => child === children[index])
    ) {
      out.push(node);
      continue;
    }
    const rebuilt = rebuildNode(node, next);
    // A record holding only left-out anchors holds nothing now.
    if (!isEmptyRecord(rebuilt)) out.push(rebuilt);
  }
  return out;
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
  const sliceKinds = anchorKindsIn(content);
  if (sliceKinds.size === 0) return false;
  const standing = standingAnchorKinds(document);
  for (const [id, kinds] of sliceKinds)
    if (keptAnchors(kinds, standing.get(id)).includes(false)) return true;
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
  const sliceKinds = anchorKindsIn(content);
  const carried = op.comments ?? [];
  if (sliceKinds.size === 0 && carried.length === 0)
    return Result.ok({ content: [...content], added: [] });
  const threads = carriedThreads(carried);
  if (threads.isErr()) return Result.err(threads.error);
  const standing = standingAnchorKinds(document);
  const duplicated = (id: number): AnchorKind[] =>
    (sliceKinds.get(id) ?? []).filter(
      (kind) => standing.get(id)?.has(kind) === true,
    );
  const records = new Set(
    (document.package.document.comments ?? []).map(({ id }) => id),
  );
  const renamed = new Map<number, number>();
  const added: Comment[] = [];
  const fresh = [...(op.newCommentIds ?? [])];
  const taken = takenCommentIds(document);
  for (const { id } of carried) taken.add(id);
  for (const id of sliceKinds.keys()) taken.add(id);
  const copied = new Set<Comment>();
  const joining = new Set<Comment>();
  for (const { root, members } of threads.value) {
    if (!sliceKinds.has(root))
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
          kinds: sliceKinds.get(id),
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
    if (sliceKinds.has(comment.id)) renamed.set(comment.id, next);
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
  for (const [id, kinds] of sliceKinds) {
    if (renamed.has(id)) continue;
    const keeps = keptAnchors(kinds, standing.get(id));
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
      : settleAnchors({ nodes: content, renamed, kept }).filter(
          isParagraphContent,
        );
  return Result.ok({ content: settled, added });
};
