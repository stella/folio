/** Semantic comment definitions and anchors, with narrow exact-state inverses. */
import { Result } from "better-result";
import {
  MAX_REVISION_ID,
  COMMENT_PART_RELATIONSHIPS,
  type Comment,
  type Document,
  type ParagraphContent,
} from "../model/document";
import { replaceParagraphs, storyParagraphs } from "./blocks";
import { contractViolation } from "./contract";
import type { DocumentEdit } from "./edits";
import { structurallyEqual } from "./equality";
import { IDENTITY_SPACES, idKey, isParaId, identityKeysIn, packageIdentityKeys } from "./ids";
import {
  asParagraphContent,
  childNodes,
  compareGaps,
  defaultInsertionGap,
  isCommentAnchor,
  isEmptyRecord,
  leafSpans,
  mergeAlike,
  rebuildNode,
  type Gap,
  type InlineNode,
} from "./leaves";
import { cloneModel } from "./modelClone";
import { applyFormattingPatch } from "./patch";
import {
  DOCUMENT_OP_REFUSAL_REASONS,
  DocumentOpRefusal,
  type DocumentOpRefusalReason,
} from "./refusal";
import type { ApplyOps } from "./resolve";
import { documentStories, findStoryBody, sameStory, storyBody } from "./stories";
import {
  DOCUMENT_OP_TYPES,
  type CommentOp,
  type CommentState,
  type CreateCommentOp,
  type NewIds,
  type OpStory,
  type RestoreCommentStateOp,
  type TextPosition,
} from "./types";

const fail = (
  op: CommentOp,
  message: string,
  reason: DocumentOpRefusalReason = DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
) => Result.err(new DocumentOpRefusal({ opType: op.type, reason, message }));
const validId = (id: number): boolean => Number.isInteger(id) && id >= 0 && id <= MAX_REVISION_ID;
const anchorId = (node: InlineNode): number | undefined => {
  switch (node.type) {
    case "commentRangeStart":
    case "commentRangeEnd":
    case "commentReference":
      return node.id;
    default:
      return undefined;
  }
};
const anchorsIn = (document: Document) =>
  documentStories(document).flatMap((story) =>
    storyParagraphs(storyBody(document, story)).flatMap(({ paragraph }) =>
      leafSpans(paragraph.content).flatMap((span) =>
        isCommentAnchor(span.node) ? [{ story, blockId: paragraph.paraId ?? "", ...span }] : [],
      ),
    ),
  );

/**
 * Root comments without any owned source anchor. Loaded packages may carry
 * them; they stay preserved as read-only entries rather than refusing the
 * document, and no comment operation may create or target one.
 */
export const orphanRootCommentIds = (document: Document): Set<number> => {
  const comments = document.package.document.comments ?? [];
  const ids = new Set(comments.map(({ id }) => id));
  const anchored = new Set(anchorsIn(document).map(({ node }) => anchorId(node)));
  return new Set(
    comments
      .filter(
        ({ id, parentId }) => (parentId === undefined || !ids.has(parentId)) && !anchored.has(id),
      )
      .map(({ id }) => id),
  );
};

/**
 * Shared loaded-comment activation contract; legacy parentId revision associations stay explicit.
 * `allowedOrphans` names unanchored root comments the caller accepts ("all" for a loaded package).
 */
export const commentDocumentIssue = (
  document: Document,
  allowedOrphans: ReadonlySet<number> | "all" = new Set(),
): string | undefined => {
  const comments = document.package.document.comments ?? [];
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  if (comments.some(({ id }) => !validId(id)) || byId.size !== comments.length)
    return "Comment identities must be unique bounded integers.";
  const anchors = anchorsIn(document);
  if (anchors.some(({ node }) => !byId.has(anchorId(node) ?? -1)))
    return "A comment anchor has no definition.";
  for (const comment of comments) {
    const own = anchors.filter(({ node }) => anchorId(node) === comment.id);
    const starts = own.filter(({ node }) => node.type === "commentRangeStart");
    const ends = own.filter(({ node }) => node.type === "commentRangeEnd");
    const references = own.filter(({ node }) => node.type === "commentReference");
    const start = starts.at(0);
    const end = ends.at(0);
    if (starts.length !== ends.length || starts.length > 1 || references.length > 1)
      return "Comment anchors are unbalanced or duplicated.";
    if (
      start &&
      end &&
      (!sameStory(start.story, end.story) || anchors.indexOf(start) > anchors.indexOf(end))
    )
      return "Comment range boundaries must be ordered in one story.";
    const parent = comment.parentId === undefined ? undefined : byId.get(comment.parentId);
    if (!parent && own.length === 0 && allowedOrphans !== "all" && !allowedOrphans.has(comment.id))
      return "A root or revision-associated comment needs an owned source anchor.";
    if (comment.parentId !== undefined && !validId(comment.parentId))
      return "A comment parent identity is invalid.";
    const visited = new Set([comment.id]);
    let current = parent;
    while (current) {
      if (visited.has(current.id)) return "Comment parent relations cannot cycle.";
      visited.add(current.id);
      current = current.parentId === undefined ? undefined : byId.get(current.parentId);
    }
  }
  return undefined;
};

/** Allocate from the package's annotation domains without a hidden mutable counter. */
export const freshCommentId = (document: Document): Result<number, DocumentOpRefusal> => {
  let maximum = -1;
  for (const comment of document.package.document.comments ?? []) {
    maximum = Math.max(maximum, comment.id, comment.parentId ?? -1);
  }
  for (const { node } of anchorsIn(document)) maximum = Math.max(maximum, anchorId(node) ?? -1);
  for (const key of packageIdentityKeys(document.package)) {
    if (!Object.values(IDENTITY_SPACES).some((space) => key.startsWith(`${space}:`))) continue;
    maximum = Math.max(maximum, Number(key.slice(key.lastIndexOf(":") + 1)));
  }
  if (!Number.isInteger(maximum) || maximum >= MAX_REVISION_ID)
    return Result.err(
      new DocumentOpRefusal({
        opType: DOCUMENT_OP_TYPES.CREATE_COMMENT,
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
        message: "The package annotation id space is exhausted.",
      }),
    );
  return Result.ok(maximum + 1);
};

const propertyPresence = (record: object, key: string): CommentState["listPresence"] => {
  if (!Object.hasOwn(record, key)) return "absent";
  return Reflect.get(record, key) === undefined ? "undefined" : "present";
};

const capture = (
  document: Document,
  ids: ReadonlySet<number>,
  addressed: CommentState["anchors"],
): CommentState => {
  const body = document.package.document;
  return {
    relationshipPresence: propertyPresence(document.package, "relationships"),
    relationships: [...(document.package.relationships ?? [])].flatMap(
      ([key, relationship], index) =>
        isCommentRelationship(relationship.type) ? [{ index, key, relationship }] : [],
    ),
    listPresence: propertyPresence(body, "comments"),
    records: (body.comments ?? []).flatMap((comment, index) =>
      ids.has(comment.id) ? [{ index, comment }] : [],
    ),
    anchors: addressed.flatMap(({ story, blockId }) => {
      const found = findStoryBody(document, story);
      const paragraph =
        found &&
        storyParagraphs(found).find(
          ({ paragraph: value }) => idKey(value.paraId ?? "") === idKey(blockId),
        )?.paragraph;
      return paragraph ? [{ story, blockId, content: paragraph.content }] : [];
    }),
  };
};

const changedAnchors = (before: Document, after: Document): CommentState["anchors"] =>
  documentStories(before).flatMap((story) => {
    const next = new Map(
      storyParagraphs(storyBody(after, story)).map(({ paragraph }) => [
        idKey(paragraph.paraId ?? ""),
        paragraph,
      ]),
    );
    return storyParagraphs(storyBody(before, story)).flatMap(({ paragraph }) =>
      paragraph.content !== next.get(idKey(paragraph.paraId ?? ""))?.content
        ? [{ story, blockId: paragraph.paraId ?? "", content: paragraph.content }]
        : [],
    );
  });

const scaffoldDelta = (before: CommentState, after: CommentState): NewIds => {
  const old = new Set(identityKeysIn(before.anchors));
  const next = new Set(identityKeysIn(after.anchors));
  const changed = [...old]
    .filter((id) => !next.has(id))
    .concat([...next].filter((id) => !old.has(id)));
  const values = (space: string) =>
    changed
      .filter((key) => key.startsWith(`${space}:`))
      .map((key) => Number(key.slice(key.indexOf(":") + 1)));
  return { revision: values(IDENTITY_SPACES.REVISION), control: values(IDENTITY_SPACES.CONTROL) };
};

/** Remove only owned anchors, then merge scaffolding with the shared exact field predicate. */
const ownedSeams = (content: readonly InlineNode[], ids: ReadonlySet<number>): number[] => {
  let references = 0;
  return leafSpans(content).flatMap(({ node, before }) => {
    const id = anchorId(node);
    if (id === undefined || !ids.has(id)) return [];
    const position = before.offset - references;
    if (node.type === "commentReference") references++;
    return [position];
  });
};
type NormalizeAnchorCutsOptions = {
  content: readonly InlineNode[];
  ids: ReadonlySet<number>;
  seams: ReadonlySet<number>;
  start?: number;
};
const normalizedWithoutAnchors = ({
  content,
  ids,
  seams,
  start = 0,
}: NormalizeAnchorCutsOptions): InlineNode[] => {
  let normalized: InlineNode[] = [];
  let position = start;
  for (const node of withoutAnchors(content, ids)) {
    const children = childNodes(node);
    const normalizedNode =
      children === undefined
        ? node
        : rebuildNode(
            node,
            normalizedWithoutAnchors({ content: children, ids, seams, start: position }),
          );
    if (seams.has(position)) normalized = mergeAlike(normalized, [normalizedNode]);
    else normalized.push(normalizedNode);
    position += leafSpans([normalizedNode]).at(-1)?.after.offset ?? 0;
  }
  return normalized;
};

const isCommentRelationship = (type: string): boolean =>
  Object.values(COMMENT_PART_RELATIONSHIPS).some((part) => part.type === type);

/** Keep the same part ids as the packager: numeric rId maximum plus one. */
const reconcileRelationships = (document: Document): Document => {
  const comments = document.package.document.comments ?? [];
  const relationships = new Map(document.package.relationships ?? []);
  let next =
    Math.max(
      0,
      ...[...relationships.keys()].map((id) => Number(/^rId(\d+)$/u.exec(id)?.at(1) ?? 0)),
    ) + 1;
  const extended = comments.some(
    (comment) => comment.parentId !== undefined || comment.done !== undefined,
  );
  for (const [kind, part] of Object.entries(COMMENT_PART_RELATIONSHIPS)) {
    const needed = kind === "comments" ? comments.length > 0 : extended;
    const matching = [...relationships].filter(
      ([, relationship]) => relationship.type === part.type,
    );
    if (needed && matching.length === 0) {
      const id = `rId${next++}`;
      relationships.set(id, { id, type: part.type, target: part.target });
    }
    if (kind === "commentsExtended" && !needed)
      for (const [id] of matching) relationships.delete(id);
  }
  if (structurallyEqual([...relationships], [...(document.package.relationships ?? [])]))
    return document;
  return { ...document, package: { ...document.package, relationships } };
};

const edited = (
  before: Document,
  after: Document,
  op: CommentOp,
  ids: ReadonlySet<number>,
): Result<DocumentEdit, DocumentOpRefusal> => {
  if (op.type !== DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE) after = reconcileRelationships(after);
  const invalid = contractViolation(after);
  if (invalid) return fail(op, invalid.message, invalid.reason);
  // An edit may keep a loaded orphan, never strand another comment.
  const issue = commentDocumentIssue(after, orphanRootCommentIds(before));
  if (issue) return fail(op, issue);
  const anchors = changedAnchors(before, after);
  const prior = capture(before, ids, anchors);
  const next = capture(after, ids, anchors);
  if (structurallyEqual(prior, next))
    return Result.ok({
      document: before,
      inverse: [],
      touched: { modified: [], inserted: [], removed: [] },
    });
  return Result.ok({
    document: after,
    inverse: [
      {
        type: DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE,
        ids: [...ids],
        scaffoldIds: scaffoldDelta(prior, next),
        expected: cloneModel(next),
        state: cloneModel(prior),
      },
    ],
    touched: { modified: anchors.map(({ blockId }) => blockId), inserted: [], removed: [] },
  });
};

const withComments = (
  document: Document,
  comments: Comment[],
  presence: CommentState["listPresence"] = "present",
): Document => {
  const body = { ...document.package.document };
  switch (presence) {
    case "present":
      body.comments = comments;
      break;
    case "undefined":
      Reflect.set(body, "comments", undefined);
      break;
    case "absent":
      delete body.comments;
      break;
    default: {
      const unreachable: never = presence;
      return unreachable;
    }
  }
  return { ...document, package: { ...document.package, document: body } };
};

const withoutAnchors = (
  nodes: readonly InlineNode[],
  ids: ReadonlySet<number>,
): readonly InlineNode[] => {
  let changed = false;
  const kept: InlineNode[] = [];
  for (const node of nodes) {
    const id = anchorId(node);
    if (id !== undefined && ids.has(id)) {
      changed = true;
      continue;
    }
    const children = childNodes(node);
    const filtered = children === undefined ? undefined : withoutAnchors(children, ids);
    if (filtered === children || filtered === undefined) {
      kept.push(node);
      continue;
    }
    changed = true;
    const next = rebuildNode(node, filtered);
    if (!isEmptyRecord(next)) kept.push(next);
  }
  return changed ? kept : nodes;
};

type RestoreAnchorsOptions = {
  document: Document;
  anchors: CommentState["anchors"];
  op: CommentOp;
};
const restoreAnchors = ({
  document,
  anchors,
  op,
}: RestoreAnchorsOptions): Result<Document, DocumentOpRefusal> => {
  let current = document;
  for (const { story, blockId, content } of anchors) {
    const body = findStoryBody(current, story);
    const location =
      body &&
      storyParagraphs(body).find(
        ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(blockId),
      );
    if (!location)
      return fail(op, "An owned comment paragraph is missing.", DOCUMENT_OP_REFUSAL_REASONS.STALE);
    const changed = replaceParagraphs({
      document: current,
      story,
      at: location,
      count: 1,
      replacement: [{ ...location.paragraph, content: cloneModel([...content]) }],
    });
    if (changed.isErr()) return fail(op, changed.error.message, changed.error.reason);
    current = changed.value;
  }
  return Result.ok(current);
};

const restore = (
  document: Document,
  op: RestoreCommentStateOp,
): Result<DocumentEdit, DocumentOpRefusal> => {
  const ids = new Set(op.ids);
  if (
    ids.size !== op.ids.length ||
    op.ids.some((id) => !validId(id)) ||
    !structurallyEqual(capture(document, ids, op.expected.anchors), op.expected)
  )
    return fail(op, "The comment inverse is stale.", DOCUMENT_OP_REFUSAL_REASONS.STALE);
  if (
    op.state.records.some(
      ({ comment, index }) => !ids.has(comment.id) || !Number.isInteger(index) || index < 0,
    ) ||
    new Set(op.state.records.map(({ comment }) => comment.id)).size !== op.state.records.length
  )
    return fail(op, "The restored comment records do not match the owned identities.");
  if (!op.scaffoldIds || !structurallyEqual(scaffoldDelta(op.expected, op.state), op.scaffoldIds))
    return fail(op, "The comment inverse scaffold identities do not match its owned cuts.");
  const allowed = new Set([
    ...(op.scaffoldIds.revision ?? []).map((id) => `${IDENTITY_SPACES.REVISION}:${id}`),
    ...(op.scaffoldIds.control ?? []).map((id) => `${IDENTITY_SPACES.CONTROL}:${id}`),
  ]);
  const identities = (state: CommentState) =>
    identityKeysIn(state.anchors)
      .filter((id) => !allowed.has(id))
      .toSorted();
  if (!structurallyEqual(identities(op.expected), identities(op.state)))
    return fail(op, "The comment inverse changes an unowned identified record.");
  if (
    op.state.anchors.length !== op.expected.anchors.length ||
    op.state.anchors.some((entry, index) => {
      const expected = op.expected.anchors.at(index);
      if (
        !expected ||
        !sameStory(entry.story, expected.story) ||
        idKey(entry.blockId) !== idKey(expected.blockId)
      )
        return true;
      const seams = new Set([
        ...ownedSeams(entry.content, ids),
        ...ownedSeams(expected.content, ids),
      ]);
      const preserved = (content: readonly ParagraphContent[]) =>
        normalizedWithoutAnchors({ content, ids, seams });
      const unowned = (content: readonly ParagraphContent[]) =>
        leafSpans(content)
          .filter(({ node }) => {
            const id = anchorId(node);
            return id !== undefined && !ids.has(id);
          })
          .map(({ node }) => node);
      return (
        !structurallyEqual(preserved(entry.content), preserved(expected.content)) ||
        !structurallyEqual(unowned(entry.content), unowned(expected.content))
      );
    })
  )
    return fail(op, "A comment inverse cannot replace text or unowned anchors.");
  const anchors = restoreAnchors({ document, op, anchors: op.state.anchors });
  if (anchors.isErr()) return anchors;
  const comments = (document.package.document.comments ?? []).filter(({ id }) => !ids.has(id));
  for (const { index, comment } of op.state.records) {
    if (index > comments.length)
      return fail(
        op,
        "A restored comment declaration position is stale.",
        DOCUMENT_OP_REFUSAL_REASONS.STALE,
      );
    comments.splice(index, 0, cloneModel(comment));
  }
  if (op.state.listPresence !== "present" && comments.length > 0)
    return fail(op, "A missing comment list cannot discard unowned definitions.");
  const relationships = [...(anchors.value.package.relationships ?? [])].filter(
    ([, relation]) => !isCommentRelationship(relation.type),
  );
  for (const { index, key, relationship } of op.state.relationships) {
    const existing = op.expected.relationships.find((entry) => entry.key === key);
    if (existing && !structurallyEqual(existing.relationship, relationship))
      return fail(op, "A comment inverse cannot alter an existing relationship payload.");
    if (
      !isCommentRelationship(relationship.type) ||
      relationship.id !== key ||
      index < 0 ||
      index > relationships.length ||
      relationships.some(([id]) => id === key)
    )
      return fail(op, "The comment inverse cannot replace an unowned relationship.");
    relationships.splice(index, 0, [key, cloneModel(relationship)]);
  }
  if (op.state.relationshipPresence !== "present" && relationships.length > 0)
    return fail(op, "A missing relationship map cannot discard unowned relationships.");
  const restored = withComments(anchors.value, comments, op.state.listPresence);
  const pkg = { ...restored.package };
  switch (op.state.relationshipPresence) {
    case "present":
      pkg.relationships = new Map(relationships);
      break;
    case "undefined":
      Reflect.set(pkg, "relationships", undefined);
      break;
    case "absent":
      delete pkg.relationships;
      break;
  }
  return edited(document, { ...restored, package: pkg }, op, ids);
};

const revisionRange = (
  document: Document,
  story: OpStory,
  revisionId: number,
): { from: TextPosition; to: TextPosition } | undefined => {
  const body = findStoryBody(document, story);
  if (!body) return undefined;
  const spans = storyParagraphs(body).flatMap(({ paragraph }) => {
    const leaves = leafSpans(paragraph.content);
    const position = (gap: Gap): TextPosition => ({
      story,
      blockId: paragraph.paraId ?? "",
      offset: gap.offset,
      zeroWidthBefore: gap.zeroWidthBefore,
    });
    if (
      paragraph.propertyChanges?.some(({ info }) => info.id === revisionId) ||
      paragraph.pPrMark?.info.id === revisionId
    ) {
      const start = leaves.at(0)?.before ?? { offset: 0, zeroWidthBefore: 0 };
      const end = leaves.at(-1)?.after ?? start;
      return [{ from: position(start), to: position(end) }];
    }
    return leaves.flatMap((span) => {
      const selected = [span.node, ...span.ancestors].some(
        (node) =>
          ((node.type === "insertion" ||
            node.type === "deletion" ||
            node.type === "moveFrom" ||
            node.type === "moveTo") &&
            node.info.id === revisionId) ||
          (node.type === "run" && node.propertyChanges?.some(({ info }) => info.id === revisionId)),
      );
      return selected ? [{ from: position(span.before), to: position(span.after) }] : [];
    });
  });
  const first = spans.at(0);
  const last = spans.at(-1);
  return first && last ? { from: first.from, to: last.to } : undefined;
};

const remainingIds = (document: Document, ids: NewIds | undefined): NewIds => {
  const used = new Set(packageIdentityKeys(document.package));
  return {
    ...(ids?.revision === undefined
      ? {}
      : { revision: ids.revision.filter((id) => !used.has(`${IDENTITY_SPACES.REVISION}:${id}`)) }),
    ...(ids?.control === undefined
      ? {}
      : { control: ids.control.filter((id) => !used.has(`${IDENTITY_SPACES.CONTROL}:${id}`)) }),
  };
};

const create = (
  document: Document,
  op: CreateCommentOp,
  applyOps: ApplyOps,
): Result<DocumentEdit, DocumentOpRefusal> => {
  const comments = document.package.document.comments ?? [];
  if (
    !validId(op.comment.id) ||
    comments.some(({ id }) => id === op.comment.id) ||
    anchorsIn(document).some(({ node }) => anchorId(node) === op.comment.id) ||
    new Set(packageIdentityKeys(document.package)).has(
      `${IDENTITY_SPACES.REVISION}:${op.comment.id}`,
    )
  )
    return fail(
      op,
      "The fresh comment identity is invalid or already used.",
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
    );
  if (Object.hasOwn(op.comment, "parentId"))
    return fail(op, "Comment parent ownership belongs to the anchor discriminator.");
  if (
    op.comment.content.length === 0 ||
    op.comment.content.some(({ paraId }) => paraId === undefined || !isParaId(paraId))
  )
    return fail(
      op,
      "Comment paragraphs need explicit package paragraph identities.",
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
    );
  let comment: Comment = cloneModel(op.comment);
  let positions: { at: TextPosition; content: ParagraphContent[] }[] = [];
  let range: { from: TextPosition; to: TextPosition } | undefined;
  switch (op.anchor.kind) {
    case "reply": {
      const parentId = op.anchor.parentId;
      if (!comments.some(({ id }) => id === parentId))
        return fail(op, "The parent comment does not exist.");
      comment = { ...comment, parentId: op.anchor.parentId };
      // Loaded replies may carry no body markers of their own; the thread is
      // anchored by its nearest anchored ancestor, as for any unanchored reply.
      const anchors = anchorsIn(document);
      const parentById = new Map(comments.map(({ id, parentId: owner }) => [id, owner]));
      const visited = new Set<number>();
      let anchorOwner: number | undefined = parentId;
      let parentAnchors: typeof anchors = [];
      while (anchorOwner !== undefined && !visited.has(anchorOwner)) {
        visited.add(anchorOwner);
        const owner = anchorOwner;
        parentAnchors = anchors.filter(({ node }) => anchorId(node) === owner);
        if (parentAnchors.length > 0) break;
        const next = parentById.get(owner);
        anchorOwner = typeof next === "number" && parentById.has(next) ? next : undefined;
      }
      if (parentAnchors.length === 0)
        return fail(op, "The parent comment has no representable owned anchors.");
      positions = parentAnchors.toReversed().map(({ story, blockId, node, after }) => {
        switch (node.type) {
          case "commentRangeStart":
          case "commentRangeEnd":
          case "commentReference":
            return {
              at: { story, blockId, offset: after.offset, zeroWidthBefore: after.zeroWidthBefore },
              content: [{ type: node.type, id: comment.id }],
            };
          default:
            return {
              at: { story, blockId, offset: after.offset, zeroWidthBefore: after.zeroWidthBefore },
              content: [],
            };
        }
      });
      break;
    }
    case "revision": {
      const revisionId = op.anchor.revisionId;
      if (!validId(revisionId) || comments.some(({ id }) => id === revisionId))
        return fail(op, "Revision association cannot ambiguously name a comment parent.");
      range = revisionRange(document, op.anchor.story, op.anchor.revisionId);
      if (!range) return fail(op, "The associated revision has no representable source range.");
      comment = { ...comment, parentId: op.anchor.revisionId };
      if (
        range.from.blockId === range.to.blockId &&
        range.from.offset === range.to.offset &&
        range.from.zeroWidthBefore === range.to.zeroWidthBefore
      ) {
        positions = [{ at: range.from, content: [{ type: "commentReference", id: comment.id }] }];
        range = undefined;
      }
      break;
    }
    case "point":
      positions = [{ at: op.anchor.at, content: [{ type: "commentReference", id: comment.id }] }];
      break;
    case "range":
      range = { from: op.anchor.from, to: op.anchor.to };
      break;
    default: {
      const unreachable: never = op.anchor;
      void unreachable;
      return fail(op, "The comment anchor discriminator is invalid.");
    }
  }
  if (range) {
    if (!sameStory(range.from.story, range.to.story))
      return fail(
        op,
        "A comment range must stay in one story.",
        DOCUMENT_OP_REFUSAL_REASONS.CROSS_BLOCK_RANGE,
      );
    const body = findStoryBody(document, range.from.story);
    if (!body)
      return fail(
        op,
        "The comment story does not exist.",
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      );
    const paragraphs = storyParagraphs(body).map(({ paragraph }) => paragraph);
    const fromIndex = paragraphs.findIndex(
      ({ paraId }) => idKey(paraId ?? "") === idKey(range.from.blockId),
    );
    const toIndex = paragraphs.findIndex(
      ({ paraId }) => idKey(paraId ?? "") === idKey(range.to.blockId),
    );
    const fromParagraph = paragraphs.at(fromIndex);
    const toParagraph = paragraphs.at(toIndex);
    if (fromIndex < 0 || toIndex < 0 || !fromParagraph || !toParagraph)
      return fail(
        op,
        "A comment boundary paragraph does not exist.",
        DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      );
    const fromGap =
      range.from.zeroWidthBefore === undefined
        ? defaultInsertionGap(fromParagraph.content, range.from.offset)
        : { offset: range.from.offset, zeroWidthBefore: range.from.zeroWidthBefore };
    const toGap =
      range.to.zeroWidthBefore === undefined
        ? defaultInsertionGap(toParagraph.content, range.to.offset)
        : { offset: range.to.offset, zeroWidthBefore: range.to.zeroWidthBefore };
    if (fromIndex > toIndex || (fromIndex === toIndex && compareGaps(fromGap, toGap) > 0))
      return fail(
        op,
        "A comment range cannot run backwards.",
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
      );
    positions = [
      {
        at: { ...range.to, zeroWidthBefore: toGap.zeroWidthBefore },
        content: [
          { type: "commentRangeEnd", id: comment.id },
          { type: "commentReference", id: comment.id },
        ],
      },
      {
        at: { ...range.from, zeroWidthBefore: fromGap.zeroWidthBefore },
        content: [{ type: "commentRangeStart", id: comment.id }],
      },
    ];
  }
  let current = withComments(document, [...comments, comment]);
  for (const { at, content } of positions) {
    const inserted = applyOps(current, [
      {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at,
        slice: { content, openStart: 0, openEnd: 0 },
        newIds: remainingIds(current, op.newIds),
      },
    ]);
    if (inserted.isErr()) return Result.err(inserted.error);
    current = inserted.value.document;
  }
  return edited(document, current, op, new Set([comment.id]));
};

export const applyCommentOp = (
  document: Document,
  op: CommentOp,
  applyOps: ApplyOps,
): Result<DocumentEdit, DocumentOpRefusal> => {
  if (op.type === DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE) {
    if (
      !Array.isArray(op.ids) ||
      !op.expected ||
      !op.state ||
      !Array.isArray(op.expected.relationships) ||
      !Array.isArray(op.state.relationships) ||
      !Array.isArray(op.expected.records) ||
      !Array.isArray(op.expected.anchors) ||
      !Array.isArray(op.state.records) ||
      !Array.isArray(op.state.anchors)
    )
      return fail(op, "The comment inverse payload is invalid.");
    return restore(document, op);
  }
  const issue = commentDocumentIssue(document, "all");
  if (issue) return fail(op, issue);
  const orphans = orphanRootCommentIds(document);
  if (op.type === DOCUMENT_OP_TYPES.CREATE_COMMENT) {
    if (op.anchor?.kind === "reply" && orphans.has(op.anchor.parentId))
      return fail(op, "A comment without a source anchor is preserved read-only.");
    if (
      !op.comment ||
      typeof op.comment.author !== "string" ||
      !Array.isArray(op.comment.content) ||
      !op.anchor ||
      typeof op.anchor !== "object"
    )
      return fail(op, "The comment creation payload is invalid.");
    return create(document, op, applyOps);
  }
  const comments = document.package.document.comments ?? [];
  const comment = comments.find(({ id }) => id === op.id);
  if (!comment) return fail(op, "The comment does not exist.");
  if (orphans.has(comment.id))
    return fail(op, "A comment without a source anchor is preserved read-only.");
  if (op.type === DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT) {
    if (
      !Array.isArray(op.content) ||
      op.content.length === 0 ||
      op.content.some(({ paraId }) => paraId === undefined || !isParaId(paraId))
    )
      return fail(
        op,
        "Comment content needs explicit paragraph identities.",
        DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
      );
    if (
      op.patch &&
      Object.keys(op.patch).some((key) => key !== "author" && key !== "initials" && key !== "date")
    )
      return fail(op, "A content edit cannot change comment ownership.");
    if (
      op.patch &&
      Object.hasOwn(op.patch, "author") &&
      (op.patch.author === null || op.patch.author === undefined)
    )
      return fail(op, "A comment must retain an author.");
    const updated = {
      ...comment,
      ...applyFormattingPatch(
        {
          author: comment.author,
          ...(Object.hasOwn(comment, "initials") ? { initials: comment.initials } : {}),
          ...(Object.hasOwn(comment, "date") ? { date: comment.date } : {}),
        },
        op.patch ?? {},
      ),
      content: cloneModel([...op.content]),
    };
    if (op.patch?.initials === null) delete updated.initials;
    if (op.patch?.date === null) delete updated.date;
    return edited(
      document,
      withComments(
        document,
        comments.map((value) => (value === comment ? updated : value)),
      ),
      op,
      new Set([comment.id]),
    );
  }
  if (op.type === DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION) {
    if (op.status !== "open" && op.status !== "resolved")
      return fail(op, "Comment resolution status is invalid.");
    return edited(
      document,
      withComments(
        document,
        comments.map((value) =>
          value === comment ? { ...comment, done: op.status === "resolved" } : value,
        ),
      ),
      op,
      new Set([comment.id]),
    );
  }
  const parentIsComment =
    comment.parentId !== undefined && comments.some(({ id }) => id === comment.parentId);
  if ((op.scope !== "thread" && op.scope !== "reply") || (op.scope === "reply") !== parentIsComment)
    return fail(op, "Comment deletion scope must match root or reply ownership.");
  const ids = new Set([comment.id]);
  for (let previousSize = -1; previousSize !== ids.size;) {
    previousSize = ids.size;
    for (const value of comments)
      if (value.parentId !== undefined && ids.has(value.parentId)) ids.add(value.id);
  }
  let current = document;
  for (const story of documentStories(document)) {
    for (const { paragraph } of storyParagraphs(storyBody(document, story))) {
      const seams = new Set(ownedSeams(paragraph.content, ids));
      if (seams.size === 0) continue;
      const content = normalizedWithoutAnchors({ content: paragraph.content, ids, seams });
      if (content === paragraph.content) continue;
      const changed = restoreAnchors({
        document: current,
        op,
        anchors: [{ story, blockId: paragraph.paraId ?? "", content: asParagraphContent(content) }],
      });
      if (changed.isErr()) return changed;
      current = changed.value;
    }
  }
  return edited(
    document,
    withComments(
      current,
      comments.filter(({ id }) => !ids.has(id)),
      comments.every(({ id }) => ids.has(id)) ? "absent" : "present",
    ),
    op,
    ids,
  );
};
