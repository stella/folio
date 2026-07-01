/**
 * Reply-range marker synthesis for serialization.
 *
 * Word anchors every comment — including a REPLY — with
 * `commentRangeStart` / `commentRangeEnd` markers plus a
 * `commentReference` run in `document.xml`. A reply shares its parent
 * thread's text, so a freshly created reply (see {@link replyToComment})
 * has no markers of its own yet. Before serialization we walk the
 * comment-bearing surfaces and, for each reply that lacks its own
 * markers, duplicate the parent's range markers next to the parent's so
 * the reply gets an identical anchor. The paragraph serializer then
 * emits the reply's `commentReference` run from its `commentRangeEnd`.
 *
 * This is the save-time counterpart to `normalizeCommentReferences`
 * (parse time): it produces BALANCED ranges referencing valid comment
 * ids, so a later parse+normalize leaves them intact rather than
 * re-anchoring them. Idempotent — a reply that already has a marker
 * (parsed from a Word thread, or synthesized on a prior save) is skipped.
 *
 * Approach adapted from `injectReplyRangeMarkers` in
 * eigenpal/docx-editor (Apache-2.0).
 */

import type {
  BlockContent,
  Comment,
  Document,
  Endnote,
  Footnote,
  HeaderFooter,
  Paragraph,
} from "../types/document";

type ReplyThreadSurfaces = {
  body: BlockContent[];
  comments: readonly Comment[];
  headers?: Map<string, HeaderFooter>;
  footers?: Map<string, HeaderFooter>;
  footnotes?: readonly Footnote[];
  endnotes?: readonly Endnote[];
};

const surfacesFromDocument = (doc: Document): ReplyThreadSurfaces => ({
  body: doc.package.document.content,
  comments: doc.package.document.comments ?? [],
  ...(doc.package.headers !== undefined ? { headers: doc.package.headers } : {}),
  ...(doc.package.footers !== undefined ? { footers: doc.package.footers } : {}),
  ...(doc.package.footnotes !== undefined ? { footnotes: doc.package.footnotes } : {}),
  ...(doc.package.endnotes !== undefined ? { endnotes: doc.package.endnotes } : {}),
});

// Malformed input can omit a container's children; guard every content / rows /
// cells access with the same `?? []` so a walk never crashes on a bad model.
const surfaceBlockGroups = (surfaces: ReplyThreadSurfaces): BlockContent[][] => {
  const groups: BlockContent[][] = [surfaces.body ?? []];
  for (const header of surfaces.headers?.values() ?? []) {
    groups.push(header.content ?? []);
  }
  for (const footer of surfaces.footers?.values() ?? []) {
    groups.push(footer.content ?? []);
  }
  for (const footnote of surfaces.footnotes ?? []) {
    groups.push(footnote.content ?? []);
  }
  for (const endnote of surfaces.endnotes ?? []) {
    groups.push(endnote.content ?? []);
  }
  return groups;
};

const eachParagraph = (blocks: BlockContent[], visit: (paragraph: Paragraph) => void): void => {
  for (const block of blocks ?? []) {
    if (block.type === "paragraph") {
      visit(block);
      continue;
    }
    if (block.type === "table") {
      for (const row of block.rows ?? []) {
        for (const cell of row.cells ?? []) {
          eachParagraph(cell.content ?? [], visit);
        }
      }
      continue;
    }
    eachParagraph(block.content ?? [], visit);
  }
};

type AnchorScan = {
  /** Comment ids that carry a `commentRangeStart` (a real text range). */
  rangeIds: Set<number>;
  /** Comment ids with any anchor (range start OR a `commentReference` run). */
  anchoredIds: Set<number>;
};

const scanAnchors = (surfaces: ReplyThreadSurfaces): AnchorScan => {
  const rangeIds = new Set<number>();
  const anchoredIds = new Set<number>();
  for (const group of surfaceBlockGroups(surfaces)) {
    eachParagraph(group, (paragraph) => {
      for (const item of paragraph.content ?? []) {
        if (item.type === "commentRangeStart") {
          rangeIds.add(item.id);
          anchoredIds.add(item.id);
        } else if (item.type === "commentReference") {
          anchoredIds.add(item.id);
        }
      }
    });
  }
  return { rangeIds, anchoredIds };
};

/** Replies whose parent is another comment (the threaded-discussion shape). */
const commentParentedReplies = (comments: readonly Comment[]): Comment[] => {
  const commentIds = new Set(comments.map((comment) => comment.id));
  return comments.filter(
    (comment) => comment.parentId !== undefined && commentIds.has(comment.parentId),
  );
};

const injectIntoParagraph = (
  paragraph: Paragraph,
  replyIdsByParent: Map<number, number[]>,
  parentsWithRange: ReadonlySet<number>,
): number => {
  const content = paragraph.content ?? [];
  const touchesReplyParent = content.some((item) => {
    if (item.type === "commentRangeStart" || item.type === "commentRangeEnd") {
      return replyIdsByParent.has(item.id);
    }
    // A point comment (bare reference, no range) needs a duplicated reference.
    return item.type === "commentReference" && !parentsWithRange.has(item.id)
      ? replyIdsByParent.has(item.id)
      : false;
  });
  if (!touchesReplyParent) {
    return 0;
  }

  const next: Paragraph["content"] = [];
  let injected = 0;
  for (const item of content) {
    next.push(item);
    if (item.type === "commentRangeStart") {
      for (const replyId of replyIdsByParent.get(item.id) ?? []) {
        next.push({ type: "commentRangeStart", id: replyId });
        injected += 1;
      }
    } else if (item.type === "commentRangeEnd") {
      for (const replyId of replyIdsByParent.get(item.id) ?? []) {
        next.push({ type: "commentRangeEnd", id: replyId });
      }
    } else if (item.type === "commentReference" && !parentsWithRange.has(item.id)) {
      for (const replyId of replyIdsByParent.get(item.id) ?? []) {
        next.push({ type: "commentReference", id: replyId });
        injected += 1;
      }
    }
  }
  paragraph.content = next;
  return injected;
};

const synthesizeReplyRangeMarkers = (surfaces: ReplyThreadSurfaces): number => {
  const replies = commentParentedReplies(surfaces.comments);
  if (replies.length === 0) {
    return 0;
  }

  const { rangeIds, anchoredIds } = scanAnchors(surfaces);

  // Only replies that have no anchor of their own need synthesis.
  const replyIdsByParent = new Map<number, number[]>();
  for (const reply of replies) {
    if (anchoredIds.has(reply.id) || reply.parentId === undefined) {
      continue;
    }
    const siblings = replyIdsByParent.get(reply.parentId) ?? [];
    siblings.push(reply.id);
    replyIdsByParent.set(reply.parentId, siblings);
  }
  if (replyIdsByParent.size === 0) {
    return 0;
  }

  let injected = 0;
  const visit = (paragraph: Paragraph): void => {
    injected += injectIntoParagraph(paragraph, replyIdsByParent, rangeIds);
  };
  for (const group of surfaceBlockGroups(surfaces)) {
    eachParagraph(group, visit);
  }
  return injected;
};

/**
 * Synthesize reply range markers into a document model in place, so every
 * reply comment picks up its parent's anchor before `document.xml` is
 * serialized. Idempotent and safe to call on both save paths. Returns the
 * number of markers injected.
 */
export const applyReplyThreadMarkers = (doc: Document): number =>
  synthesizeReplyRangeMarkers(surfacesFromDocument(doc));

/**
 * Whether the model holds a reply comment that has no anchor of its own yet —
 * i.e. one that still needs {@link applyReplyThreadMarkers} to run. The
 * selective-save path uses this to bail to a full repack, which is the path
 * that owns marker synthesis (selective only patches individually-changed
 * paragraphs, so it cannot add a reply's markers to the parent's paragraph).
 */
export const hasUnsynthesizedReplyRanges = (doc: Document): boolean => {
  const surfaces = surfacesFromDocument(doc);
  const replies = commentParentedReplies(surfaces.comments);
  if (replies.length === 0) {
    return false;
  }
  const { anchoredIds } = scanAnchors(surfaces);
  return replies.some((reply) => !anchoredIds.has(reply.id));
};
