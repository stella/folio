/**
 * Comment subsystem state for {@link DocxEditor}: thread storage, author
 * visibility filtering, the in-progress "add comment" flow, and the DOM
 * highlight sync. Extracted to keep the comment surface in one named seam
 * rather than scattered across the editor component. Save-coupled mutators
 * (`replaceComments`/`updateComments`) stay in the component because they
 * depend on its document-build pipeline.
 */
import {
  type RefObject,
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";

import type { Comment, Paragraph } from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";
import { isValidHexId } from "@stll/folio-core/utils/hexId";
import { PENDING_COMMENT_ID, getCommentAuthorKey, getCommentParentId } from "./commentsHelpers";

/**
 * Sanitize a `paraId`/`textId`-bearing paragraph: drop either field when it
 * isn't a well-formed OOXML long-hex id (see `isValidHexId`), leaving
 * everything else untouched. Comment paragraphs from a controlled
 * `commentsProp` come from the host app rather than our own parser/
 * serializer, so they aren't guaranteed to satisfy the id shape those trust.
 */
function sanitizeCommentParagraph(paragraph: Paragraph): Paragraph {
  const paraIdValid = paragraph.paraId === undefined || isValidHexId(paragraph.paraId);
  const textIdValid = paragraph.textId === undefined || isValidHexId(paragraph.textId);
  if (paraIdValid && textIdValid) {
    return paragraph;
  }
  const { paraId, textId, ...rest } = paragraph;
  return {
    ...rest,
    ...(paraIdValid && paraId !== undefined ? { paraId } : {}),
    ...(textIdValid && textId !== undefined ? { textId } : {}),
  };
}

/**
 * Validate a controlled `commentsProp` at the boundary before it becomes
 * editor state. A host app (or a collaboration payload relayed through it)
 * can hand back arbitrary JSON, so `id` isn't guaranteed numeric and
 * paragraph `paraId`/`textId` aren't guaranteed to be real Word ids — both
 * eventually reach XML/CSS-adjacent serialization (comment threading,
 * `[data-comment-id]` selectors). Comments with a non-finite `id` are
 * dropped entirely; malformed paraId/textId are stripped in place.
 */
function sanitizeControlledComments(comments: Comment[]): Comment[] {
  const sanitized: Comment[] = [];
  for (const comment of comments) {
    if (!Number.isFinite(comment.id)) {
      continue;
    }
    sanitized.push({
      ...comment,
      content: comment.content.map(sanitizeCommentParagraph),
    });
  }
  return sanitized;
}

type UseFolioCommentsOptions = {
  /** Current document (history head); seeds comments on first load. */
  doc: Document | null;
  autoOpenReviewSidebar: boolean;
  /** Anchor offsets from layout; re-triggers highlight sync when they shift. */
  anchorPositions: Map<string, number>;
  /** Editor content root used to locate comment-marked run nodes. */
  editorContentRef: RefObject<HTMLElement | null>;
  /**
   * Controlled comments array. When provided, thread metadata is read from
   * this prop and every mutation routes through `onCommentsChange` (e.g. Yjs
   * comment sync in collaboration backends).
   */
  commentsProp?: Comment[] | undefined;
  /** Fires whenever the comments array changes (controlled and uncontrolled). */
  onCommentsChange?: ((comments: Comment[]) => void) | undefined;
};

export function useFolioComments({
  doc,
  autoOpenReviewSidebar,
  anchorPositions,
  editorContentRef,
  commentsProp,
  onCommentsChange,
}: UseFolioCommentsOptions) {
  const [showCommentsSidebar, setShowCommentsSidebar] = useState(false);
  const [visibleCommentAuthors, setVisibleCommentAuthors] = useState<Set<string> | null>(null);
  const [activeCommentId, setActiveCommentId] = useState<number | null>(null);
  const [internalComments, setInternalComments] = useState<Comment[]>([]);
  const sanitizedCommentsProp = useMemo(
    () => (commentsProp !== undefined ? sanitizeControlledComments(commentsProp) : undefined),
    [commentsProp],
  );
  const isControlledComments = sanitizedCommentsProp !== undefined;
  const comments = isControlledComments ? sanitizedCommentsProp : internalComments;

  const commentsDirtyRef = useRef(false);
  const commentsRef = useRef(comments);
  commentsRef.current = comments;
  const onCommentsChangeRef = useRef(onCommentsChange);
  onCommentsChangeRef.current = onCommentsChange;

  const setComments = useCallback(
    (next: Comment[] | ((prev: Comment[]) => Comment[])) => {
      const resolved =
        typeof next === "function"
          ? (next as (prev: Comment[]) => Comment[])(commentsRef.current)
          : next;
      if (resolved === commentsRef.current) {
        return;
      }
      commentsRef.current = resolved;
      if (!isControlledComments) {
        setInternalComments(resolved);
      }
      onCommentsChangeRef.current?.(resolved);
    },
    [isControlledComments],
  );

  const [isAddingComment, setIsAddingComment] = useState(false);
  const [commentSelectionRange, setCommentSelectionRange] = useState<{
    from: number;
    to: number;
  } | null>(null);
  const [addCommentYPosition, setAddCommentYPosition] = useState<number | null>(null);

  // Floating "add comment" button position (relative to scroll container, null = hidden)
  const [floatingCommentBtn, setFloatingCommentBtn] = useState<{
    top: number;
    left: number;
    from: number;
    to: number;
  } | null>(null);

  // Extract comments from document model on initial load (uncontrolled only).
  const commentsLoadedRef = useRef(false);
  useEffect(() => {
    if (isControlledComments || commentsLoadedRef.current) {
      return;
    }
    if (!doc) {
      return;
    }
    const bodyComments = doc.package.document.comments;
    if (bodyComments && bodyComments.length > 0) {
      setComments(bodyComments);
      setVisibleCommentAuthors(null);
      setActiveCommentId(null);
      if (autoOpenReviewSidebar) {
        setShowCommentsSidebar(true);
      }
      commentsLoadedRef.current = true;
    }
  }, [autoOpenReviewSidebar, doc, isControlledComments, setComments]);

  const commentAuthors = useMemo(() => {
    const seen = new Set<string>();
    const authors: string[] = [];
    for (const comment of comments) {
      const commentAuthor = getCommentAuthorKey(comment.author);
      if (!seen.has(commentAuthor)) {
        seen.add(commentAuthor);
        authors.push(commentAuthor);
      }
    }
    return authors;
  }, [comments]);

  const visibleCommentAuthorSet = useMemo(
    () => visibleCommentAuthors ?? new Set(commentAuthors),
    [visibleCommentAuthors, commentAuthors],
  );

  const visibleCommentIds = useMemo(() => {
    const ids = new Set<number>([PENDING_COMMENT_ID]);
    for (const comment of comments) {
      if (visibleCommentAuthorSet.has(getCommentAuthorKey(comment.author))) {
        ids.add(comment.id);
      }
    }
    return ids;
  }, [comments, visibleCommentAuthorSet]);

  const visibleComments = useMemo(() => {
    const visibleRootIds = new Set<number>();
    for (const comment of comments) {
      const parentId = getCommentParentId(comment);
      if (parentId === null || parentId === undefined || !visibleCommentIds.has(comment.id)) {
        continue;
      }
      visibleRootIds.add(parentId);
    }
    return comments.filter((comment) => {
      const parentId = getCommentParentId(comment);
      if (parentId !== null && parentId !== undefined) {
        return visibleCommentIds.has(comment.id);
      }
      return visibleCommentIds.has(comment.id) || visibleRootIds.has(comment.id);
    });
  }, [comments, visibleCommentIds]);

  const activeCommentVisible = activeCommentId !== null && visibleCommentIds.has(activeCommentId);

  useEffect(() => {
    if (!activeCommentVisible) {
      setActiveCommentId(null);
    }
  }, [activeCommentVisible]);

  const syncCommentHighlightStyles = useCallback(() => {
    const root = editorContentRef.current;
    if (!root) {
      return;
    }

    const nodes = root.querySelectorAll<HTMLElement>(".layout-run-text[data-comment-id]");
    for (const node of nodes) {
      const commentId = Number.parseInt(node.dataset["commentId"] ?? "", 10);
      const isPending = commentId === PENDING_COMMENT_ID;
      const isVisible = isPending || visibleCommentIds.has(commentId);
      if (!isVisible) {
        node.style.backgroundColor = "transparent";
        node.style.borderBottom = "2px solid transparent";
        node.style.boxShadow = "none";
        delete node.dataset["activeComment"];
        continue;
      }

      if (activeCommentId === commentId) {
        node.style.backgroundColor = "var(--doc-comment-active-bg, rgba(255, 212, 0, 0.22))";
        node.style.borderBottom =
          "1px solid var(--doc-comment-active-border, rgba(180, 130, 0, 0.62))";
        node.style.boxShadow = "none";
        node.dataset["activeComment"] = "true";
        continue;
      }

      node.style.backgroundColor = "var(--doc-comment-bg, rgba(255, 212, 0, 0.08))";
      node.style.borderBottom = "1px solid var(--doc-comment-border, rgba(180, 130, 0, 0.24))";
      node.style.boxShadow = "none";
      delete node.dataset["activeComment"];
    }
  }, [visibleCommentIds, activeCommentId, editorContentRef]);

  useLayoutEffect(() => {
    syncCommentHighlightStyles();
  }, [syncCommentHighlightStyles, anchorPositions]);

  useEffect(() => {
    syncCommentHighlightStyles();
    let secondFrame: number | null = null;
    const firstFrame = requestAnimationFrame(() => {
      syncCommentHighlightStyles();
      secondFrame = requestAnimationFrame(syncCommentHighlightStyles);
    });
    const timeout = setTimeout(syncCommentHighlightStyles, 120);
    return () => {
      cancelAnimationFrame(firstFrame);
      if (secondFrame !== null) {
        cancelAnimationFrame(secondFrame);
      }
      clearTimeout(timeout);
    };
  }, [comments, syncCommentHighlightStyles]);

  return {
    comments,
    setComments,
    isControlledComments,
    commentsRef,
    commentsDirtyRef,
    commentsLoadedRef,
    showCommentsSidebar,
    setShowCommentsSidebar,
    setVisibleCommentAuthors,
    activeCommentId,
    setActiveCommentId,
    isAddingComment,
    setIsAddingComment,
    commentSelectionRange,
    setCommentSelectionRange,
    addCommentYPosition,
    setAddCommentYPosition,
    floatingCommentBtn,
    setFloatingCommentBtn,
    commentAuthors,
    visibleCommentAuthorSet,
    visibleComments,
    syncCommentHighlightStyles,
  };
}
