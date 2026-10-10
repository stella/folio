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
  useEffectEvent,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";

import {
  COMMENT_ANCHOR_SELECTOR,
  commentAnchorIds,
} from "@stll/folio-core/render-dom/commentAnchorAttributes";
import type { Comment, Paragraph } from "@stll/folio-core/types/content";
import type { Document } from "@stll/folio-core/types/document";
import { isValidHexId } from "@stll/folio-core/utils/hexId";
import {
  PENDING_COMMENT_ID,
  countOpenCommentThreads,
  createComment as buildComment,
  getCommentAuthorKey,
  getCommentParentId,
  seedCommentIdAbove,
} from "./commentsHelpers";

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
  /**
   * Comments the canonical session has committed. When present, the sidebar
   * lists and filters these instead of the host or internal array, so canonical
   * creation, resolution and deletion show without waiting for host feedback.
   */
  committedComments?: readonly Comment[] | null | undefined;
};

export function useFolioComments({
  doc,
  autoOpenReviewSidebar,
  anchorPositions,
  editorContentRef,
  commentsProp,
  onCommentsChange,
  committedComments,
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

  // Reserve before browser events can mint comments, including controlled replies.
  useLayoutEffect(() => {
    for (const { id } of doc?.package.document.comments ?? []) {
      seedCommentIdAbove(id);
    }
    for (const { id } of comments) {
      seedCommentIdAbove(id);
    }
  }, [comments, doc]);

  const commentsDirtyRef = useRef(false);
  const getCommentsDirty = useCallback(() => commentsDirtyRef.current, []);
  const setCommentsDirty = useCallback((dirty: boolean) => {
    commentsDirtyRef.current = dirty;
  }, []);
  // Reconcile every commit, including when a controlled host rejects an update
  // and rerenders the same array. Event-time writes remain available until that
  // commit. The hook keeps the writable handle private and exports a read-only
  // view; mutate only through `setComments`.
  const commentsRef = useRef(comments);
  useLayoutEffect(() => {
    commentsRef.current = comments;
  });
  const readonlyCommentsRef: Readonly<RefObject<Comment[]>> = commentsRef;
  const onCommentsChangeRef = useRef(onCommentsChange);
  useLayoutEffect(() => {
    onCommentsChangeRef.current = onCommentsChange;
  }, [onCommentsChange]);

  const createComment = useCallback(
    (text: string, author: string, parentId?: number) => {
      for (const { id } of doc?.package.document.comments ?? []) {
        seedCommentIdAbove(id);
      }
      for (const { id } of commentsRef.current) {
        seedCommentIdAbove(id);
      }
      return buildComment(text, author, parentId);
    },
    [doc],
  );

  const setComments = useCallback(
    (next: Comment[] | ((prev: Comment[]) => Comment[])) => {
      const resolved =
        typeof next === "function"
          ? (next as (prev: Comment[]) => Comment[])(commentsRef.current)
          : next;
      if (resolved === commentsRef.current) {
        return;
      }
      for (const { id } of resolved) {
        seedCommentIdAbove(id);
      }
      // The owning setter: the ref write is paired with the state update below
      // (or, when controlled, with the host applying `onCommentsChange`), so
      // same-tick readers and the next render agree.
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

  // Initialize once when authoritative comments become available. The loader
  // explicitly resets this lifecycle when it replaces the document.
  const [commentsLoaded, setCommentsLoaded] = useState(false);
  const [loadedCommentsNotification, setLoadedCommentsNotification] = useState<{
    comments: Comment[];
  } | null>(null);
  const resetLoadedComments = useCallback(() => setCommentsLoaded(false), []);
  const bodyComments = committedComments ?? doc?.package.document.comments;
  if (
    !commentsLoaded &&
    !(committedComments == null && isControlledComments) &&
    bodyComments &&
    bodyComments.length > 0
  ) {
    setCommentsLoaded(true);
    if (committedComments == null && !isControlledComments && doc?.package.document.comments) {
      const loadedComments = doc.package.document.comments;
      setInternalComments(loadedComments);
      setLoadedCommentsNotification({ comments: loadedComments });
    }
    setVisibleCommentAuthors(null);
    setActiveCommentId(null);
    if (autoOpenReviewSidebar && countOpenCommentThreads(bodyComments) > 0) {
      setShowCommentsSidebar(true);
    }
  }

  const notifyLoadedComments = useEffectEvent((loaded: Comment[]) => onCommentsChange?.(loaded));
  const lastLoadedNotificationRef = useRef<typeof loadedCommentsNotification>(null);
  useEffect(() => {
    if (
      loadedCommentsNotification === null ||
      lastLoadedNotificationRef.current === loadedCommentsNotification
    )
      return;
    lastLoadedNotificationRef.current = loadedCommentsNotification;
    notifyLoadedComments(loadedCommentsNotification.comments);
  }, [loadedCommentsNotification]);

  const listedComments = committedComments ?? comments;

  const commentAuthors = useMemo(() => {
    const seen = new Set<string>();
    const authors: string[] = [];
    for (const comment of listedComments) {
      const commentAuthor = getCommentAuthorKey(comment.author);
      if (!seen.has(commentAuthor)) {
        seen.add(commentAuthor);
        authors.push(commentAuthor);
      }
    }
    return authors;
  }, [listedComments]);

  const visibleCommentAuthorSet = useMemo(
    () => visibleCommentAuthors ?? new Set(commentAuthors),
    [visibleCommentAuthors, commentAuthors],
  );

  const visibleCommentIds = useMemo(() => {
    const ids = new Set<number>([PENDING_COMMENT_ID]);
    for (const comment of listedComments) {
      if (visibleCommentAuthorSet.has(getCommentAuthorKey(comment.author))) {
        ids.add(comment.id);
      }
    }
    return ids;
  }, [listedComments, visibleCommentAuthorSet]);

  const visibleComments = useMemo(() => {
    const visibleRootIds = new Set<number>();
    for (const comment of listedComments) {
      const parentId = getCommentParentId(comment);
      if (parentId === null || parentId === undefined || !visibleCommentIds.has(comment.id)) {
        continue;
      }
      visibleRootIds.add(parentId);
    }
    return listedComments.filter((comment) => {
      const parentId = getCommentParentId(comment);
      if (parentId !== null && parentId !== undefined) {
        return visibleCommentIds.has(comment.id);
      }
      return visibleCommentIds.has(comment.id) || visibleRootIds.has(comment.id);
    });
  }, [listedComments, visibleCommentIds]);

  const activeCommentVisible = activeCommentId !== null && visibleCommentIds.has(activeCommentId);

  if (activeCommentId !== null && !activeCommentVisible) {
    setActiveCommentId(null);
  }

  const syncCommentHighlightStyles = useCallback(() => {
    const root = editorContentRef.current;
    if (!root) {
      return;
    }

    const nodes = root.querySelectorAll<HTMLElement>(`.layout-run-text${COMMENT_ANCHOR_SELECTOR}`);
    for (const node of nodes) {
      // A run inside overlapping ranges belongs to every one of them, so it
      // stays lit while any of its comments is visible and goes active for
      // whichever of them the user is on.
      const commentIds = commentAnchorIds(node).map((id) => Number.parseInt(id, 10));
      const isVisible = commentIds.some(
        (commentId) => commentId === PENDING_COMMENT_ID || visibleCommentIds.has(commentId),
      );
      if (!isVisible) {
        node.style.backgroundColor = "transparent";
        node.style.borderBottom = "2px solid transparent";
        node.style.boxShadow = "none";
        delete node.dataset["activeComment"];
        continue;
      }

      if (activeCommentId !== null && commentIds.includes(activeCommentId)) {
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

  // A layout-map commit can replace painted marks without changing thread metadata.
  const highlightLayoutCommit = useEffectEvent(
    (_positions: Map<string, number>, sync: () => void) => sync(),
  );
  useLayoutEffect(() => {
    highlightLayoutCommit(anchorPositions, syncCommentHighlightStyles);
  }, [syncCommentHighlightStyles, anchorPositions]);

  const scheduleHighlightCommit = useEffectEvent((_comments: Comment[], sync: () => void) => {
    sync();
    let secondFrame: number | null = null;
    const firstFrame = requestAnimationFrame(() => {
      sync();
      secondFrame = requestAnimationFrame(sync);
    });
    const timeout = setTimeout(sync, 120);
    return () => {
      cancelAnimationFrame(firstFrame);
      if (secondFrame !== null) {
        cancelAnimationFrame(secondFrame);
      }
      clearTimeout(timeout);
    };
  });
  useEffect(
    () => scheduleHighlightCommit(comments, syncCommentHighlightStyles),
    [comments, syncCommentHighlightStyles],
  );

  return {
    comments,
    setComments,
    createComment,
    isControlledComments,
    commentsRef: readonlyCommentsRef,
    getCommentsDirty,
    setCommentsDirty,
    resetLoadedComments,
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
