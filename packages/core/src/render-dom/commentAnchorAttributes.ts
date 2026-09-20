/**
 * How a painted run advertises which comments it sits inside.
 *
 * Comment ranges overlap: a run covered by two `w:commentRangeStart` /
 * `w:commentRangeEnd` pairs belongs to both, and hover styling, the active
 * highlight and the sidebar anchor all have to see both memberships.
 *
 * `data-comment-id` carries the first id, which is what every existing reader
 * scrolls to and what keeps a single-range run's DOM exactly as the painter
 * emitted it before. `data-comment-ids` carries the whole membership and is
 * written only when there is more than one id. Readers ask through the
 * helpers here instead of reading either attribute, so the two spellings
 * cannot come to answer differently.
 */

/** Attribute carrying the run's first comment id. */
const PRIMARY_ATTRIBUTE = "data-comment-id";
/** Attribute carrying every comment id the run is inside, space separated. */
const MEMBERSHIP_ATTRIBUTE = "data-comment-ids";

/** The painted anchors inside a container, whichever comment they belong to. */
export const COMMENT_ANCHOR_SELECTOR = `[${PRIMARY_ATTRIBUTE}]`;

/** What a painter, a reader and their test fakes all offer. */
type CommentAnchor = { dataset: Record<string, string | undefined> };

/** Record the comments a painted run sits inside. */
export const writeCommentAnchorIds = (
  element: CommentAnchor,
  commentIds: readonly number[],
): void => {
  const primary = commentIds.at(0);
  if (primary === undefined) {
    return;
  }
  element.dataset["commentId"] = String(primary);
  if (commentIds.length > 1) {
    element.dataset["commentIds"] = commentIds.join(" ");
  }
};

/** Every comment id a painted anchor sits inside, first id first. */
export const commentAnchorIds = (element: CommentAnchor): readonly string[] => {
  const membership = element.dataset["commentIds"];
  if (membership !== undefined && membership.length > 0) {
    return membership.split(" ");
  }
  const primary = element.dataset["commentId"];
  return primary === undefined || primary.length === 0 ? [] : [primary];
};

/**
 * Selector for every painted anchor inside one comment, optionally scoped.
 *
 * `comment.id` is typed as a number, but a controlled `comments` prop supplied
 * by the host app is not runtime-checked, so the value is escaped before it is
 * spliced into the selector.
 */
export const commentAnchorSelector = (id: string | number, scope = ""): string => {
  const value = CSS.escape(String(id));
  return `${scope}[${PRIMARY_ATTRIBUTE}="${value}"], ${scope}[${MEMBERSHIP_ATTRIBUTE}~="${value}"]`;
};

/** First painted anchor per comment id, in one pass over `container`. */
export const indexCommentAnchors = (container: ParentNode): Map<string, HTMLElement> => {
  const anchors = new Map<string, HTMLElement>();
  for (const element of container.querySelectorAll<HTMLElement>(COMMENT_ANCHOR_SELECTOR)) {
    for (const id of commentAnchorIds(element)) {
      if (!anchors.has(id)) {
        anchors.set(id, element);
      }
    }
  }
  return anchors;
};
