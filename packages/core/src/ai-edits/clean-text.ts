import type { Mark, Node as PMNode } from "prosemirror-model";

/**
 * "Post-tracked-changes" view of a textblock: the string the user
 * would see if every existing tracked change were accepted.
 * `deletion`-marked text is skipped, `insertion`-marked text is
 * included as plain text, everything else is included as-is.
 *
 * `offsets[i]` is the absolute ProseMirror position to use when you
 * want to anchor at the character at clean-offset `i`. `offsets`
 * has length `text.length + 1` so callers can ask for the position
 * immediately after the last character, which is the right anchor
 * for an insertion at end-of-block.
 *
 * This is the view the AI should reason against (so it doesn't see
 * `"shallmust"` smashed together) and the view the apply engine's
 * find-string lookup should run against (so the same offsets it
 * sent us still resolve to the right PM positions on a doc with
 * pending tracked changes).
 */
export type CleanBlockText = {
  text: string;
  offsets: number[];
};

const DELETION_MARK = "deletion";
const INSERTION_MARK = "insertion";
const COMMENT_MARK = "comment";
const HIDDEN_MARK = "hidden";

export const buildCleanBlockText = (blockNode: PMNode, blockFrom: number): CleanBlockText => {
  let text = "";
  const offsets: number[] = [];
  let lastEnd = blockFrom + 1;
  blockNode.descendants((node, pos) => {
    if (!node.isText || node.text === undefined) {
      return true;
    }
    if (
      node.marks.some(
        (mark) => mark.type.name === DELETION_MARK || mark.type.name === HIDDEN_MARK,
      )
    ) {
      // Skip the run entirely (deleted, or OOXML w:vanish hidden text).
      // Don't update lastEnd — if the next surviving char sits right
      // after the skipped run in the live doc, we still want offsets to
      // anchor at the live position (which sits past the skipped run).
      return false;
    }
    const startPos = blockFrom + 1 + pos;
    for (let i = 0; i < node.text.length; i++) {
      offsets.push(startPos + i);
    }
    text += node.text;
    lastEnd = startPos + node.text.length;
    return true;
  });
  offsets.push(lastEnd);
  return { text, offsets };
};

/**
 * A "redline-aware" view of a textblock: the same left-to-right traversal as
 * {@link buildCleanBlockText}, but every tracked change and comment anchor is
 * rendered inline with a simple tag rather than being flattened away:
 *
 * - `<ins author="…">text</ins>` for insertion-marked runs,
 * - `<del author="…">text</del>` for deletion-marked runs (tracked moves
 *   surface as plain ins/del, since a move carries the same marks),
 * - `<comment id="N">quoted text</comment>` for comment-anchored runs.
 *
 * Nested annotations (e.g. an inserted run that is also commented) nest their
 * tags in a stable `comment > ins > del` order. Text content and attribute
 * values are XML-escaped so the tags stay unambiguous when embedded in a
 * prompt. Adjacent runs sharing the same annotation coalesce into one tag.
 *
 * This is the view a consumer embeds when it wants the model to reason about
 * the redline itself, in contrast to {@link buildCleanBlockText}'s
 * post-tracked-changes view.
 */
export const buildAnnotatedBlockText = (blockNode: PMNode): string => {
  const segments: { annotation: RunAnnotation; text: string }[] = [];
  blockNode.descendants((node) => {
    if (!node.isText || node.text === undefined) {
      return true;
    }
    const annotation = annotationOf(node.marks);
    const previous = segments.at(-1);
    if (previous && sameAnnotation(previous.annotation, annotation)) {
      previous.text += node.text;
      return false;
    }
    segments.push({ annotation, text: node.text });
    return false;
  });
  return segments.map(renderAnnotatedSegment).join("");
};

type RunAnnotation = {
  commentId: number | null;
  insertionAuthor: string | null;
  deletionAuthor: string | null;
};

const authorOf = (attrs: Mark["attrs"]): string => {
  const author = attrs["author"];
  return typeof author === "string" ? author : "";
};

const annotationOf = (marks: readonly Mark[]): RunAnnotation => {
  let commentId: number | null = null;
  let insertionAuthor: string | null = null;
  let deletionAuthor: string | null = null;
  for (const mark of marks) {
    if (mark.type.name === COMMENT_MARK) {
      const id = mark.attrs["commentId"];
      if (typeof id === "number") {
        commentId = id;
      }
    } else if (mark.type.name === INSERTION_MARK) {
      insertionAuthor = authorOf(mark.attrs);
    } else if (mark.type.name === DELETION_MARK) {
      deletionAuthor = authorOf(mark.attrs);
    }
  }
  return { commentId, insertionAuthor, deletionAuthor };
};

const sameAnnotation = (a: RunAnnotation, b: RunAnnotation): boolean =>
  a.commentId === b.commentId &&
  a.insertionAuthor === b.insertionAuthor &&
  a.deletionAuthor === b.deletionAuthor;

const escapeText = (text: string): string =>
  text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const escapeAttr = (value: string): string => escapeText(value).replaceAll('"', "&quot;");

const renderAnnotatedSegment = ({
  annotation,
  text,
}: {
  annotation: RunAnnotation;
  text: string;
}): string => {
  let inner = escapeText(text);
  if (annotation.deletionAuthor !== null) {
    inner = `<del author="${escapeAttr(annotation.deletionAuthor)}">${inner}</del>`;
  }
  if (annotation.insertionAuthor !== null) {
    inner = `<ins author="${escapeAttr(annotation.insertionAuthor)}">${inner}</ins>`;
  }
  if (annotation.commentId !== null) {
    inner = `<comment id="${annotation.commentId}">${inner}</comment>`;
  }
  return inner;
};
