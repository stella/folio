import type { Mark, Node as PMNode } from "prosemirror-model";

import { expectPageBreakRunAttrs } from "../prosemirror/attrs";
import {
  runFormattingInlineAtomCleanText,
  runFormattingInlineControlCharacter,
} from "../prosemirror/runFormattingInlineCarriers";
import type { PageBreakRunAttrs } from "../prosemirror/schema/nodes";

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
  /**
   * Zero-width structural carriers projected at their clean-text boundary.
   *
   * A single clean offset cannot identify both sides of an atom: the PM
   * position immediately after the preceding character is before the atom,
   * while the next character starts after it. Callers that select text must
   * resolve through {@link resolveCleanTextRange}; indexing `offsets`
   * directly can accidentally absorb the carrier into an adjacent range.
   */
  structuralBoundaries: readonly CleanTextStructuralBoundary[];
};

export type CleanTextStructuralBoundary =
  | {
      type: "pageBreakRun";
      /** Clean-text offset at which the zero-width carrier occurs. */
      offset: number;
      /** PM range owned by the carrier. */
      from: number;
      to: number;
      /** Preserved even though `clear` does not alter page-break layout. */
      clear?: PageBreakRunAttrs["clear"];
      /** Whether the post-tracked-changes projection retains this carrier. */
      presentInCleanView: boolean;
    }
  | {
      /**
       * A field result: several clean-text characters over a single PM
       * position. Only the whole span is addressable, so a text range may
       * start or end at its edges but never inside it.
       */
      type: "field";
      offset: number;
      /** Characters the result contributes; always greater than one. */
      length: number;
      from: number;
      to: number;
    };

const EMPTY_CLEAN_TEXT_STRUCTURAL_BOUNDARIES: readonly CleanTextStructuralBoundary[] =
  Object.freeze([]);

const cutsIntoSpan = (
  { offset, length }: { offset: number; length: number },
  boundaryOffset: number,
): boolean => boundaryOffset > offset && boundaryOffset < offset + length;

type ResolveCleanTextRangeOptions = {
  cleanBlock: CleanBlockText;
  startOffset: number;
  endOffset: number;
};

/**
 * Resolve clean-text offsets without selecting a structural atom.
 *
 * A range wholly on one side of a zero-width boundary is biased away from the
 * atom. A range spanning both sides, or cutting into a field result, is
 * unrepresentable as a generic text selection and returns `null`; a structural
 * operation must own that mutation instead.
 */
export const resolveCleanTextRange = ({
  cleanBlock,
  startOffset,
  endOffset,
}: ResolveCleanTextRangeOptions): { from: number; to: number } | null => {
  if (
    !Number.isInteger(startOffset) ||
    !Number.isInteger(endOffset) ||
    startOffset < 0 ||
    endOffset < startOffset ||
    endOffset > cleanBlock.text.length
  ) {
    return null;
  }

  const baseFrom = cleanBlock.offsets[startOffset];
  const baseTo = cleanBlock.offsets[endOffset];
  if (baseFrom === undefined || baseTo === undefined) {
    return null;
  }

  const { structuralBoundaries } = cleanBlock;
  if (structuralBoundaries.length === 0) {
    return baseFrom <= baseTo ? { from: baseFrom, to: baseTo } : null;
  }

  let from = baseFrom;
  let to = baseTo;
  for (const boundary of structuralBoundaries) {
    if (boundary.type === "field") {
      if (cutsIntoSpan(boundary, startOffset) || cutsIntoSpan(boundary, endOffset)) {
        return null;
      }
      continue;
    }
    if (boundary.offset > startOffset && boundary.offset < endOffset) {
      return null;
    }
    if (boundary.offset === startOffset) {
      from = Math.max(from, boundary.to);
    }
    if (startOffset !== endOffset && boundary.offset === endOffset) {
      to = Math.min(to, boundary.from);
    }
  }

  if (startOffset === endOffset) {
    return { from, to: from };
  }

  return from <= to ? { from, to } : null;
};

const DELETION_MARK = "deletion";
const INSERTION_MARK = "insertion";
const COMMENT_MARK = "comment";
const HIDDEN_MARK = "hidden";
const isOmittedFromCleanView = (node: PMNode): boolean =>
  node.marks.some((mark) => mark.type.name === DELETION_MARK || mark.type.name === HIDDEN_MARK);

export type BuildCleanBlockTextOptions = {
  /**
   * `"text"` reads a field as the result Word shows, which is the view a reader
   * and the AI reason against. `"omitted"` drops it, which is what an alignment
   * coordinate needs: an atom present on only one side must not shift the
   * offsets that locate it.
   */
  fieldResults: "text" | "omitted";
};

const DEFAULT_BUILD_CLEAN_BLOCK_TEXT_OPTIONS: BuildCleanBlockTextOptions = { fieldResults: "text" };

export const buildCleanBlockText = (
  blockNode: PMNode,
  blockFrom: number,
  { fieldResults }: BuildCleanBlockTextOptions = DEFAULT_BUILD_CLEAN_BLOCK_TEXT_OPTIONS,
): CleanBlockText => {
  let text = "";
  const offsets: number[] = [];
  let structuralBoundaries: CleanTextStructuralBoundary[] | undefined;
  let lastEnd = blockFrom + 1;
  blockNode.descendants((node, pos) => {
    if (node.type.name === "pageBreakRun") {
      const from = blockFrom + 1 + pos;
      const { clear } = expectPageBreakRunAttrs(node);
      (structuralBoundaries ??= []).push({
        type: "pageBreakRun",
        offset: text.length,
        from,
        to: from + node.nodeSize,
        ...(clear !== undefined ? { clear } : {}),
        presentInCleanView: !node.marks.some(({ type }) => type.name === DELETION_MARK),
      });
      return false;
    }
    const atomText =
      fieldResults === "omitted"
        ? runFormattingInlineControlCharacter(node)
        : runFormattingInlineAtomCleanText(node);
    if (atomText !== null) {
      if (isOmittedFromCleanView(node)) {
        return false;
      }
      const startPos = blockFrom + 1 + pos;
      // Every character of a multi-character atom anchors at the atom itself:
      // its interior has no PM position of its own, so the boundary is what
      // stops a caller slicing into it.
      if (atomText.length > 1) {
        (structuralBoundaries ??= []).push({
          type: "field",
          offset: text.length,
          length: atomText.length,
          from: startPos,
          to: startPos + node.nodeSize,
        });
      }
      for (let index = 0; index < atomText.length; index++) {
        offsets.push(startPos);
      }
      text += atomText;
      lastEnd = startPos + node.nodeSize;
      return false;
    }
    if (!node.isText || node.text === undefined) {
      return true;
    }
    if (isOmittedFromCleanView(node)) {
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
  return {
    text,
    offsets,
    structuralBoundaries: structuralBoundaries ?? EMPTY_CLEAN_TEXT_STRUCTURAL_BOUNDARIES,
  };
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
    const text = node.isText ? node.text : runFormattingInlineAtomCleanText(node);
    if (text === undefined || text === null) {
      return true;
    }
    const annotation = annotationOf(node.marks);
    const previous = segments.at(-1);
    if (previous && sameAnnotation(previous.annotation, annotation)) {
      previous.text += text;
      return false;
    }
    segments.push({ annotation, text });
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
