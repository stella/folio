import type { Node as PMNode, ResolvedPos } from "prosemirror-model";

/**
 * A node the FILE holds as a child of the container it appears in.
 *
 * The editor keeps a text box as a block-level sibling of the paragraph that
 * anchors it; the file hangs it off that paragraph's own run, so it is not a
 * child of the container at all. A paragraph and a table are children in both.
 */
const isAContainerChild = (node: PMNode, paragraphTypeName: string): boolean =>
  node.type.name === paragraphTypeName || node.type.spec["tableRole"] === "table";

/** Whether every child from `fromIndex` on is one the file does not hold here. */
const nothingFollowsIn = (
  parent: PMNode,
  fromIndex: number,
  paragraphTypeName: string,
): boolean => {
  for (let index = fromIndex; index < parent.childCount; index++) {
    if (isAContainerChild(parent.child(index), paragraphTypeName)) {
      return false;
    }
  }
  return true;
};

/**
 * Whether the paragraph starting at `at` is the one its container ENDS with,
 * which is the one place a paragraph mark cannot carry a revision.
 *
 * A deleted mark asks a reader to join the paragraph it ends with the one
 * after it, and an inserted mark says that break was added, so rejecting it
 * closes the paragraph back over the next one. A body, a table cell, a header
 * or footer, a note and a text box each end with a paragraph that nothing
 * follows, so neither direction can be resolved there: a consumer is left with
 * a revision it can only ignore, and the redline no longer resolves to the two
 * documents it was written from.
 *
 * A paragraph before a TABLE is not one of those: the table is a following
 * sibling, so the paragraph does not end anything and carries its own mark
 * like any other.
 */
export const paragraphEndsItsContainer = (at: ResolvedPos, paragraphTypeName: string): boolean =>
  nothingFollowsIn(at.parent, at.index() + 1, paragraphTypeName);

/**
 * Every paragraph a container ENDS with in `doc`, with its position.
 *
 * One walk for the whole document: asking the question of each paragraph in
 * turn costs a scan of its siblings each time, which is quadratic in a body
 * whose paragraphs a batch has just filled.
 */
export const finalParagraphsOf = (
  doc: PMNode,
  paragraphTypeName: string,
): { position: number; node: PMNode }[] => {
  const found: { position: number; node: PMNode }[] = [];
  const collect = (node: PMNode, contentStart: number): void => {
    let last: { position: number; node: PMNode } | null = null;
    node.forEach((child, offset) => {
      const position = contentStart + offset;
      if (child.type.name === paragraphTypeName) {
        last = { position, node: child };
      } else if (isAContainerChild(child, paragraphTypeName)) {
        last = null;
      }
      if (!child.isTextblock && child.childCount > 0) {
        collect(child, position + 1);
      }
    });
    if (last !== null) {
      found.push(last);
    }
  };
  collect(doc, 0);
  return found;
};

/** A paragraph's mark, when it says the break was ADDED. */
const marksAnAddedBreak = (node: PMNode): boolean => {
  const mark: unknown = node.attrs["pPrMark"];
  if (typeof mark !== "object" || mark === null || !("kind" in mark)) {
    return false;
  }
  return mark.kind === "ins" || mark.kind === "moveTo";
};

/**
 * The paragraph an added break can rotate BACK onto from the one at `at`: the
 * paragraph the run of inserted ones was appended after.
 *
 * The walk crosses the inserted paragraphs — their own marks stay where they
 * are, and only which paragraph is left markless changes — and stops at the
 * first one whose mark is free. `null` when there is none: a table in between,
 * which no mark joins across; a paragraph whose mark says something else,
 * which the rotation must not overwrite; or a container whose every paragraph
 * is new.
 */
export const addedBreakCarrierBefore = (
  at: ResolvedPos,
  paragraphTypeName: string,
): { position: number; node: PMNode } | null => {
  let position = at.pos;
  for (let index = at.index() - 1; index >= 0; index--) {
    const sibling = at.parent.child(index);
    position -= sibling.nodeSize;
    if (sibling.type.name !== paragraphTypeName) {
      // A table stops the walk; a node the file does not hold here at all — a
      // text box — is not between the two paragraphs in the first place.
      if (isAContainerChild(sibling, paragraphTypeName)) {
        return null;
      }
      continue;
    }
    if (sibling.attrs["pPrMark"] == null) {
      return { position, node: sibling };
    }
    if (!marksAnAddedBreak(sibling)) {
      return null;
    }
  }
  return null;
};
