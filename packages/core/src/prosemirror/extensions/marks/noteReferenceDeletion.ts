import type { Node as PMNode } from "prosemirror-model";

type DeletionRange = { from: number; to: number };

/** Keep the text carrying one OOXML note reference together during deletion. */
export const expandNoteReferenceDeletionRange = (
  doc: PMNode,
  from: number,
  to: number,
): DeletionRange | null => {
  let expandedFrom = from;
  let expandedTo = to;
  let intersectsReference = false;

  doc.nodesBetween(from, to, (node, pos) => {
    const reference = node.marks.find((mark) => mark.type.name === "footnoteRef");
    if (!node.isText || !reference) {
      return;
    }
    intersectsReference = true;
    const $pos = doc.resolve(pos);
    const parent = $pos.parent;
    const index = $pos.index();
    let referenceFrom = pos;
    let referenceTo = pos + node.nodeSize;

    for (let left = index - 1; left >= 0; left--) {
      const sibling = parent.child(left);
      if (!sibling.isText || !sibling.marks.some((mark) => mark.eq(reference))) break;
      referenceFrom -= sibling.nodeSize;
    }
    for (let right = index + 1; right < parent.childCount; right++) {
      const sibling = parent.child(right);
      if (!sibling.isText || !sibling.marks.some((mark) => mark.eq(reference))) break;
      referenceTo += sibling.nodeSize;
    }

    expandedFrom = Math.min(expandedFrom, referenceFrom);
    expandedTo = Math.max(expandedTo, referenceTo);
  });

  return intersectsReference ? { from: expandedFrom, to: expandedTo } : null;
};

/** Each whole note reference [`from`, `to`) of `doc` touches, in document order. */
export const noteReferenceRanges = (doc: PMNode, from: number, to: number): DeletionRange[] => {
  const ranges: DeletionRange[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText || !node.marks.some((mark) => mark.type.name === "footnoteRef")) {
      return;
    }
    const whole = expandNoteReferenceDeletionRange(doc, pos, pos + node.nodeSize);
    const last = ranges.at(-1);
    if (whole && (!last || whole.from >= last.to)) {
      ranges.push(whole);
    }
  });
  return ranges;
};
