import type { Node as PMNode } from "prosemirror-model";

/**
 * Whether a node, or any node under it, is one `carries` names, decided once
 * per node.
 *
 * ProseMirror nodes are immutable and an edit shares every subtree it does
 * not touch, so a walk that only needs the nodes `carries` names can skip
 * every child this answers `false` for: an edit then costs the path it
 * changed, not every node of the document. Leaves are answered directly
 * rather than kept.
 */
export const subtreeHolds = (
  carries: (node: PMNode) => boolean,
): ((node: PMNode) => boolean) => {
  const known = new WeakMap<PMNode, boolean>();
  const holds = (node: PMNode): boolean => {
    if (node.childCount === 0) return carries(node);
    const kept = known.get(node);
    if (kept !== undefined) return kept;
    let held = carries(node);
    for (let index = 0; !held && index < node.childCount; index += 1)
      held = holds(node.child(index));
    known.set(node, held);
    return held;
  };
  return holds;
};
