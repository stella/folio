import type { Node as PMNode } from "prosemirror-model";

import { createResolvedCommentPointCollector } from "../../prosemirror/replacedAnnotations";

type ResolvedCommentPointsOptions = {
  doc: PMNode;
  removes: (node: PMNode, position: number) => boolean;
  retiresStructure: (node: PMNode, position: number) => boolean;
};

/** Ranged comments whose covered non-carrier content is removed by a resolution. */
export const resolvedCommentPoints = ({
  doc,
  removes,
  retiresStructure,
}: ResolvedCommentPointsOptions): Map<number, number> => {
  const collector = createResolvedCommentPointCollector({
    removes,
    retiresStructure,
  });
  doc.descendants((node, position, parent, index) =>
    collector.visit(
      node,
      position,
      index > 0 ? (parent?.maybeChild(index - 1) ?? undefined) : undefined,
    ),
  );
  return collector.finish();
};
