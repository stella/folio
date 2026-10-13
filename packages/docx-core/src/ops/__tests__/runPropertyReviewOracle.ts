import type { Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { storyParagraphs } from "../blocks";
import { childNodes, type InlineNode } from "../leaves";
import { DOCUMENT_OP_TYPES, OP_STORIES, type RevisionDecision } from "../types";

/** Compare complete formatting outcomes when one review consumes an earlier one. */
export const resolvePendingRunFormatting = (
  document: Document,
  decision: RevisionDecision,
): Document => {
  const revisionIds = new Set<number>();
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (node.type === "run")
        for (const change of node.propertyChanges ?? []) revisionIds.add(change.info.id);
      visit(childNodes(node));
    }
  };
  for (const { paragraph } of storyParagraphs(document.package.document)) visit(paragraph.content);
  if (revisionIds.size === 0) return document;
  return applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...revisionIds],
    decision,
  }).unwrap().document;
};
