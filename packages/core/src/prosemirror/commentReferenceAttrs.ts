import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import type { CommentReferenceAttrs } from "./schema/nodes";

const attrsCache = new WeakMap<PMNode, CommentReferenceAttrs>();

export const readCommentReferenceAttrs = (
  node: PMNode,
): ReadProseMirrorAttrsResult<CommentReferenceAttrs> => {
  const issues: ProseMirrorAttrIssue[] = [];
  if (node.type.name !== "commentReference") {
    issues.push({
      path: "commentReference.type.name",
      message: `Expected commentReference, got ${node.type.name}.`,
    });
  }
  const commentId = node.attrs["commentId"];
  if (typeof commentId !== "number" || !Number.isSafeInteger(commentId)) {
    issues.push({
      path: "commentReference.attrs.commentId",
      message: "Expected a safe integer.",
    });
  }

  if (issues.length > 0 || typeof commentId !== "number") {
    return { ok: false, issues };
  }
  return { ok: true, value: { commentId } };
};

export const expectCommentReferenceAttrs = (node: PMNode): CommentReferenceAttrs => {
  const cached = attrsCache.get(node);
  if (cached) {
    return cached;
  }
  const result = readCommentReferenceAttrs(node);
  if (!result.ok) {
    const details = result.issues.map((issue) => `${issue.path}: ${issue.message}`).join("\n");
    panic(`Invalid ProseMirror comment reference attrs:\n${details}`);
  }
  attrsCache.set(node, result.value);
  return result.value;
};
