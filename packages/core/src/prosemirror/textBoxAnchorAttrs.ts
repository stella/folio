import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import type { TextBoxAnchorAttrs } from "./schema/nodes";

const attrsCache = new WeakMap<PMNode, TextBoxAnchorAttrs>();

export const readTextBoxAnchorAttrs = (
  node: PMNode,
): ReadProseMirrorAttrsResult<TextBoxAnchorAttrs> => {
  const issues: ProseMirrorAttrIssue[] = [];
  if (node.type.name !== "textBoxAnchor") {
    issues.push({
      path: "textBoxAnchor.type.name",
      message: `Expected textBoxAnchor, got ${node.type.name}.`,
    });
  }
  const anchorId = node.attrs["anchorId"];
  if (typeof anchorId !== "string" || anchorId.length === 0) {
    issues.push({
      path: "textBoxAnchor.attrs.anchorId",
      message: "Expected a non-empty string.",
    });
  }
  if (issues.length > 0 || typeof anchorId !== "string") {
    return { ok: false, issues };
  }
  return { ok: true, value: { anchorId } };
};

export const expectTextBoxAnchorAttrs = (node: PMNode): TextBoxAnchorAttrs => {
  const cached = attrsCache.get(node);
  if (cached) {
    return cached;
  }
  const result = readTextBoxAnchorAttrs(node);
  if (!result.ok) {
    const details = result.issues.map((issue) => `${issue.path}: ${issue.message}`).join("\n");
    panic(`Invalid ProseMirror text box anchor attrs:\n${details}`);
  }
  attrsCache.set(node, result.value);
  return result.value;
};
