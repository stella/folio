import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { ParagraphContent } from "../types/content";
import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import { isDisplacedByCustomXml } from "./bookmarkBoundaryAttrs";

type CommentRangeBoundaryMarker = Extract<
  ParagraphContent,
  { type: "commentRangeStart" | "commentRangeEnd" }
>;

type CommentRangeBoundaryAttrs = {
  marker: CommentRangeBoundaryMarker;
};

const attrsCache = new WeakMap<PMNode, CommentRangeBoundaryAttrs>();

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

export const readCommentRangeBoundaryMarker = (
  value: unknown,
): CommentRangeBoundaryMarker | undefined => {
  if (!isRecord(value)) return undefined;
  const marker = value;
  const type = marker["type"];
  const id = marker["id"];
  const displacedByCustomXml = marker["displacedByCustomXml"];
  if (
    (type !== "commentRangeStart" && type !== "commentRangeEnd") ||
    typeof id !== "number" ||
    !Number.isSafeInteger(id) ||
    id < 0 ||
    (displacedByCustomXml !== undefined &&
      !isDisplacedByCustomXml(displacedByCustomXml))
  ) {
    return undefined;
  }
  return {
    type,
    id,
    ...(displacedByCustomXml === undefined ? {} : { displacedByCustomXml }),
  };
};

export const readCommentRangeBoundaryAttrs = (
  node: PMNode,
): ReadProseMirrorAttrsResult<CommentRangeBoundaryAttrs> => {
  const issues: ProseMirrorAttrIssue[] = [];
  if (node.type.name !== "commentRangeBoundary") {
    issues.push({
      path: "commentRangeBoundary.type.name",
      message: `Expected commentRangeBoundary, got ${node.type.name}.`,
    });
  }
  const marker = readCommentRangeBoundaryMarker(node.attrs["marker"]);
  if (!marker) {
    issues.push({
      path: "commentRangeBoundary.attrs.marker",
      message: "Expected a comment range start or end marker.",
    });
  }
  if (issues.length > 0 || !marker) return { ok: false, issues };
  return { ok: true, value: { marker } };
};

export const expectCommentRangeBoundaryAttrs = (
  node: PMNode,
): CommentRangeBoundaryAttrs => {
  const cached = attrsCache.get(node);
  if (cached) return cached;
  const result = readCommentRangeBoundaryAttrs(node);
  if (!result.ok) {
    const details = result.issues
      .map((issue) => `${issue.path}: ${issue.message}`)
      .join("\n");
    panic(`Invalid ProseMirror comment range boundary attrs:\n${details}`);
  }
  attrsCache.set(node, result.value);
  return result.value;
};
