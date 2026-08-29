import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import type { BookmarkBoundaryAttrs } from "./schema/nodes";

const attrsCache = new WeakMap<PMNode, BookmarkBoundaryAttrs>();

export const readBookmarkBoundaryAttrs = (
  node: PMNode,
): ReadProseMirrorAttrsResult<BookmarkBoundaryAttrs> => {
  const issues: ProseMirrorAttrIssue[] = [];
  if (node.type.name !== "bookmarkBoundary") {
    issues.push({
      path: "bookmarkBoundary.type.name",
      message: `Expected bookmarkBoundary, got ${node.type.name}.`,
    });
  }
  const type = node.attrs["type"];
  if (type !== "start" && type !== "end") {
    issues.push({
      path: "bookmarkBoundary.attrs.type",
      message: 'Expected "start" or "end".',
    });
  }
  const id = node.attrs["id"];
  if (typeof id !== "number" || !Number.isInteger(id) || id < 0) {
    issues.push({
      path: "bookmarkBoundary.attrs.id",
      message: "Expected a non-negative integer.",
    });
  }
  const name = node.attrs["name"];
  if (type === "start" && (typeof name !== "string" || name.length === 0)) {
    issues.push({
      path: "bookmarkBoundary.attrs.name",
      message: "Expected a non-empty string for a bookmark start.",
    });
  }
  const colFirst = node.attrs["colFirst"];
  if (
    colFirst !== undefined &&
    colFirst !== null &&
    (typeof colFirst !== "number" || !Number.isInteger(colFirst) || colFirst < 0)
  ) {
    issues.push({
      path: "bookmarkBoundary.attrs.colFirst",
      message: "Expected a non-negative integer.",
    });
  }
  const colLast = node.attrs["colLast"];
  if (
    colLast !== undefined &&
    colLast !== null &&
    (typeof colLast !== "number" || !Number.isInteger(colLast) || colLast < 0)
  ) {
    issues.push({
      path: "bookmarkBoundary.attrs.colLast",
      message: "Expected a non-negative integer.",
    });
  }

  if (issues.length > 0 || typeof id !== "number") {
    return { ok: false, issues };
  }
  if (type === "start" && typeof name === "string") {
    return {
      ok: true,
      value: {
        type,
        id,
        name,
        ...(typeof colFirst === "number" ? { colFirst } : {}),
        ...(typeof colLast === "number" ? { colLast } : {}),
      },
    };
  }
  return { ok: true, value: { type: "end", id } };
};

export const expectBookmarkBoundaryAttrs = (node: PMNode): BookmarkBoundaryAttrs => {
  const cached = attrsCache.get(node);
  if (cached) {
    return cached;
  }
  const result = readBookmarkBoundaryAttrs(node);
  if (!result.ok) {
    const details = result.issues.map((issue) => `${issue.path}: ${issue.message}`).join("\n");
    panic(`Invalid ProseMirror bookmark boundary attrs:\n${details}`);
  }
  attrsCache.set(node, result.value);
  return result.value;
};
