import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { BookmarkEnd, BookmarkStart, DisplacedByCustomXml } from "../types/document";
import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import type { BookmarkBoundaryAttrs } from "./schema/nodes";

export const isDisplacedByCustomXml = (value: unknown): value is DisplacedByCustomXml =>
  value === "next" || value === "prev";

const attrsCache = new WeakMap<PMNode, BookmarkBoundaryAttrs>();

/**
 * A boundary's `w:displacedByCustomXml`, as a fragment to spread into the model
 * content it rebuilds. Both a start and an end may carry it, so the callers
 * that build either one share this rather than each remembering the attribute.
 */
export const bookmarkBoundaryDisplacement = (
  attrs: BookmarkBoundaryAttrs,
): { displacedByCustomXml?: DisplacedByCustomXml } =>
  attrs.displacedByCustomXml === undefined
    ? {}
    : { displacedByCustomXml: attrs.displacedByCustomXml };

/**
 * The model marker a boundary node carries, at either level.
 *
 * A paragraph, a hyperlink, a revision wrapper and a block container each
 * rebuild the same two elements from the same attributes; spelling that out
 * per caller is how one of them comes to forget `w:colFirst`.
 */
export const bookmarkMarkerFromAttrs = (
  attrs: BookmarkBoundaryAttrs,
): BookmarkStart | BookmarkEnd =>
  attrs.type === "start"
    ? {
        type: "bookmarkStart",
        id: attrs.id,
        name: attrs.name,
        ...(attrs.colFirst !== undefined ? { colFirst: attrs.colFirst } : {}),
        ...(attrs.colLast !== undefined ? { colLast: attrs.colLast } : {}),
        ...bookmarkBoundaryDisplacement(attrs),
      }
    : { type: "bookmarkEnd", id: attrs.id, ...bookmarkBoundaryDisplacement(attrs) };

export const readBookmarkBoundaryAttrs = (
  node: PMNode,
): ReadProseMirrorAttrsResult<BookmarkBoundaryAttrs> => {
  const issues: ProseMirrorAttrIssue[] = [];
  // Either level: the inline atom inside a paragraph, or the block one beside
  // paragraphs. Both carry the same marker, so both read through here.
  if (node.type.name !== "bookmarkBoundary" && node.type.name !== "blockBookmarkBoundary") {
    issues.push({
      path: "bookmarkBoundary.type.name",
      message: `Expected a bookmark boundary, got ${node.type.name}.`,
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

  const displaced = node.attrs["displacedByCustomXml"];
  if (displaced !== undefined && displaced !== null && !isDisplacedByCustomXml(displaced)) {
    issues.push({
      path: "bookmarkBoundary.attrs.displacedByCustomXml",
      message: 'Expected "next" or "prev".',
    });
  }

  if (issues.length > 0 || typeof id !== "number") {
    return { ok: false, issues };
  }
  const displacement = isDisplacedByCustomXml(displaced) ? { displacedByCustomXml: displaced } : {};
  if (type === "start" && typeof name === "string") {
    return {
      ok: true,
      value: {
        type,
        id,
        name,
        ...(typeof colFirst === "number" ? { colFirst } : {}),
        ...(typeof colLast === "number" ? { colLast } : {}),
        ...displacement,
      },
    };
  }
  return { ok: true, value: { type: "end", id, ...displacement } };
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
