import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import { isDisplacedByCustomXml } from "./bookmarkBoundaryAttrs";
import {
  RANGE_END_FOR_START,
  type RangeAnchorAttrs,
  type RangeMarker,
} from "./emptyRangeAnchor";

const attrsCache = new WeakMap<PMNode, RangeAnchorAttrs>();

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const isMarkerId = (value: unknown): value is number =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;

const isRangeStartType = (
  value: unknown,
): value is keyof typeof RANGE_END_FOR_START =>
  typeof value === "string" && Object.hasOwn(RANGE_END_FOR_START, value);

const validateRangeMarkerAttrs = (
  marker: Record<string, unknown>,
  path: string,
  issues: ProseMirrorAttrIssue[],
): void => {
  if (!isMarkerId(marker["id"])) {
    issues.push({
      path: `${path}.id`,
      message: "Expected a non-negative integer.",
    });
  }
  const displacedByCustomXml = marker["displacedByCustomXml"];
  if (
    displacedByCustomXml !== undefined &&
    !isDisplacedByCustomXml(displacedByCustomXml)
  ) {
    issues.push({
      path: `${path}.displacedByCustomXml`,
      message: 'Expected "next" or "prev".',
    });
  }
};

const validateMoveRangeStartAttrs = (
  marker: Record<string, unknown>,
  path: string,
  issues: ProseMirrorAttrIssue[],
): void => {
  for (const required of ["name", "author"]) {
    if (
      typeof marker[required] !== "string" ||
      marker[required].trim() === ""
    ) {
      issues.push({
        path: `${path}.${required}`,
        message: "Expected a non-empty string on a move range start.",
      });
    }
  }
  const date = marker["date"];
  if (date !== undefined && typeof date !== "string") {
    issues.push({
      path: `${path}.date`,
      message: "Expected a string.",
    });
  }
  for (const column of ["colFirst", "colLast"]) {
    const value = marker[column];
    if (
      value !== undefined &&
      (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0)
    ) {
      issues.push({
        path: `${path}.${column}`,
        message: "Expected a non-negative integer.",
      });
    }
  }
};

/** An ordered, nonempty balanced sequence from one zero-width source cluster. */
export const readRangeAnchorValues = (
  markers: unknown,
): ReadProseMirrorAttrsResult<RangeAnchorAttrs> => {
  const issues: ProseMirrorAttrIssue[] = [];
  if (!Array.isArray(markers) || markers.length < 2) {
    return {
      ok: false,
      issues: [
        {
          path: "rangeAnchor.attrs.markers",
          message: "Expected a nonempty range-marker sequence.",
        },
      ],
    };
  }
  const validMarkers = markers.filter(
    (marker, index): marker is RangeMarker => {
      const path = `rangeAnchor.attrs.markers[${index}]`;
      if (
        !isRecord(marker) ||
        !(
          isRangeStartType(marker["type"]) ||
          marker["type"] === "commentRangeEnd" ||
          marker["type"] === "moveFromRangeEnd" ||
          marker["type"] === "moveToRangeEnd"
        )
      ) {
        issues.push({ path, message: "Expected a range marker." });
        return false;
      }
      const before = issues.length;
      validateRangeMarkerAttrs(marker, path, issues);
      if (
        marker["type"] === "moveFromRangeStart" ||
        marker["type"] === "moveToRangeStart"
      )
        validateMoveRangeStartAttrs(marker, path, issues);
      return issues.length === before;
    },
  );
  if (validMarkers.length !== markers.length) return { ok: false, issues };
  const open = new Map<string, number>();
  for (const [index, marker] of validMarkers.entries()) {
    const path = `rangeAnchor.attrs.markers[${index}]`;
    if (
      marker.type === "commentRangeStart" ||
      marker.type === "moveFromRangeStart" ||
      marker.type === "moveToRangeStart"
    ) {
      const key = `${RANGE_END_FOR_START[marker.type]}:${marker.id}`;
      open.set(key, (open.get(key) ?? 0) + 1);
    } else {
      const key = `${marker.type}:${marker.id}`;
      const count = open.get(key) ?? 0;
      if (count === 0)
        issues.push({ path, message: "Expected a preceding matching start." });
      else if (count === 1) open.delete(key);
      else open.set(key, count - 1);
    }
  }
  if (open.size > 0)
    issues.push({
      path: "rangeAnchor.attrs.markers",
      message: "Expected every start to have a matching end.",
    });
  // Copied move boundaries resolve independently; their existing atom owns one
  // pair. Only comment points may share an indivisible interleaved cluster.
  if (
    markers.length !== 2 &&
    validMarkers.some(
      (marker) =>
        marker.type !== "commentRangeStart" &&
        marker.type !== "commentRangeEnd",
    )
  )
    issues.push({
      path: "rangeAnchor.attrs.markers",
      message: "Expected a single move pair or a comment-only cluster.",
    });
  return issues.length > 0
    ? { ok: false, issues }
    : { ok: true, value: { markers: validMarkers } };
};

/**
 * The anchor's attrs, checked before a save can call string methods or emit
 * invalid range-marker XML.
 */
export const readRangeAnchorAttrs = (
  node: PMNode,
): ReadProseMirrorAttrsResult<RangeAnchorAttrs> => {
  const result = readRangeAnchorValues(node.attrs["markers"]);
  if (node.type.name === "rangeAnchor") {
    return result;
  }
  return {
    ok: false,
    issues: [
      {
        path: "rangeAnchor.type.name",
        message: `Expected rangeAnchor, got ${node.type.name}.`,
      },
      ...(result.ok ? [] : result.issues),
    ],
  };
};

export const expectRangeAnchorAttrs = (node: PMNode): RangeAnchorAttrs => {
  const cached = attrsCache.get(node);
  if (cached) {
    return cached;
  }
  const result = readRangeAnchorAttrs(node);
  if (!result.ok) {
    const details = result.issues
      .map((issue) => `${issue.path}: ${issue.message}`)
      .join("\n");
    panic(`Invalid ProseMirror range anchor attrs:\n${details}`);
  }
  attrsCache.set(node, result.value);
  return result.value;
};
