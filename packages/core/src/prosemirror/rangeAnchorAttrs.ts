import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import { isDisplacedByCustomXml } from "./bookmarkBoundaryAttrs";
import { RANGE_END_FOR_START, type RangeAnchorAttrs } from "./emptyRangeAnchor";

const attrsCache = new WeakMap<PMNode, RangeAnchorAttrs>();

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const isMarkerId = (value: unknown): value is number =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;

const isRangeStartType = (value: unknown): value is keyof typeof RANGE_END_FOR_START =>
  typeof value === "string" && value in RANGE_END_FOR_START;

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
  if (displacedByCustomXml !== undefined && !isDisplacedByCustomXml(displacedByCustomXml)) {
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
    if (typeof marker[required] !== "string" || marker[required].trim() === "") {
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

/**
 * The anchor's two markers, checked for every field the DOCX serializer reads.
 * This is an editor boundary: malformed attrs must be rejected before a save
 * can call string methods or emit invalid range-marker XML.
 */
export const readRangeAnchorAttrs = (
  node: PMNode,
): ReadProseMirrorAttrsResult<RangeAnchorAttrs> => {
  const issues: ProseMirrorAttrIssue[] = [];
  if (node.type.name !== "rangeAnchor") {
    issues.push({
      path: "rangeAnchor.type.name",
      message: `Expected rangeAnchor, got ${node.type.name}.`,
    });
  }
  const start = node.attrs["start"];
  const end = node.attrs["end"];
  if (!isRecord(start) || !isRangeStartType(start["type"])) {
    issues.push({
      path: "rangeAnchor.attrs.start",
      message: `Expected one of ${Object.keys(RANGE_END_FOR_START).join(", ")}.`,
    });
    return { ok: false, issues };
  }
  const expectedEndType = RANGE_END_FOR_START[start["type"]];
  if (!isRecord(end) || end["type"] !== expectedEndType) {
    issues.push({
      path: "rangeAnchor.attrs.end",
      message: `Expected ${expectedEndType}.`,
    });
  }
  validateRangeMarkerAttrs(start, "rangeAnchor.attrs.start", issues);
  if (isRecord(end)) {
    validateRangeMarkerAttrs(end, "rangeAnchor.attrs.end", issues);
    if (end["id"] !== start["id"]) {
      issues.push({
        path: "rangeAnchor.attrs.end.id",
        message: "Expected the id its start carries.",
      });
    }
  }
  if (start["type"] !== "commentRangeStart") {
    validateMoveRangeStartAttrs(start, "rangeAnchor.attrs.start", issues);
  }

  if (issues.length > 0) {
    return { ok: false, issues };
  }
  // SAFETY: the checks above are exactly the discriminator, the pairing and the
  // required fields that tell `RangeAnchorAttrs`'s branches apart; nothing in
  // ProseMirror's untyped attrs bag can carry that narrowing for us.
  return { ok: true, value: node.attrs as unknown as RangeAnchorAttrs };
};

export const expectRangeAnchorAttrs = (node: PMNode): RangeAnchorAttrs => {
  const cached = attrsCache.get(node);
  if (cached) {
    return cached;
  }
  const result = readRangeAnchorAttrs(node);
  if (!result.ok) {
    const details = result.issues.map((issue) => `${issue.path}: ${issue.message}`).join("\n");
    panic(`Invalid ProseMirror range anchor attrs:\n${details}`);
  }
  attrsCache.set(node, result.value);
  return result.value;
};
