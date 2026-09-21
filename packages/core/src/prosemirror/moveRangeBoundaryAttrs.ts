import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { DisplacedByCustomXml, ParagraphContent } from "../types/document";
import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";

export type MoveRangeMarker = Extract<
  ParagraphContent,
  {
    type: "moveFromRangeStart" | "moveFromRangeEnd" | "moveToRangeStart" | "moveToRangeEnd";
  }
>;

const MOVE_RANGE_START_TYPES = new Set(["moveFromRangeStart", "moveToRangeStart"]);
const MOVE_RANGE_END_TYPES = new Set(["moveFromRangeEnd", "moveToRangeEnd"]);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isNonNegativeSafeInteger = (value: unknown): value is number =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;

const displacementOf = (value: unknown): DisplacedByCustomXml | undefined =>
  value === "next" || value === "prev" ? value : undefined;

/** Read one hidden move-range boundary from an untyped editor or clipboard attribute. */
export const readMoveRangeMarker = (
  value: unknown,
  path = "moveRangeBoundary.attrs.marker",
): ReadProseMirrorAttrsResult<MoveRangeMarker> => {
  const issues: ProseMirrorAttrIssue[] = [];
  if (!isRecord(value)) {
    return { ok: false, issues: [{ path, message: "Expected a move range marker object." }] };
  }

  const type = value["type"];
  const isStart = typeof type === "string" && MOVE_RANGE_START_TYPES.has(type);
  const isEnd = typeof type === "string" && MOVE_RANGE_END_TYPES.has(type);
  if (!isStart && !isEnd) {
    issues.push({ path: `${path}.type`, message: "Expected a move range marker type." });
  }

  const id = value["id"];
  if (!isNonNegativeSafeInteger(id)) {
    issues.push({ path: `${path}.id`, message: "Expected a non-negative safe integer." });
  }

  const displacedByCustomXml = displacementOf(value["displacedByCustomXml"]);
  if (value["displacedByCustomXml"] !== undefined && displacedByCustomXml === undefined) {
    issues.push({ path: `${path}.displacedByCustomXml`, message: 'Expected "next" or "prev".' });
  }

  const name = value["name"];
  const author = value["author"];
  const date = value["date"];
  const colFirst = value["colFirst"];
  const colLast = value["colLast"];
  if (isStart) {
    if (typeof name !== "string" || name.trim().length === 0) {
      issues.push({ path: `${path}.name`, message: "Expected a non-empty string." });
    }
    if (typeof author !== "string" || author.trim().length === 0) {
      issues.push({ path: `${path}.author`, message: "Expected a non-empty string." });
    }
    if (date !== undefined && typeof date !== "string") {
      issues.push({ path: `${path}.date`, message: "Expected a string." });
    }
    for (const [key, column] of [
      ["colFirst", colFirst],
      ["colLast", colLast],
    ] as const) {
      if (column !== undefined && !isNonNegativeSafeInteger(column)) {
        issues.push({ path: `${path}.${key}`, message: "Expected a non-negative safe integer." });
      }
    }
  }

  if (issues.length > 0 || !isNonNegativeSafeInteger(id) || (!isStart && !isEnd)) {
    return { ok: false, issues };
  }
  const displacement = displacedByCustomXml === undefined ? {} : { displacedByCustomXml };
  if (type === "moveFromRangeEnd" || type === "moveToRangeEnd") {
    return { ok: true, value: { type, id, ...displacement } };
  }
  if (
    (type !== "moveFromRangeStart" && type !== "moveToRangeStart") ||
    typeof name !== "string" ||
    typeof author !== "string"
  ) {
    return { ok: false, issues };
  }
  return {
    ok: true,
    value: {
      type,
      id,
      name,
      author,
      ...(typeof date === "string" ? { date } : {}),
      ...(typeof colFirst === "number" ? { colFirst } : {}),
      ...(typeof colLast === "number" ? { colLast } : {}),
      ...displacement,
    },
  };
};

const attrsCache = new WeakMap<PMNode, MoveRangeMarker>();

export const readMoveRangeBoundaryAttrs = (
  node: PMNode,
): ReadProseMirrorAttrsResult<MoveRangeMarker> => {
  const result = readMoveRangeMarker(node.attrs["marker"]);
  if (node.type.name === "moveRangeBoundary") {
    return result;
  }
  const issue = {
    path: "moveRangeBoundary.type.name",
    message: `Expected moveRangeBoundary, got ${node.type.name}.`,
  };
  return result.ok
    ? { ok: false, issues: [issue] }
    : { ok: false, issues: [issue, ...result.issues] };
};

export const expectMoveRangeBoundaryAttrs = (node: PMNode): MoveRangeMarker => {
  const cached = attrsCache.get(node);
  if (cached) {
    return cached;
  }
  const result = readMoveRangeBoundaryAttrs(node);
  if (!result.ok) {
    const details = result.issues.map((issue) => `${issue.path}: ${issue.message}`).join("\n");
    panic(`Invalid ProseMirror move range boundary attrs:\n${details}`);
  }
  attrsCache.set(node, result.value);
  return result.value;
};
