import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { ProseMirrorAttrIssue, ReadProseMirrorAttrsResult } from "./attrs";
import { RANGE_END_FOR_START, type RangeAnchorAttrs } from "./emptyRangeAnchor";

const attrsCache = new WeakMap<PMNode, RangeAnchorAttrs>();

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

const isMarkerId = (value: unknown): value is number =>
  typeof value === "number" && Number.isInteger(value) && value >= 0;

const isRangeStartType = (value: unknown): value is keyof typeof RANGE_END_FOR_START =>
  typeof value === "string" && value in RANGE_END_FOR_START;

/**
 * The anchor's two markers, checked for the invariants that make it an anchor
 * at all: a start, the end that closes that start, and one `w:id` between them.
 *
 * The fields beyond those are the model's own and are not restated here; a
 * move start's `w:name` and `w:author` are, because the schema requires them
 * and a marker written without either is a package Word refuses.
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
  if (!isMarkerId(start["id"])) {
    issues.push({
      path: "rangeAnchor.attrs.start.id",
      message: "Expected a non-negative integer.",
    });
  }
  if (isRecord(end) && end["id"] !== start["id"]) {
    issues.push({
      path: "rangeAnchor.attrs.end.id",
      message: "Expected the id its start carries.",
    });
  }
  if (start["type"] !== "commentRangeStart") {
    for (const required of ["name", "author"]) {
      if (typeof start[required] !== "string" || start[required] === "") {
        issues.push({
          path: `rangeAnchor.attrs.start.${required}`,
          message: "Expected a non-empty string on a move range start.",
        });
      }
    }
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
