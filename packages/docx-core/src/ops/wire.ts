import { cloneModel } from "./modelClone";
import { Result } from "better-result";

import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import type { DocumentOp } from "./types";

const PRESENCE_FIELD = "undefinedFields";
const RESERVED_SEGMENTS = new Set(["__proto__", "constructor", "prototype"]);

/** Capture own undefined fields before an operation enters its JSON journal. */
export const captureDocumentOp = (op: DocumentOp): DocumentOp => {
  const fields = [...(op.undefinedFields ?? [])];
  const known = new Set(fields.map((path) => JSON.stringify(path)));
  const visit = (value: unknown, path: readonly string[]): void => {
    if (typeof value !== "object" || value === null) return;
    for (const [key, child] of Object.entries(value)) {
      if (path.length === 0 && key === PRESENCE_FIELD) continue;
      const next = [...path, key];
      if (child !== undefined) {
        visit(child, next);
        continue;
      }
      const identity = JSON.stringify(next);
      if (known.has(identity)) continue;
      known.add(identity);
      fields.push(next);
    }
  };
  visit(op, []);
  return fields.length === 0 ? op : { ...cloneModel(op), undefinedFields: cloneModel(fields) };
};

/** Restore validated presence metadata without mutating the supplied operation. */
export const restoreDocumentOp = (op: DocumentOp): Result<DocumentOp, DocumentOpRefusal> => {
  if (op === null || typeof op !== "object" || Array.isArray(op))
    return Result.err(
      new DocumentOpRefusal({
        opType: undefined,
        reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        message: "A document operation must be an object.",
      }),
    );
  if (op.undefinedFields === undefined) return Result.ok(op);
  const invalid = (message: string) =>
    Result.err(
      new DocumentOpRefusal({
        opType: op.type,
        reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        message,
      }),
    );
  if (!Array.isArray(op.undefinedFields))
    return invalid("Operation presence metadata is not an array.");
  const restored = cloneModel(op);
  const paths = op.undefinedFields;
  const known = new Set<string>();
  for (const path of paths) {
    if (
      !Array.isArray(path) ||
      path.length === 0 ||
      path.some((part) => typeof part !== "string" || RESERVED_SEGMENTS.has(part))
    )
      return invalid("Operation presence metadata contains an invalid path.");
    const first = path.at(0);
    if (first === "type" || first === PRESENCE_FIELD)
      return invalid("Operation presence metadata cannot replace journal identity.");
    const identity = JSON.stringify(path);
    if (known.has(identity))
      return invalid("Operation presence metadata contains duplicate paths.");
    known.add(identity);
    let parent: unknown = restored;
    for (const part of path.slice(0, -1)) {
      if (typeof parent !== "object" || parent === null || !Object.hasOwn(parent, part))
        return invalid("Operation presence metadata has no owned parent field.");
      parent = Reflect.get(parent, part);
    }
    const field = path.at(-1);
    if (field === undefined || typeof parent !== "object" || parent === null)
      return invalid("Operation presence metadata has no object parent.");
    const value = Reflect.get(parent, field);
    // JSON preserves an undefined array slot as null, while omitting an object field.
    if (
      Object.hasOwn(parent, field) &&
      value !== undefined &&
      !(Array.isArray(parent) && value === null)
    )
      return invalid("Operation presence metadata conflicts with an existing value.");
    if (
      Array.isArray(parent) &&
      (!/^(0|[1-9][0-9]*)$/u.test(field) || Number(field) >= parent.length)
    )
      return invalid("Operation presence metadata names an unavailable array slot.");
    Reflect.set(parent, field, undefined);
  }
  delete restored.undefinedFields;
  return Result.ok(restored);
};
