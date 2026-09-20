import { panic } from "better-result";
import { headingOutlineLevel, type OutlineLevel } from "@stll/docx-core/model";

/**
 * The one shape check over a persisted `outlineLevel` node attr.
 *
 * The attr validator and the readers that go straight to `node.attrs` share
 * it, so a snapshot shape can be accepted in one place and rejected in the
 * other only by deleting this function.
 *
 * It is deliberately strict. ProseMirror's `computeAttrs` copies whatever a
 * stored document holds into the node without validating, so the pre-union
 * shape — a bare `number`, where `9` meant body text — would otherwise arrive
 * as an `outlineLevel` no consumer recognises and every consumer silently
 * ignores. Returning `null` for it is what lets the caller refuse.
 */
export const outlineLevelFromAttrValue = (value: unknown): OutlineLevel | null => {
  if (value === null || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  const kind: unknown = Reflect.get(value, "kind");
  if (kind === "bodyText") {
    return { kind: "bodyText" };
  }
  if (kind !== "heading") {
    return null;
  }
  const level: unknown = Reflect.get(value, "level");
  return typeof level === "number" ? (headingOutlineLevel(level) ?? null) : null;
};

/**
 * A paragraph's stated outline level, read straight from an unvalidated attrs
 * record. `null` is the attr's absent state, which is what the node spec's
 * default stores.
 *
 * A value that is present and not an {@link OutlineLevel} panics rather than
 * being ignored: it can only come from a document written under an attr schema
 * this build does not read, and `yjsDocumentMetadata`'s version gate exists to
 * make that unreachable.
 */
export const readOutlineLevelAttr = (value: unknown): OutlineLevel | null => {
  if (value === undefined || value === null) {
    return null;
  }
  const outlineLevel = outlineLevelFromAttrValue(value);
  if (outlineLevel === null) {
    panic(`Invalid paragraph outlineLevel attr: ${JSON.stringify(value)}`);
  }
  return outlineLevel;
};
