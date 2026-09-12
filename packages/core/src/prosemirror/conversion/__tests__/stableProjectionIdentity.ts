import type { Node as PMNode } from "prosemirror-model";

const stableId = (ids: Map<string, number>, value: unknown, family: string): unknown => {
  if (typeof value !== "string") {
    return value;
  }
  const existing = ids.get(value);
  if (existing !== undefined) {
    return `<${family}:${String(existing)}>`;
  }
  const ordinal = ids.size;
  ids.set(value, ordinal);
  return `<${family}:${String(ordinal)}>`;
};

/**
 * Removes allocation-specific values while retaining every equality relation
 * among group and anchor identifiers.
 */
export const stableProjectionIdentity = (document: PMNode): string => {
  const groupIds = new Map<string, number>();
  const anchorIds = new Map<string, number>();

  return JSON.stringify(document.toJSON(), (key, value: unknown) => {
    if (key === "_docxGroupId") {
      return stableId(groupIds, value, "group");
    }
    if (key === "_docxAnchorId") {
      return stableId(anchorIds, value, "anchor");
    }
    return value;
  });
};
