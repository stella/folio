/** Test projection patch: reference unchanged top-level blocks, emit changed blocks. */
import type { Document } from "../../../packages/docx-core/src/model/document";

export const sameValue = (left: unknown, right: unknown): boolean => {
  if (left === right) return true;
  if (
    typeof left !== "object" ||
    left === null ||
    typeof right !== "object" ||
    right === null ||
    Array.isArray(left) !== Array.isArray(right)
  )
    return false;
  const keys = Object.keys(left);
  return (
    keys.length === Object.keys(right).length &&
    keys.every(
      (key) =>
        Object.hasOwn(right, key) && sameValue(Reflect.get(left, key), Reflect.get(right, key)),
    )
  );
};

export const projectionPatch = (before: Document, after: Document) => {
  const oldBlocks = before.package.document.content;
  const byId = new Map(
    oldBlocks.flatMap((block, index) =>
      block.type === "paragraph" && block.paraId !== undefined
        ? [[block.paraId, index] satisfies [string, number]]
        : [],
    ),
  );
  const content = after.package.document.content.map((block, index) => {
    if (sameValue(oldBlocks.at(index), block)) return index;
    if (block.type === "paragraph" && block.paraId !== undefined) {
      const candidate = byId.get(block.paraId);
      if (candidate !== undefined && sameValue(oldBlocks.at(candidate), block)) return candidate;
      return block;
    }
    const candidate = oldBlocks.findIndex((old) => sameValue(old, block));
    return candidate < 0 ? block : candidate;
  });
  if (after.package.document.sections === undefined) return { content };
  return {
    content,
    sectionMetadata: after.package.document.sections.map(
      ({ content: ignoredContent, ...metadata }) => {
        void ignoredContent;
        return metadata;
      },
    ),
  };
};
