import { FolioDocxReviewer } from "@stll/folio-core/server";
import { getCachedNumberingMap } from "@stll/folio-core/docx/numberingParser";
import { canonicalJson } from "@stll/folio-core/utils/canonicalJson";

/**
 * A resolved redline may retain both numbering definitions and move its list
 * onto a fresh numId. Compare the effective levels and the identity of each
 * numbering sequence, including every level that can restart a nested list.
 */
export const equivalentNumberingAliases = async (
  originalBuffer: ArrayBuffer,
  resolvedBuffer: ArrayBuffer,
): Promise<boolean> => {
  const original = await FolioDocxReviewer.fromBuffer(originalBuffer);
  const resolved = await FolioDocxReviewer.fromBuffer(resolvedBuffer);
  const originalDefinitions = original.toDocument().package.numbering;
  const resolvedDefinitions = resolved.toDocument().package.numbering;
  if (!originalDefinitions || !resolvedDefinitions) return false;
  const originalNumbering = getCachedNumberingMap(originalDefinitions);
  const resolvedNumbering = getCachedNumberingMap(resolvedDefinitions);
  const originalStories = original.listStories();
  const resolvedStories = resolved.listStories();
  if (originalStories.length !== resolvedStories.length) return false;

  const forward = new Map<number, number>();
  const reverse = new Map<number, number>();
  let remapped = false;
  for (const [storyIndex, { handle }] of originalStories.entries()) {
    const counterpart = resolvedStories[storyIndex];
    if (!counterpart || JSON.stringify(handle) !== JSON.stringify(counterpart.handle)) return false;
    const before = original.snapshotStory(handle)?.blocks;
    const after = resolved.snapshotStory(handle)?.blocks;
    if (!before || !after || before.length !== after.length) return false;
    for (const [index, block] of before.entries()) {
      const other = after[index];
      if (!other || block.kind !== other.kind || block.text !== other.text) return false;
      const source = block.listReference;
      const target = other.listReference;
      if (!source && !target) continue;
      if (!source || !target || source.level !== target.level) return false;
      if (!originalNumbering.getLevel(source.numId, source.level)) return false;
      if (!resolvedNumbering.getLevel(target.numId, target.level)) return false;
      const mapped = forward.get(source.numId);
      const reverseMapped = reverse.get(target.numId);
      if (
        (mapped !== undefined && mapped !== target.numId) ||
        (reverseMapped !== undefined && reverseMapped !== source.numId)
      )
        return false;
      forward.set(source.numId, target.numId);
      reverse.set(target.numId, source.numId);
      remapped ||= source.numId !== target.numId;
    }
  }
  if (!remapped) return false;

  for (const [sourceNumId, targetNumId] of forward) {
    for (let level = 0; level <= 8; level += 1) {
      if (
        canonicalJson(originalNumbering.getLevel(sourceNumId, level)) !==
        canonicalJson(resolvedNumbering.getLevel(targetNumId, level))
      )
        return false;
    }
  }
  return true;
};
