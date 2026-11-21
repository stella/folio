import { storyParagraphs } from "@stll/docx-core/ops";

import type { BlockContent } from "../types/document";

/**
 * Name the merge each tracked merge's property change belongs to.
 *
 * The default editing engine records a merge's property change, on the
 * paragraph that ends the merged group, under the id of the merge's
 * paragraph-mark deletion. A save gives every record its own `w:id`, so the
 * change also says which revision it belongs to (`resolvesWith`), and the
 * reopened document decides the two together again.
 */
export const linkMergePropertyChanges = (blocks: BlockContent[]): void => {
  const paragraphs = storyParagraphs({ content: blocks });
  const deletions = new Set<number>();
  for (const { paragraph } of paragraphs) {
    const mark = paragraph.pPrMark;
    const deletion =
      mark?.kind === "del" ? mark.info.id : mark?.deletion?.info.id;
    if (deletion !== undefined) deletions.add(deletion);
  }
  if (deletions.size === 0) return;
  for (const { paragraph } of paragraphs) {
    const [change, ...rest] = paragraph.propertyChanges ?? [];
    if (
      change === undefined ||
      change.resolvesWith !== undefined ||
      !deletions.has(change.info.id)
    )
      continue;
    paragraph.propertyChanges = [
      { ...change, resolvesWith: change.info.id },
      ...rest,
    ];
  }
};
