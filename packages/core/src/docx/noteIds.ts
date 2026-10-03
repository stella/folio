import { createNumericIdAllocator } from "./numericIdAllocator";

/**
 * The next free endnote id among `existingIds`, a document's endnotes. Ids are
 * unique per document, so the allocator lives for one call: the result
 * depends only on the document, never on what other documents minted.
 */
export const mintEndnoteId = (existingIds: Iterable<number>): number => {
  const allocator = createNumericIdAllocator({ space: "endnote", firstId: 1 });
  allocator.reserve(existingIds);
  return allocator.next();
};
