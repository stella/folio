import { createNumericIdAllocator } from "./numericIdAllocator";

let endnoteIds: ReturnType<typeof createNumericIdAllocator> | undefined;

export const mintEndnoteId = (existingIds: Iterable<number>): number => {
  const allocator = (endnoteIds ??= createNumericIdAllocator({ space: "endnote", firstId: 1 }));
  allocator.reserve(existingIds);
  return allocator.next();
};
