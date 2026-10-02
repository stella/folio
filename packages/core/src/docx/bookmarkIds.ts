import type { Node } from "prosemirror-model";

import { createNumericIdAllocator } from "./numericIdAllocator";

let allocator: ReturnType<typeof createNumericIdAllocator> | undefined;
const getAllocator = () =>
  (allocator ??= createNumericIdAllocator({ space: "bookmark", firstId: 0 }));

export const reserveBookmarkIds = (ids: Iterable<number>): void => getAllocator().reserve(ids);
export const mintBookmarkId = (): number => getAllocator().next();

export const createBookmarkIdAllocator = (existingIds: Iterable<number>) => {
  const scopedAllocator = createNumericIdAllocator({ space: "bookmark", firstId: 0 });
  scopedAllocator.reserve(existingIds);
  return scopedAllocator;
};

export const reserveProseBookmarkIds = (doc: Node): void => {
  const ids: number[] = [];
  doc.descendants((node) => {
    const bookmarks: unknown = node.attrs["bookmarks"];
    if (Array.isArray(bookmarks)) {
      for (const bookmark of bookmarks) {
        if (
          typeof bookmark === "object" &&
          bookmark !== null &&
          "id" in bookmark &&
          typeof bookmark.id === "number"
        )
          ids.push(bookmark.id);
      }
    }
    if (node.type.name === "bookmarkBoundary" || node.type.name === "blockBookmarkBoundary") {
      const id: unknown = node.attrs["id"];
      if (typeof id === "number") ids.push(id);
    }
  });
  reserveBookmarkIds(ids);
};
