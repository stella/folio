import { createNumericIdAllocator } from "./numericIdAllocator";

type NumberingKind = "num" | "abstract";

type MintNumberingIdOptions = { kind: NumberingKind; existingIds: Iterable<number> };

/** Independent offline packages can allocate deterministically without retaining realm state. */
export const createNumberingIdAllocator = (kind: NumberingKind, existingIds: Iterable<number>) => {
  const allocator = createNumericIdAllocator({
    space: kind === "num" ? "numbering instance" : "abstract numbering",
    firstId: kind === "num" ? 1 : 0,
  });
  allocator.reserve(existingIds);
  return allocator;
};

export const mintNumberingId = ({ kind, existingIds }: MintNumberingIdOptions): number => {
  return createNumberingIdAllocator(kind, existingIds).next();
};
