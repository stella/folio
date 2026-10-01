import { createNumericIdAllocator } from "./numericIdAllocator";

type NumberingKind = "num" | "abstract";
let numIds: ReturnType<typeof createNumericIdAllocator> | undefined;
let abstractIds: ReturnType<typeof createNumericIdAllocator> | undefined;

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
  const allocator =
    kind === "num"
      ? (numIds ??= createNumericIdAllocator({ space: "numbering instance", firstId: 1 }))
      : (abstractIds ??= createNumericIdAllocator({ space: "abstract numbering", firstId: 0 }));
  allocator.reserve(existingIds);
  return allocator.next();
};
