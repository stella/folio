import type { NumberingDefinitions } from "../types/document";
import { createNumericIdAllocator } from "./numericIdAllocator";

/**
 * Every `w:numId` `definitions` holds: the instances it defines and the ids a
 * paragraph or style names without a definition, which a new list never takes.
 */
export const numberingNumIdsTaken = (
  definitions: NumberingDefinitions | null | undefined,
): number[] => [
  ...(definitions?.nums ?? []).map(({ numId }) => numId),
  ...(definitions?.danglingNumIds ?? []),
];

type NumberingKind = "num" | "abstract";

/**
 * Numbering ids (`w:num`, `w:abstractNum`) defined outside the numbering a
 * caller allocates against: the lists the package's other stories define. The
 * ids are package-wide, so a list minted in one story takes none of them.
 */
export type NumberingIdsInUse = {
  numIds: ReadonlySet<number>;
  abstractNumIds: ReadonlySet<number>;
};

export const NO_NUMBERING_IDS_IN_USE: NumberingIdsInUse = {
  numIds: new Set(),
  abstractNumIds: new Set(),
};

type MintNumberingIdOptions = {
  kind: NumberingKind;
  existingIds: Iterable<number>;
};

/** Independent offline packages can allocate deterministically without retaining realm state. */
export const createNumberingIdAllocator = (
  kind: NumberingKind,
  existingIds: Iterable<number>,
): {
  reserve: (ids: Iterable<number>) => void;
  next: () => number;
} => {
  const allocator = createNumericIdAllocator({
    space: kind === "num" ? "numbering instance" : "abstract numbering",
    firstId: kind === "num" ? 1 : 0,
  });
  allocator.reserve(existingIds);
  return allocator;
};

export const mintNumberingId = ({
  kind,
  existingIds,
}: MintNumberingIdOptions): number => {
  return createNumberingIdAllocator(kind, existingIds).next();
};
