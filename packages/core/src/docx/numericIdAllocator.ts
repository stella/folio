import { MAX_REVISION_ID } from "@stll/docx-core/model";
import { TaggedError } from "better-result";

class OoxmlIdSpaceExhaustedError extends TaggedError("OoxmlIdSpaceExhaustedError")<{
  message: string;
  space: string;
}> {}

type NumericIdAllocatorOptions = { space: string; firstId: 0 | 1 };

/** Each OOXML id space owns one instance, retaining loaded and minted ids through rollover. */
export const createNumericIdAllocator = ({ space, firstId }: NumericIdAllocatorOptions) => {
  const reserved = new Set<number>();
  let nextId: number = firstId;
  return {
    reserve: (ids: Iterable<number>): void => {
      let max = -1;
      for (const id of ids) {
        if (!Number.isInteger(id) || id < firstId || id > MAX_REVISION_ID) continue;
        reserved.add(id);
        max = Math.max(max, id);
      }
      if (max >= nextId) nextId = max === MAX_REVISION_ID ? firstId : max + 1;
    },
    next: (): number => {
      if (reserved.size >= MAX_REVISION_ID - firstId + 1) {
        throw new OoxmlIdSpaceExhaustedError({
          message: `OOXML ${space} id space is exhausted`,
          space,
        });
      }
      while (reserved.has(nextId)) nextId = nextId === MAX_REVISION_ID ? firstId : nextId + 1;
      const id = nextId;
      reserved.add(id);
      nextId = id === MAX_REVISION_ID ? firstId : id + 1;
      return id;
    },
  };
};
