/** Source of truth shared by relation execution and oracle mutation coverage. */
export const RELATIONS = [
  "directTracked",
  "rejectAll",
  "saveIdempotent",
  "undo",
  "batchSequential",
  "readerStability",
] as const;
export type Relation = (typeof RELATIONS)[number];
