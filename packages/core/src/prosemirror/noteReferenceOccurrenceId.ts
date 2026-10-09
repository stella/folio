/** Occurrence identities are opaque, nonblank attribution tokens. */
export const isNoteReferenceOccurrenceId = (value: unknown): value is string =>
  typeof value === "string" && value.trim().length > 0;
