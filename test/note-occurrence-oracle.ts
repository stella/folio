/** Compare independently minted identities while retaining occurrence equivalence classes. */
export const normalizeNoteOccurrenceIds = (value: unknown): unknown => {
  const ids = new Map<string, number>();
  const visit = (entry: unknown): unknown => {
    if (Array.isArray(entry)) return entry.map(visit);
    if (typeof entry !== "object" || entry === null) return entry;
    return Object.fromEntries(
      Object.entries(entry).map(([key, child]) => {
        if (key !== "occurrenceId" || typeof child !== "string") return [key, visit(child)];
        let id = ids.get(child);
        if (id === undefined) {
          id = ids.size;
          ids.set(child, id);
        }
        return [key, id];
      }),
    );
  };
  return visit(value);
};
