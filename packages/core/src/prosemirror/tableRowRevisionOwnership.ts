type TableRowRevisionCompanionOptions = {
  row: {
    id: number;
    author?: string | undefined;
    date?: string | null | undefined;
  };
  candidate: { id: unknown; author: unknown; date: unknown };
};

/**
 * A row's structural and inline marks can share a revision id. Word can also
 * mint separate ids for them, so a matching author and known date count too.
 */
export const isTableRowRevisionCompanion = ({
  row,
  candidate,
}: TableRowRevisionCompanionOptions): boolean =>
  candidate.id === row.id ||
  (row.author !== undefined &&
    candidate.author === row.author &&
    typeof row.date === "string" &&
    row.date.length > 0 &&
    candidate.date === row.date);
