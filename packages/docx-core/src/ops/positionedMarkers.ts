/**
 * Bookmark markers a table holds beside its rows, or a row beside its cells
 * ({@link PositionedBookmarkMarker}), kept in place when rows or cells come
 * and go. A marker's index counts the siblings before it, so an edit that
 * adds or removes siblings moves it: a start stays before the sibling that
 * followed it, as the bookmark it opens selects that sibling, and an end
 * stays after the sibling that preceded it. Captured XML stays before the
 * next surviving sibling. Where a marker's sibling is gone, the marker
 * stands where it was, among the siblings that remain.
 */
import type {
  PositionedBookmarkMarker,
  Table,
  TableCell,
  TablePreservedMarkup,
  TableRow,
} from "../model/document";
import { structurallyEqual } from "./equality";
import { idKey, paragraphIdsIn } from "./ids";

/**
 * For each sibling after an edit, the index of the sibling before it that it
 * continues, or `null` for a new one. A sibling continues the latest one
 * whose paragraphs it holds: a row that takes over a removed row's merged
 * cells holds both rows' paragraphs, and it is the later one.
 */
const originsOf = (
  before: readonly (TableRow | TableCell)[],
  after: readonly (TableRow | TableCell)[],
): (number | null)[] => {
  const owner = new Map<string, number>();
  for (const [index, sibling] of before.entries())
    for (const id of paragraphIdsIn(sibling)) owner.set(idKey(id), index);
  return after.map((sibling) => {
    let origin: number | null = null;
    for (const id of paragraphIdsIn(sibling)) {
      const index = owner.get(idKey(id));
      if (index !== undefined && (origin === null || index > origin))
        origin = index;
    }
    return origin;
  });
};

const beforeFollowingSibling = (
  index: number,
  origins: readonly (number | null)[],
): number => {
  const following = origins.findIndex(
    (origin) => origin !== null && origin >= index,
  );
  return following === -1 ? origins.length : following;
};

/** `markers`, indexed among the siblings `origins` describes. */
const relocated = (
  markers: PositionedBookmarkMarker[],
  origins: readonly (number | null)[],
): PositionedBookmarkMarker[] => {
  let changed = false;
  const nextMarkers = markers.map((positioned) => {
    const { index } = positioned;
    let next: number;
    if (positioned.marker.type === "bookmarkStart") {
      next = beforeFollowingSibling(index, origins);
    } else
      next =
        origins.findLastIndex((origin) => origin !== null && origin < index) +
        1;
    if (next === index) return positioned;
    changed = true;
    return { ...positioned, index: next };
  });
  return changed ? nextMarkers : markers;
};

/** Preserve a captured child's gap before the next surviving sibling. */
const relocateCaptured = (
  preserved: TablePreservedMarkup,
  origins: readonly (number | null)[],
): TablePreservedMarkup => {
  if (preserved.children === undefined) return preserved;
  let changed = false;
  const children = preserved.children.map((child) => {
    const index = beforeFollowingSibling(child.index, origins);
    if (index === child.index) return child;
    changed = true;
    return { ...child, index };
  });
  return changed ? { ...preserved, children } : preserved;
};

/**
 * `after`, with its bookmark markers and captured XML moved to the gaps
 * where the edit leaves them beside table rows and row cells.
 */
export const relocatePositionedMarkers = (
  before: Table,
  after: Table,
): Table => {
  const rowOrigins = originsOf(before.rows, after.rows);
  let rowsChanged = false;
  const rows = after.rows.map((row, index) => {
    const origin = rowOrigins[index] ?? null;
    const source = origin === null ? undefined : before.rows[origin];
    if (
      source === undefined ||
      (row.bookmarks === undefined && row.preserved === undefined)
    )
      return row;
    const cellOrigins = originsOf(source.cells, row.cells);
    // Exact row replacements can state their own gaps. Re-address inherited
    // captures against the cells they originally stood beside.
    const markers =
      row.bookmarks !== undefined &&
      structurallyEqual(row.bookmarks, source.bookmarks)
        ? relocated(row.bookmarks, cellOrigins)
        : row.bookmarks;
    const preserved =
      row.preserved !== undefined &&
      structurallyEqual(row.preserved, source.preserved)
        ? relocateCaptured(row.preserved, cellOrigins)
        : row.preserved;
    if (preserved === row.preserved && markers === row.bookmarks) return row;
    rowsChanged = true;
    return {
      ...row,
      ...(markers === undefined ? {} : { bookmarks: markers }),
      ...(preserved === undefined ? {} : { preserved }),
    };
  });
  const tableMarkers =
    after.bookmarks === undefined
      ? undefined
      : relocated(after.bookmarks, rowOrigins);
  const preserved =
    after.preserved === undefined
      ? undefined
      : relocateCaptured(after.preserved, rowOrigins);
  if (
    !rowsChanged &&
    tableMarkers === after.bookmarks &&
    preserved === after.preserved
  )
    return after;
  return {
    ...after,
    rows,
    ...(tableMarkers === undefined ? {} : { bookmarks: tableMarkers }),
    ...(preserved === undefined ? {} : { preserved }),
  };
};

/** Whether the inverse needs the original row and cell gaps. */
export const holdsPositionedMarkup = (table: Table): boolean =>
  table.preserved?.children !== undefined ||
  table.bookmarks !== undefined ||
  table.rows.some(
    (row) =>
      row.bookmarks !== undefined || row.preserved?.children !== undefined,
  );
