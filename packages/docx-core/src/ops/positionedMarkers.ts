/**
 * Bookmark markers a table holds beside its rows, or a row beside its cells
 * ({@link PositionedBookmarkMarker}), kept in place when rows or cells come
 * and go. A marker's index counts the siblings before it, so an edit that
 * adds or removes siblings moves it: a start stays before the sibling that
 * followed it, as the bookmark it opens selects that sibling, and an end
 * stays after the sibling that preceded it. Where that sibling is gone, the
 * marker stands where it was, among the siblings that remain.
 */
import type {
  PositionedBookmarkMarker,
  Table,
  TableCell,
  TableRow,
} from "../model/document";
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

/** `markers`, indexed among the siblings `origins` describes. */
const relocated = (
  markers: readonly PositionedBookmarkMarker[],
  origins: readonly (number | null)[],
): PositionedBookmarkMarker[] =>
  markers.map((positioned) => {
    const { index } = positioned;
    let next: number;
    if (positioned.marker.type === "bookmarkStart") {
      const following = origins.findIndex(
        (origin) => origin !== null && origin >= index,
      );
      next = following === -1 ? origins.length : following;
    } else
      next =
        origins.findLastIndex((origin) => origin !== null && origin < index) +
        1;
    return next === index ? positioned : { ...positioned, index: next };
  });

const sameIndices = (
  markers: readonly PositionedBookmarkMarker[],
  next: readonly PositionedBookmarkMarker[],
): boolean => markers.every((marker, index) => marker === next[index]);

/**
 * `after`, the table an edit made of `before`, with the markers `before`
 * holds beside its rows, and those each continuing row holds beside its
 * cells, moved to where the edit leaves them.
 */
export const relocatePositionedMarkers = (
  before: Table,
  after: Table,
): Table => {
  const rowOrigins = originsOf(before.rows, after.rows);
  let rowsChanged = false;
  const rows = after.rows.map((row, index) => {
    const origin = rowOrigins[index];
    const source = origin === null ? undefined : before.rows[origin];
    if (row.bookmarks === undefined || source === undefined) return row;
    const markers = relocated(
      row.bookmarks,
      originsOf(source.cells, row.cells),
    );
    if (sameIndices(row.bookmarks, markers)) return row;
    rowsChanged = true;
    return { ...row, bookmarks: markers };
  });
  const tableMarkers =
    after.bookmarks === undefined
      ? undefined
      : relocated(after.bookmarks, rowOrigins);
  const tableChanged =
    after.bookmarks !== undefined &&
    tableMarkers !== undefined &&
    !sameIndices(after.bookmarks, tableMarkers);
  if (!rowsChanged && !tableChanged) return after;
  return {
    ...after,
    rows,
    ...(tableMarkers === undefined ? {} : { bookmarks: tableMarkers }),
  };
};

/** Whether `table` holds a marker beside its rows or beside a row's cells. */
export const holdsPositionedMarkers = (table: Table): boolean =>
  table.bookmarks !== undefined ||
  table.rows.some((row) => row.bookmarks !== undefined);
