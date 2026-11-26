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
  SdtProperties,
  Table,
  TableCell,
  TableContentCarrier,
  TablePreservedMarkup,
  TableRow,
} from "../model/document";
import { structurallyEqual } from "./equality";
import { idKey, paragraphIdsIn } from "./ids";

type Wrapped = {
  carrierStack?: TableContentCarrier[];
  contentControls?: SdtProperties[];
};

const carriers = (item: Wrapped | undefined): TableContentCarrier[] =>
  item?.carrierStack ??
  (item?.contentControls ?? []).map((properties) => ({
    type: "sdt",
    properties,
  }));

const sameCarrier = (left: TableContentCarrier, right: TableContentCarrier) =>
  left.type === "customXml"
    ? right.type === "customXml" && left.wrapper.id === right.wrapper.id
    : right.type === "sdt" &&
      structurallyEqual(left.properties, right.properties);

const sharedDepth = (
  left: readonly TableContentCarrier[],
  right: readonly TableContentCarrier[],
): number => {
  let depth = 0;
  for (const [index, carrier] of left.entries()) {
    const peer = right[index];
    if (peer === undefined || !sameCarrier(carrier, peer)) break;
    depth += 1;
  }
  return depth;
};

type InheritCarriersOptions<T> = {
  siblings: T[];
  origins: readonly (number | null)[];
  preserveCarriers?: ReadonlySet<T> | undefined;
};
/** A new sibling between two members of one wrapper stays inside it. */
const inheritCarriers = <T extends TableRow | TableCell>({
  siblings,
  origins,
  preserveCarriers,
}: InheritCarriersOptions<T>): T[] => {
  if (!origins.includes(null)) return siblings;
  let changed = false;
  const next = [...siblings];
  for (let start = 0; start < siblings.length; start++) {
    if (origins[start] !== null) continue;
    const left = carriers(siblings[start - 1]);
    let end = start + 1;
    while (origins[end] === null) end += 1;
    const right = carriers(siblings[end]);
    const depth = sharedDepth(left, right);
    if (depth > 0) {
      const inherited = left.slice(0, depth);
      for (let index = start; index < end; index++) {
        const sibling = siblings[index];
        if (sibling === undefined || preserveCarriers?.has(sibling)) continue;
        const existing = carriers(sibling);
        const shared = sharedDepth(inherited, existing);
        if (shared === inherited.length) continue;
        const carrierStack = [...inherited, ...existing.slice(shared)];
        const contentControls = carrierStack.flatMap((carrier) =>
          carrier.type === "sdt" ? [carrier.properties] : [],
        );
        changed = true;
        next[index] = {
          ...sibling,
          ...(contentControls.length === 0 ? {} : { contentControls }),
          // Preserve the model's legacy SDT-only representation when the
          // neighbors have no recorded mixed order.
          ...(sibling.carrierStack === undefined &&
          siblings[start - 1]?.carrierStack === undefined &&
          siblings[end]?.carrierStack === undefined
            ? {}
            : { carrierStack }),
        };
      }
    }
    start = end - 1;
  }
  return changed ? next : siblings;
};

/**
 * For each sibling after an edit, the index of the sibling before it that it
 * continues, or `null` for a new one. A sibling continues the latest one
 * whose paragraphs it holds: a row that takes over a removed row's merged
 * cells holds both rows' paragraphs, and it is the later one.
 */
type OriginsOptions<T extends TableRow | TableCell> = {
  before: readonly T[];
  after: readonly T[];
  fresh?: (item: T) => boolean;
};
const originsOf = <T extends TableRow | TableCell>({
  before,
  after,
  fresh,
}: OriginsOptions<T>): (number | null)[] => {
  const owner = new Map<string, number>();
  for (const [index, sibling] of before.entries()) {
    if (fresh?.(sibling)) continue;
    for (const id of paragraphIdsIn(sibling)) owner.set(idKey(id), index);
  }
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

type GapDestinations = {
  origins: readonly (number | null)[];
  siblings: readonly Wrapped[];
};

/** Keep positioned children beside siblings that share their wrapper. */
const relocatedIndex = (
  item: Wrapped & { index: number },
  { origins, siblings }: GapDestinations,
  sticksLeft: boolean,
): number => {
  const { index } = item;
  const right = beforeFollowingSibling(index, origins);
  const left =
    origins.findLastIndex((origin) => origin !== null && origin < index) + 1;
  if (left === right) return left;
  const wrapper = carriers(item);
  if (wrapper.length > 0) {
    const leftDepth = sharedDepth(wrapper, carriers(siblings[left - 1]));
    const rightDepth = sharedDepth(wrapper, carriers(siblings[right]));
    if (leftDepth !== rightDepth) return leftDepth > rightDepth ? left : right;
  }
  return sticksLeft ? left : right;
};

/** `markers`, indexed among the siblings `origins` describes. */
const relocated = (
  markers: PositionedBookmarkMarker[],
  destinations: GapDestinations,
): PositionedBookmarkMarker[] => {
  let changed = false;
  const nextMarkers = markers.map((positioned) => {
    const next = relocatedIndex(
      positioned,
      destinations,
      positioned.marker.type === "bookmarkEnd",
    );
    if (next === positioned.index) return positioned;
    changed = true;
    return { ...positioned, index: next };
  });
  return changed ? nextMarkers : markers;
};

/** Preserve a captured child's gap before the next surviving sibling. */
const relocateCaptured = (
  preserved: TablePreservedMarkup,
  destinations: GapDestinations,
): TablePreservedMarkup => {
  if (preserved.children === undefined) return preserved;
  let changed = false;
  const children = preserved.children.map((child) => {
    const index = relocatedIndex(child, destinations, false);
    if (index === child.index) return child;
    changed = true;
    return { ...child, index };
  });
  return changed ? { ...preserved, children } : preserved;
};

export type MergedTableOwner = { rowIndex: number; cellIndex: number };

export const pendingMergedOwner = (
  table: Table,
): MergedTableOwner | undefined => {
  const change = table.formatting?.gridChange?.geometryChange;
  if (change?.type !== "mergeCells") return undefined;
  const ids = new Set(change.revisionIds);
  for (const [rowIndex, row] of table.rows.entries()) {
    const cellIndex = row.cells.findIndex((cell) =>
      cell.propertyChanges?.some(({ info }) => ids.has(info.id)),
    );
    if (cellIndex !== -1) return { rowIndex, cellIndex };
  }
  return undefined;
};

type RelocateTableMarkupOptions = {
  mergedOwner?: MergedTableOwner | undefined;
  freshCellRevisions?: ReadonlySet<number>;
};
/** Keep bookmarks and captured XML in their gaps beside the surviving members. */
export const relocatePositionedMarkers = (
  before: Table,
  after: Table,
  options?: RelocateTableMarkupOptions,
): Table => {
  // A merge keeps its owner's source position even when every original
  // paragraph was replaced, or its content includes paragraphs of later cells.
  const mergedOwner = options?.mergedOwner;
  const rowOrigins = originsOf({ before: before.rows, after: after.rows });
  if (mergedOwner !== undefined)
    rowOrigins[mergedOwner.rowIndex] = mergedOwner.rowIndex;
  const wrappedRows = inheritCarriers({
    siblings: after.rows,
    origins: rowOrigins,
  });
  let rowsChanged = wrappedRows !== after.rows;
  const rows = [...wrappedRows];
  for (const [index, row] of wrappedRows.entries()) {
    const origin = rowOrigins[index] ?? null;
    const source = origin === null ? undefined : before.rows[origin];
    if (
      source === undefined ||
      (row.bookmarks === undefined &&
        row.preserved === undefined &&
        !row.cells.some(
          (cell) =>
            cell.contentControls !== undefined ||
            cell.carrierStack !== undefined,
        ))
    )
      continue;
    const fresh = options?.freshCellRevisions;
    const cellOrigins = originsOf({
      before: source.cells,
      after: row.cells,
      ...(fresh === undefined
        ? {}
        : {
            fresh: ({ structuralChange }) =>
              structuralChange?.type === "tableCellInsertion" &&
              fresh.has(structuralChange.info.id),
          }),
    });
    if (mergedOwner?.rowIndex === index)
      cellOrigins[mergedOwner.cellIndex] = mergedOwner.cellIndex;
    const preservesContinuations =
      after.formatting?.gridChange?.geometryChange?.type === "mergeCells"
        ? new Set(
            row.cells.filter(
              (cell) => cell.structuralChange?.type === "tableCellInsertion",
            ),
          )
        : undefined;
    const cells = inheritCarriers({
      siblings: row.cells,
      origins: cellOrigins,
      preserveCarriers: preservesContinuations,
    });
    const destinations = { origins: cellOrigins, siblings: cells };
    // Exact row replacements can state their own gaps. Re-address inherited
    // captures against the cells they originally stood beside.
    const markers =
      row.bookmarks !== undefined &&
      structurallyEqual(row.bookmarks, source.bookmarks)
        ? relocated(row.bookmarks, destinations)
        : row.bookmarks;
    const preserved =
      row.preserved !== undefined &&
      structurallyEqual(row.preserved, source.preserved)
        ? relocateCaptured(row.preserved, destinations)
        : row.preserved;
    if (
      preserved === row.preserved &&
      markers === row.bookmarks &&
      cells === row.cells
    )
      continue;
    rowsChanged = true;
    rows[index] = {
      ...row,
      cells,
      ...(markers === undefined ? {} : { bookmarks: markers }),
      ...(preserved === undefined ? {} : { preserved }),
    };
  }
  const destinations = { origins: rowOrigins, siblings: rows };
  const tableMarkers =
    after.bookmarks === undefined
      ? undefined
      : relocated(after.bookmarks, destinations);
  const preserved =
    after.preserved === undefined
      ? undefined
      : relocateCaptured(after.preserved, destinations);
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
      row.bookmarks !== undefined ||
      row.preserved?.children !== undefined ||
      row.contentControls !== undefined ||
      row.carrierStack !== undefined ||
      row.cells.some(
        (cell) =>
          cell.contentControls !== undefined || cell.carrierStack !== undefined,
      ),
  );
