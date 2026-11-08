/**
 * The id domains of a package.
 *
 * `w14:paraId` is unique across every part of a package, not per story: body,
 * headers, footers, notes, comments and the text boxes inside any of them
 * share one space. Revision ids (`w:id` on tracked changes and property
 * changes) share another, and content-control ids a third.
 *
 * A census walks each story's own content exactly once. `DocumentBody.sections`
 * is a derived view of the body's blocks, not a story, so it is left out: the
 * same paragraph must not count twice. The walk is otherwise structural, so a
 * story the model gains is covered without a list to keep in step.
 */

import type { DocxPackage } from "../model/document";
import { structurallyEqual } from "./equality";

const isPlainRecord = (value: object): boolean => {
  const prototype: unknown = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
};

/** A record's fields, and the key of the field (or of the list) that holds it. */
type VisitContext = {
  heldBy: string | undefined;
  ownerType: string | undefined;
  /** The innermost table the record sits in. */
  table?: object | undefined;
};
type Visit = (
  entries: readonly [string, unknown][],
  context: VisitContext,
) => void;
type WalkOptions = { value: unknown; visit: Visit; context?: VisitContext };

const walk = ({
  value,
  visit,
  context = { heldBy: undefined, ownerType: undefined },
}: WalkOptions): void => {
  if (typeof value !== "object" || value === null) return;
  if (Array.isArray(value)) {
    for (const item of value) walk({ value: item, visit, context });
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) walk({ value: item, visit, context });
    return;
  }
  if (!isPlainRecord(value)) return;
  const entries = Object.entries(value);
  visit(entries, context);
  const type = fieldOf(entries, "type");
  const ownerType = typeof type === "string" ? type : undefined;
  const table = type === "table" ? value : context.table;
  for (const [heldBy, field] of entries)
    walk({ value: field, visit, context: { heldBy, ownerType, table } });
};

/** The package's stories, each once: the body without its derived section view. */
const storiesOf = (pkg: DocxPackage): unknown => ({
  ...pkg,
  document: { ...pkg.document, sections: undefined },
});

const fieldOf = (
  entries: readonly [string, unknown][],
  name: string,
): unknown => entries.find(([key]) => key === name)?.[1];

const paragraphIdOf = (
  entries: readonly [string, unknown][],
): string | undefined => {
  if (fieldOf(entries, "type") !== "paragraph") return undefined;
  const id = fieldOf(entries, "paraId");
  return typeof id === "string" ? id : undefined;
};

/** Every `paraId` of every paragraph in a value, nested ones included, in document order. */
export const paragraphIdsIn = (value: unknown): string[] => {
  const out: string[] = [];
  walk({
    value,
    visit: (entries) => {
      const id = paragraphIdOf(entries);
      if (id !== undefined) out.push(id);
    },
  });
  return out;
};

/** Every paragraph id of a package, each story walked once. */
export const packageParagraphIds = (pkg: DocxPackage): string[] =>
  paragraphIdsIn(storiesOf(pkg));

export type PackageBookmark = { id: number; name: string };
export type PackageBookmarkCensus = {
  ids: ReadonlySet<number>;
  names: ReadonlySet<string>;
  pairs: readonly PackageBookmark[];
};

/** Bookmark ids, names and uniquely paired starts from every package story. */
export const packageBookmarkCensus = (
  pkg: DocxPackage,
): PackageBookmarkCensus => {
  const ids = new Set<number>();
  const names = new Set<string>();
  const starts = new Map<number, { name: string; count: number }>();
  const ends = new Map<number, number>();
  walk({
    value: storiesOf(pkg),
    visit: (entries) => {
      const type = fieldOf(entries, "type");
      const id = fieldOf(entries, "id");
      if (typeof id !== "number") return;
      if (type === "bookmarkStart") {
        ids.add(id);
        const name = fieldOf(entries, "name");
        if (typeof name !== "string") return;
        names.add(name);
        const prior = starts.get(id);
        starts.set(id, { name, count: (prior?.count ?? 0) + 1 });
      } else if (type === "bookmarkEnd") {
        ids.add(id);
        ends.set(id, (ends.get(id) ?? 0) + 1);
      }
    },
  });
  const pairs = [...starts].flatMap(([id, start]) =>
    start.count === 1 && ends.get(id) === 1 ? [{ id, name: start.name }] : [],
  );
  return { ids, names, pairs };
};

/**
 * The key two spellings of one id share. `ST_LongHexNumber` is hex, so
 * `0000abcd` and `0000ABCD` are the same id; every lookup and comparison of
 * paragraph ids goes through this.
 */
export const idKey = (id: string): string => id.toUpperCase();

/** How many records carry each key. */
export const countKeys = (keys: readonly string[]): Map<string, number> => {
  const counts = new Map<string, number>();
  for (const key of keys) counts.set(key, (counts.get(key) ?? 0) + 1);
  return counts;
};

/** How many paragraphs carry each id, by {@link idKey}. */
export const countIds = (ids: readonly string[]): Map<string, number> =>
  countKeys(ids.map(idKey));

/** Whether `incoming` repeats an id of `existing` or of itself, by {@link idKey}. */
export const collides = (
  existing: ReadonlyMap<string, number>,
  incoming: readonly string[],
): boolean => {
  const seen = new Set<string>();
  return incoming.some((id) => {
    const key = idKey(id);
    const repeated = seen.has(key) || (existing.get(key) ?? 0) > 0;
    seen.add(key);
    return repeated;
  });
};

/** `ST_LongHexNumber` as `w14:paraId` uses it: eight hex digits below `0x80000000`. */
const PARA_ID_PATTERN = /^[0-7][0-9A-Fa-f]{7}$/u;
/** Zero is reserved: it means "no id". */
const RESERVED_PARA_ID_PATTERN = /^0{8}$/u;

/** Whether an id can name a paragraph an operation creates. */
export const isParaId = (id: string): boolean =>
  PARA_ID_PATTERN.test(id) && !RESERVED_PARA_ID_PATTERN.test(id);

/** The id spaces of revisions and content controls. */
export const IDENTITY_SPACES: Readonly<{
  readonly REVISION: "revision";
  readonly CONTROL: "control";
}> = Object.freeze({
  REVISION: "revision",
  CONTROL: "control",
} as const);

export type IdentitySpace =
  (typeof IDENTITY_SPACES)[keyof typeof IDENTITY_SPACES];

/** One id a record carries, in its space. */
export type IdentitySlot = { space: IdentitySpace; id: number };

export const slotKey = ({ space, id }: IdentitySlot): string =>
  `${space}:${id}`;

const revisionIdOf = (
  entries: readonly [string, unknown][],
  heldBy: string | undefined,
): number | undefined => {
  // Grid changes carry a physical revision id without author metadata.
  if (
    heldBy === "gridChange" &&
    Array.isArray(fieldOf(entries, "columnWidths"))
  ) {
    const id = fieldOf(entries, "id");
    return typeof id === "number" ? id : undefined;
  }
  const info = fieldOf(entries, "info");
  if (typeof info !== "object" || info === null) return undefined;
  const infoEntries = Object.entries(info);
  const id = fieldOf(infoEntries, "id");
  return typeof id === "number" &&
    typeof fieldOf(infoEntries, "author") === "string"
    ? id
    : undefined;
};

const controlIdOf = (
  entries: readonly [string, unknown][],
  { heldBy, ownerType }: VisitContext,
): number | undefined => {
  const id = fieldOf(entries, "id");
  if (typeof id !== "number") return undefined;
  const ownedProperties =
    heldBy === "properties" &&
    (ownerType === "inlineSdt" || ownerType === "blockSdt");
  return ownedProperties ||
    heldBy === CONTENT_CONTROL_STACK ||
    typeof fieldOf(entries, "sdtType") === "string"
    ? id
    : undefined;
};

/**
 * Every revision and content-control id in a value, as slot keys: tracked
 * changes and property changes by their `info`, content controls by their
 * properties.
 */
export const identityKeysIn = (value: unknown): string[] => {
  const out: string[] = [];
  // A row- or cell-level content control holding several rows or cells is
  // recorded on each of them (`TableRow.contentControls`), and again in each
  // one's wrapper order (`carrierStack`): equal records there are one
  // control. A different record with the same id is a second one.
  // Only within one table: equal records in two tables are two controls.
  const stacked = new Map<
    object | undefined,
    Map<number, Record<string, unknown>[]>
  >();
  walk({
    value,
    visit: (entries, { heldBy, ownerType, table }) => {
      const revision = revisionIdOf(entries, heldBy);
      if (revision !== undefined) {
        out.push(slotKey({ space: IDENTITY_SPACES.REVISION, id: revision }));
      }
      const control = controlIdOf(entries, { heldBy, ownerType });
      if (control === undefined) return;
      // A table wrapper's mixed-order record (`carrierStack`) restates the
      // same control the stack holds.
      if (
        heldBy === CONTENT_CONTROL_STACK ||
        (heldBy === "properties" && ownerType === "sdt")
      ) {
        const record = Object.fromEntries(entries);
        const inTable =
          stacked.get(table) ?? new Map<number, Record<string, unknown>[]>();
        stacked.set(table, inTable);
        const seen = inTable.get(control) ?? [];
        if (seen.some((other) => structurallyEqual(other, record))) return;
        seen.push(record);
        inTable.set(control, seen);
      }
      out.push(slotKey({ space: IDENTITY_SPACES.CONTROL, id: control }));
    },
  });
  return out;
};

/** The field holding the stack of content controls a table row or cell sits inside. */
const CONTENT_CONTROL_STACK = "contentControls";

/** Pending cut references reserve IDs without becoming physical model identities. */
export const reservedIdentityKeysIn = (value: unknown): string[] => {
  const out: string[] = [];
  walk({
    value,
    visit: (entries, { heldBy }) => {
      if (heldBy === "deferredRemove") {
        const blockers = fieldOf(entries, "blockers");
        if (Array.isArray(blockers)) {
          for (const id of blockers) {
            if (typeof id === "number")
              out.push(slotKey({ space: IDENTITY_SPACES.REVISION, id }));
          }
        }
        return;
      }
      if (heldBy !== "retainedAfter") return;
      for (const field of ["source", "target"]) {
        const slots = fieldOf(entries, field);
        if (!Array.isArray(slots)) continue;
        for (const slot of slots) {
          if (
            typeof slot !== "object" ||
            slot === null ||
            !("space" in slot) ||
            !("id" in slot)
          )
            continue;
          if (
            (slot.space === IDENTITY_SPACES.REVISION ||
              slot.space === IDENTITY_SPACES.CONTROL) &&
            typeof slot.id === "number"
          ) {
            out.push(slotKey({ space: slot.space, id: slot.id }));
          }
        }
      }
    },
  });
  return out;
};

/** Every revision and content-control id of a package, each story walked once. */
/**
 * A census of a package, once per package object. Operations never change a
 * package they are given: each returns a new one, sharing what it left as it
 * was, so a package's ids are read once however often a plan asks.
 */
const censusOnce = (
  census: (pkg: DocxPackage) => string[],
): ((pkg: DocxPackage) => readonly string[]) => {
  const known = new WeakMap<DocxPackage, readonly string[]>();
  return (pkg) => {
    const cached = known.get(pkg);
    if (cached !== undefined) return cached;
    const keys = Object.freeze(census(pkg));
    known.set(pkg, keys);
    return keys;
  };
};

/** Every revision and content-control id of a package, each story walked once. */
export const packageIdentityKeys: (pkg: DocxPackage) => readonly string[] =
  censusOnce((pkg) => identityKeysIn(storiesOf(pkg)));

/** The ids a package's pending cut references reserve. */
export const packageReservedIdentityKeys: (
  pkg: DocxPackage,
) => readonly string[] = censusOnce((pkg) => reservedIdentityKeysIn(pkg));
