import type { DocxPackage } from "../model/document";
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
import { structurallyEqual } from "./equality";
import { changedListBounds } from "./listSplice";
import { recentValues } from "./recentValues";
import { readDrawingIds } from "./tablePictureMove";

const isObjectValue = (value: unknown): value is object =>
  typeof value === "object" && value !== null;

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
  for (const [heldBy, field] of entries) {
    if (isObjectValue(field))
      walk({ value: field, visit, context: { heldBy, ownerType, table } });
  }
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

/**
 * The paragraph ids each package record holds, read once per record object.
 * Operations never change a record they are given, so a record an edit left
 * as it was holds the ids it held, and a batch reads again only the records
 * its edits replaced, not every paragraph of the package per operation.
 */
const RECORD_PARAGRAPH_IDS = new WeakMap<object, readonly string[]>();

const collectPackageParagraphIds = (value: unknown, out: string[]): void => {
  if (typeof value !== "object" || value === null) return;
  if (Array.isArray(value)) {
    for (const item of value) collectPackageParagraphIds(item, out);
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) collectPackageParagraphIds(item, out);
    return;
  }
  if (!isPlainRecord(value)) return;
  let ids = RECORD_PARAGRAPH_IDS.get(value);
  if (ids === undefined) {
    const own: string[] = [];
    const entries = Object.entries(value);
    const id = paragraphIdOf(entries);
    if (id !== undefined) own.push(id);
    for (const [, field] of entries) collectPackageParagraphIds(field, own);
    ids = own;
    RECORD_PARAGRAPH_IDS.set(value, ids);
  }
  for (const id of ids) out.push(id);
};

/** Every paragraph id of a package, each story walked once. */
export const packageParagraphIds = (pkg: DocxPackage): string[] => {
  const out: string[] = [];
  collectPackageParagraphIds(storiesOf(pkg), out);
  return out;
};

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

/** Native move range pairs use their own ids and share one move name. */
export const packageMoveRangeCensus = (
  pkg: DocxPackage,
): { ids: ReadonlySet<number>; names: ReadonlySet<string> } => {
  const ids = new Set<number>();
  const names = new Set<string>();
  walk({
    value: storiesOf(pkg),
    visit: (entries) => {
      const type = fieldOf(entries, "type");
      if (
        type !== "moveFromRangeStart" &&
        type !== "moveFromRangeEnd" &&
        type !== "moveToRangeStart" &&
        type !== "moveToRangeEnd"
      )
        return;
      const id = fieldOf(entries, "id");
      const name = fieldOf(entries, "name");
      if (typeof id === "number") ids.add(id);
      if (typeof name === "string") names.add(name);
    },
  });
  return { ids, names };
};

/** Drawing IDs in editable objects and in retained drawing markup. */
export const packageDrawingIds = (pkg: DocxPackage): Set<number> => {
  const ids = new Set<number>();
  walk({
    value: storiesOf(pkg),
    visit: (entries) => {
      const type = fieldOf(entries, "type");
      if (type === "image" || type === "shape") {
        const value = fieldOf(entries, "id");
        if (typeof value === "string" && /^\d+$/u.test(value))
          ids.add(Number(value));
      }
      for (const [key, value] of entries)
        if (
          typeof value === "string" &&
          (key === "rawXml" || key === "xml") &&
          value.includes("docPr")
        )
          readDrawingIds(value, ids);
    },
  });
  return ids;
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

/** A paragraph id's spelling for its number. */
export const formatParaId = (id: number): string =>
  id.toString(16).padStart(8, "0").toUpperCase();

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
 * The ids each record and list contributes, by the context it was walked in.
 *
 * A record's contribution depends only on its own fields, the field holding
 * it and its owner's type, so it is the same wherever an unchanged record is
 * shared: an edit replaces the records on the path to what it changed and
 * keeps the rest, and a census of the new version walks only those. A record
 * a census counts against others (a control stacked in a table's order is
 * counted once per table) is not shareable: it, and every record holding it,
 * are walked anew.
 */
const NO_KEYS: readonly string[] = Object.freeze([]);

type Census = { keys: readonly string[]; shareable: boolean };
const EMPTY_CENSUS: Census = Object.freeze({ keys: NO_KEYS, shareable: true });
type CensusContext = VisitContext & {
  stacked: Map<object | undefined, Map<number, Record<string, unknown>[]>>;
};
type OwnKeys = (
  entries: readonly [string, unknown][],
  context: CensusContext,
) => Census;

const contextKey = ({ heldBy, ownerType }: VisitContext): string =>
  `${heldBy ?? ""}\u0000${ownerType ?? ""}`;

const joined = (parts: readonly (readonly string[])[]): readonly string[] => {
  if (parts.length === 0) return NO_KEYS;
  if (parts.length === 1) return parts[0] ?? NO_KEYS;
  return parts.flat();
};

type CensusOf = (value: unknown, context: CensusContext) => Census;

const ROOT_CONTEXT: VisitContext = {
  heldBy: undefined,
  ownerType: undefined,
};

const censusWalker = (
  ownKeys: OwnKeys,
  sharing: "shared" | "full",
): CensusOf => {
  // By context, then by record: the contexts are few, the records many, so a
  // record costs one weak entry, not a map of its own.
  const known = new Map<string, WeakMap<object, readonly string[]>>();
  const remembered = (
    value: object,
    key: string | undefined,
    census: Census,
  ): Census => {
    if (key === undefined || !census.shareable) return census;
    // A frozen list is one a record below already holds (or none): share it.
    const keys = Object.isFrozen(census.keys)
      ? census.keys
      : Object.freeze([...census.keys]);
    let byRecord = known.get(key);
    if (byRecord === undefined) {
      byRecord = new WeakMap();
      known.set(key, byRecord);
    }
    byRecord.set(value, keys);
    return { keys, shareable: true };
  };
  const censusOf: CensusOf = (value, context) => {
    if (!isObjectValue(value)) return EMPTY_CENSUS;
    const isList = Array.isArray(value) || value instanceof Map;
    if (!isList && !isPlainRecord(value)) return EMPTY_CENSUS;
    const key = sharing === "full" ? undefined : contextKey(context);
    const cached = key === undefined ? undefined : known.get(key)?.get(value);
    if (cached !== undefined) return { keys: cached, shareable: true };
    let shareable = true;
    const parts: (readonly string[])[] = [];
    if (isList) {
      const items = Array.isArray(value) ? value : value.values();
      for (const item of items) {
        if (!isObjectValue(item)) continue;
        const census = censusOf(item, context);
        if (census.keys.length > 0) parts.push(census.keys);
        if (!census.shareable) shareable = false;
      }
    } else {
      const entries = Object.entries(value);
      const own = ownKeys(entries, context);
      shareable = own.shareable;
      if (own.keys.length > 0) parts.push(own.keys);
      let inner: ReturnType<typeof innerContext> | undefined;
      for (const [heldBy, child] of entries) {
        if (!isObjectValue(child)) continue;
        inner ??= innerContext(fieldOf(entries, "type"), value, context);
        const census = censusOf(child, { ...inner, heldBy });
        if (census.keys.length > 0) parts.push(census.keys);
        if (!census.shareable) shareable = false;
      }
    }
    return remembered(value, key, { keys: joined(parts), shareable });
  };
  return censusOf;
};

/** The context a record's fields are walked in, but for the field's own name. */
const innerContext = (
  type: unknown,
  record: object,
  context: CensusContext,
): Omit<CensusContext, "heldBy"> => {
  return {
    ownerType: typeof type === "string" ? type : undefined,
    table: type === "table" ? record : context.table,
    stacked: context.stacked,
  };
};

const freshContext = (): CensusContext => ({
  ...ROOT_CONTEXT,
  stacked: new Map(),
});

const subtreeCensus = (
  ownKeys: OwnKeys,
  sharing: "shared" | "full" = "shared",
): ((value: unknown) => string[]) => {
  const censusOf = censusWalker(ownKeys, sharing);
  return (value) => [...censusOf(value, freshContext()).keys];
};

const ownIdentityKeys: OwnKeys = (entries, context) => {
  const { heldBy, ownerType, table } = context;
  const keys: string[] = [];
  let shareable = true;
  const revision = revisionIdOf(entries, heldBy);
  if (revision !== undefined)
    keys.push(slotKey({ space: IDENTITY_SPACES.REVISION, id: revision }));
  const control = controlIdOf(entries, { heldBy, ownerType });
  if (control === undefined) return { keys, shareable };
  // A row- or cell-level content control holding several rows or cells is
  // recorded on each of them (`TableRow.contentControls`), and again in each
  // one's wrapper order (`carrierStack`): equal records there are one
  // control. A different record with the same id is a second one.
  // Only within one table: equal records in two tables are two controls.
  if (
    heldBy === CONTENT_CONTROL_STACK ||
    (heldBy === "properties" && ownerType === "sdt")
  ) {
    shareable = false;
    const record = Object.fromEntries(entries);
    const inTable =
      context.stacked.get(table) ??
      new Map<number, Record<string, unknown>[]>();
    context.stacked.set(table, inTable);
    const seen = inTable.get(control) ?? [];
    if (seen.some((other) => structurallyEqual(other, record)))
      return { keys, shareable };
    seen.push(record);
    inTable.set(control, seen);
  }
  keys.push(slotKey({ space: IDENTITY_SPACES.CONTROL, id: control }));
  return { keys, shareable };
};

/**
 * Every revision and content-control id in a value, as slot keys: tracked
 * changes and property changes by their `info`, content controls by their
 * properties. Every record is read anew: the value may belong to a caller
 * who changes it between calls.
 */
export const identityKeysIn: (value: unknown) => string[] = subtreeCensus(
  ownIdentityKeys,
  "full",
);

/**
 * {@link identityKeysIn} reusing what it counted in records an earlier
 * census met. Only for packages, which no operation changes (see
 * {@link censusOnce}).
 */
const sharedIdentityCensus = censusWalker(ownIdentityKeys, "shared");

/** The field holding the stack of content controls a table row or cell sits inside. */
const CONTENT_CONTROL_STACK = "contentControls";

const ownReservedKeys: OwnKeys = (entries, { heldBy }) => {
  const keys: string[] = [];
  if (heldBy === "nestedMoves") {
    const properties = fieldOf(entries, "properties");
    if (Array.isArray(properties)) {
      for (const pair of properties) {
        if (!Array.isArray(pair) || pair.length !== 2) continue;
        for (const id of pair)
          if (typeof id === "number")
            keys.push(slotKey({ space: IDENTITY_SPACES.REVISION, id }));
      }
    }
    return { keys, shareable: true };
  }
  if (heldBy === "keptCut") {
    const holders = fieldOf(entries, "holders");
    if (typeof holders === "object" && holders !== null) {
      for (const side of ["before", "after"] as const) {
        const id = Reflect.get(holders, side);
        if (typeof id === "number")
          keys.push(slotKey({ space: IDENTITY_SPACES.REVISION, id }));
      }
    }
    return { keys, shareable: true };
  }
  if (heldBy === "deferredRemove") {
    const blockers = fieldOf(entries, "blockers");
    if (Array.isArray(blockers)) {
      for (const id of blockers) {
        if (typeof id === "number")
          keys.push(slotKey({ space: IDENTITY_SPACES.REVISION, id }));
      }
    }
    return { keys, shareable: true };
  }
  if (heldBy !== "retainedAfter") return { keys, shareable: true };
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
        keys.push(slotKey({ space: slot.space, id: slot.id }));
      }
    }
  }
  return { keys, shareable: true };
};

/**
 * Pending cut references reserve IDs without becoming physical model
 * identities. Every record is read anew, as {@link identityKeysIn} does.
 */
export const reservedIdentityKeysIn: (value: unknown) => string[] =
  subtreeCensus(ownReservedKeys, "full");

/** {@link reservedIdentityKeysIn} reusing earlier counts; only for packages. */
const sharedReservedCensus = censusWalker(ownReservedKeys, "shared");

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
  censusOnce((pkg) => [
    ...sharedIdentityCensus(storiesOf(pkg), freshContext()).keys,
  ]);

/** The ids a package's pending cut references reserve. */
export const packageReservedIdentityKeys: (
  pkg: DocxPackage,
) => readonly string[] = censusOnce((pkg) => [
  ...sharedReservedCensus(pkg, freshContext()).keys,
]);

/** Which keys are taken, asked one key at a time. */
export type KeyMembership = { has: (key: string) => boolean };

/** The keys of one numeric identity space, in allocation order. */
export type IdentityKeySequence = {
  keyOf: (id: number) => string;
  numberOf: (key: string) => number | undefined;
};

export const paragraphIdSequence: IdentityKeySequence = {
  keyOf: formatParaId,
  numberOf: (key) => (isParaId(key) ? Number.parseInt(key, 16) : undefined),
};

/**
 * How many records of one package carry each key, answered without a census
 * of the whole package.
 */
export type IdentityKeyIndex = {
  /** Keys present in the other package but absent from this one. */
  introducedIn: (next: DocxPackage) => string[];
  /** How many records carry `key`. */
  count: (key: string) => number;
  /** Whether any record carries `key`. */
  has: (key: string) => boolean;
  /** The first absent positive id of a stable sequence, capped at `limit`. */
  firstMissing: (sequence: IdentityKeySequence, limit: number) => number;
  /** Every key with its count, as {@link countKeys} gives them. */
  counts: () => ReadonlyMap<string, number>;
  /** How many keys more than one record carries. */
  repeated: () => number;
};

/**
 * One package's counts: either every count (a root) or the change from the
 * package it was derived from. Lookups add the changes up the chain, which
 * {@link MAX_CHAIN} bounds.
 */
type IndexNode = {
  parent: IndexNode | undefined;
  changes: Map<string, number>;
  depth: number;
  /** Whether every record of the package can be counted apart from the rest. */
  shareable: boolean;
  /** How many keys more than one record of the package carries. */
  repeated: number;
  /** All ids strictly below each cached candidate are present. */
  prefixes: Map<IdentityKeySequence, number> | undefined;
};

const MAX_CHAIN = 32;

const countIn = (node: IndexNode, key: string): number => {
  let count = 0;
  for (let at: IndexNode | undefined = node; at !== undefined; at = at.parent)
    count += at.changes.get(key) ?? 0;
  return count;
};

/**
 * Turns a node into a root holding every count. Its own total is unchanged,
 * so nodes derived from it still add up to theirs.
 */
const settle = (node: IndexNode): void => {
  if (node.parent === undefined) return;
  const chain: IndexNode[] = [];
  for (let at: IndexNode | undefined = node; at !== undefined; at = at.parent)
    chain.push(at);
  chain.reverse();
  const counts = new Map(chain[0]?.changes);
  for (const step of chain.slice(1)) {
    for (const [key, change] of step.changes) {
      const count = (counts.get(key) ?? 0) + change;
      if (count === 0) counts.delete(key);
      else counts.set(key, count);
    }
  }
  node.parent = undefined;
  node.changes = counts;
  node.depth = 0;
};

type Kind = "leaf" | "list" | "map" | "record";

const kindOf = (value: unknown): Kind => {
  if (typeof value !== "object" || value === null) return "leaf";
  if (Array.isArray(value)) return "list";
  if (value instanceof Map) return "map";
  return isPlainRecord(value) ? "record" : "leaf";
};

const tally = (
  changes: Map<string, number>,
  keys: readonly string[],
  sign: 1 | -1,
): void => {
  for (const key of keys) changes.set(key, (changes.get(key) ?? 0) + sign);
};

/**
 * Adds to `changes` what turning `before` into `after` does to a census.
 *
 * A census is the sum of what each record contributes in its context, so
 * the change is found by walking both values side by side: a value both
 * share counts the same in each, and a record or list both hold is compared
 * field by field and item by item. The rest is counted out of `before` and
 * into `after`. `before` must be shareable (see {@link censusWalker}); the
 * walk stops, answering `false`, when `after` is not.
 */
const censusChange = (
  censusOf: CensusOf,
  ownKeys: OwnKeys,
): ((
  before: unknown,
  after: unknown,
  changes: Map<string, number>,
) => boolean) => {
  const replaced = (
    before: unknown,
    after: unknown,
    context: CensusContext,
    changes: Map<string, number>,
  ): boolean => {
    tally(changes, censusOf(before, context).keys, -1);
    const added = censusOf(after, context);
    if (!added.shareable) return false;
    tally(changes, added.keys, 1);
    return true;
  };
  const items = (
    before: readonly unknown[],
    after: readonly unknown[],
    context: CensusContext,
    changes: Map<string, number>,
  ): boolean => {
    const {
      start,
      previousEnd: beforeEnd,
      end: afterEnd,
    } = changedListBounds(before, after);
    const paired = Math.min(beforeEnd, afterEnd) - start;
    for (let offset = 0; offset < paired; offset += 1) {
      if (
        !change(before[start + offset], after[start + offset], context, changes)
      )
        return false;
    }
    for (let index = start + paired; index < beforeEnd; index += 1)
      replaced(before[index], undefined, context, changes);
    for (let index = start + paired; index < afterEnd; index += 1) {
      if (!replaced(undefined, after[index], context, changes)) return false;
    }
    return true;
  };
  const fields = (
    before: object,
    after: object,
    context: CensusContext,
    changes: Map<string, number>,
  ): boolean => {
    const beforeEntries = Object.entries(before);
    const afterEntries = Object.entries(after);
    // A record's type is its fields' owner type: when it changes, every
    // field counts in another context.
    const type = fieldOf(afterEntries, "type");
    if (fieldOf(beforeEntries, "type") !== type)
      return replaced(before, after, context, changes);
    tally(changes, ownKeys(beforeEntries, context).keys, -1);
    const own = ownKeys(afterEntries, context);
    if (!own.shareable) return false;
    tally(changes, own.keys, 1);

    // Only object-valued fields can contribute keys below this record.
    let beforeFields: Map<string, object> | undefined;
    for (const [heldBy, value] of beforeEntries) {
      if (!isObjectValue(value)) continue;
      beforeFields ??= new Map();
      beforeFields.set(heldBy, value);
    }
    let inner: ReturnType<typeof innerContext> | undefined;
    for (const [heldBy, value] of afterEntries) {
      const was = beforeFields?.get(heldBy);
      if (was !== undefined) beforeFields?.delete(heldBy);
      if (was === value || (was === undefined && !isObjectValue(value)))
        continue;
      inner ??= innerContext(type, after, context);
      if (!change(was, value, { ...inner, heldBy }, changes)) return false;
    }
    for (const [heldBy, was] of beforeFields ?? []) {
      inner ??= innerContext(type, after, context);
      replaced(was, undefined, { ...inner, heldBy }, changes);
    }
    return true;
  };
  const change = (
    before: unknown,
    after: unknown,
    context: CensusContext,
    changes: Map<string, number>,
  ): boolean => {
    if (before === after) return true;
    const kind = kindOf(before);
    if (kind !== kindOf(after))
      return replaced(before, after, context, changes);
    if (kind === "leaf") return true;
    if (Array.isArray(before) && Array.isArray(after))
      return items(before, after, context, changes);
    if (before instanceof Map && after instanceof Map) {
      for (const [key, value] of after) {
        if (!change(before.get(key), value, context, changes)) return false;
      }
      for (const [key, value] of before) {
        if (!after.has(key)) replaced(value, undefined, context, changes);
      }
      return true;
    }
    if (typeof before !== "object" || before === null) return true;
    if (typeof after !== "object" || after === null) return true;
    return fields(before, after, context, changes);
  };
  return (before, after, changes) =>
    change(before, after, freshContext(), changes);
};

/**
 * A key index per package, each derived from the package indexed last.
 *
 * Operations never change a package: each returns a new one sharing what it
 * left as it was, so the package before an edit and the one after differ
 * only along the paths to what changed, and an index of one is found from
 * the other's by walking those paths. A package holding a record counted
 * against others (see {@link censusWalker}) is counted in full.
 */
const keyIndex = (
  censusOf: CensusOf,
  ownKeys: OwnKeys,
  viewOf: (pkg: DocxPackage) => unknown,
): ((pkg: DocxPackage) => IdentityKeyIndex) => {
  const nodes = new WeakMap<DocxPackage, IndexNode>();
  const changeBetween = censusChange(censusOf, ownKeys);
  const recent = recentValues<DocxPackage>();
  const baseFor = ():
    { previous: DocxPackage; base: IndexNode } | undefined => {
    for (const previous of recent.values()) {
      const base = nodes.get(previous);
      if (base?.shareable === true) return { previous, base };
    }
    return undefined;
  };
  const derived = (pkg: DocxPackage): IndexNode | undefined => {
    const found = baseFor();
    if (found === undefined) return undefined;
    const { previous, base } = found;
    const changes = new Map<string, number>();
    if (!changeBetween(viewOf(previous), viewOf(pkg), changes))
      return undefined;
    // Only a key whose count changed can start or stop repeating.
    let { repeated } = base;
    const prefixes =
      base.prefixes === undefined ? undefined : new Map(base.prefixes);
    for (const [key, change] of changes) {
      if (change === 0) {
        changes.delete(key);
        continue;
      }
      const was = countIn(base, key);
      const count = was + change;
      repeated += Number(count > 1) - Number(was > 1);
      if (count === 0 && prefixes !== undefined) {
        for (const [sequence, candidate] of prefixes) {
          const id = sequence.numberOf(key);
          if (id !== undefined && id > 0 && id < candidate)
            prefixes.set(sequence, id);
        }
      }
    }
    const node: IndexNode = {
      parent: base,
      changes,
      depth: base.depth + 1,
      shareable: true,
      repeated,
      prefixes,
    };
    if (node.depth >= MAX_CHAIN) settle(node);
    return node;
  };
  const counted = (pkg: DocxPackage): IndexNode => {
    const census = censusOf(viewOf(pkg), freshContext());
    const changes = countKeys(census.keys);
    let repeated = 0;
    for (const count of changes.values()) if (count > 1) repeated += 1;
    return {
      parent: undefined,
      changes,
      depth: 0,
      shareable: census.shareable,
      repeated,
      prefixes: undefined,
    };
  };
  const index = (pkg: DocxPackage): IdentityKeyIndex => {
    let node = nodes.get(pkg);
    if (node === undefined) {
      node = derived(pkg) ?? counted(pkg);
      nodes.set(pkg, node);
    }
    recent.remember(pkg);
    const indexed = node;
    return {
      introducedIn: (next) => {
        if (pkg === next) return [];
        const cached = nodes.get(next);
        const changes = cached?.parent === indexed ? cached.changes : new Map();
        if (
          cached?.parent !== indexed &&
          (!indexed.shareable ||
            !changeBetween(viewOf(pkg), viewOf(next), changes))
        ) {
          return [...index(next).counts().keys()].filter(
            (key) => countIn(indexed, key) === 0,
          );
        }
        const introduced: string[] = [];
        for (const [key, change] of changes) {
          if (change > 0 && countIn(indexed, key) === 0) introduced.push(key);
        }
        return introduced;
      },
      count: (key) => countIn(indexed, key),
      has: (key) => countIn(indexed, key) > 0,
      firstMissing: (sequence, limit) => {
        indexed.prefixes ??= new Map();
        const known = indexed.prefixes.get(sequence) ?? 1;
        let id = Math.min(known, limit);
        while (id < limit && countIn(indexed, sequence.keyOf(id)) > 0) id += 1;
        if (id > known) indexed.prefixes.set(sequence, id);
        return id;
      },
      counts: () => {
        settle(indexed);
        return indexed.changes;
      },
      repeated: () => indexed.repeated,
    };
  };
  return index;
};

/** How many records of a package carry each revision and content-control id. */
export const packageIdentityKeyIndex: (pkg: DocxPackage) => IdentityKeyIndex =
  keyIndex(sharedIdentityCensus, ownIdentityKeys, storiesOf);

/** How many of a package's pending cut references reserve each id. */
export const packageReservedKeyIndex: (pkg: DocxPackage) => IdentityKeyIndex =
  keyIndex(sharedReservedCensus, ownReservedKeys, (pkg) => pkg);

const ownParagraphKeys: OwnKeys = (entries) => {
  const id = paragraphIdOf(entries);
  return { keys: id === undefined ? NO_KEYS : [idKey(id)], shareable: true };
};

/** The numeric spelling lease accounting has always accepted. */
export const parseLeasedParaId = (id: string): number | undefined => {
  const value = Number.parseInt(id, 16);
  return Number.isSafeInteger(value) ? value : undefined;
};

const ownLeasedParagraphKeys: OwnKeys = (entries) => {
  const raw = paragraphIdOf(entries);
  const id = raw === undefined ? undefined : parseLeasedParaId(idKey(raw));
  return {
    keys: id === undefined ? NO_KEYS : [id.toString(16)],
    shareable: true,
  };
};

/** Numeric lease identities can have multiple source spellings. */
export const packageLeasedParagraphKeyIndex: (
  pkg: DocxPackage,
) => IdentityKeyIndex = keyIndex(
  censusWalker(ownLeasedParagraphKeys, "shared"),
  ownLeasedParagraphKeys,
  storiesOf,
);

/**
 * How many paragraphs of a package carry each id, by {@link idKey}: what
 * {@link countIds} of {@link packageParagraphIds} counts, each package found
 * from one indexed before it.
 */
export const packageParagraphIdIndex: (pkg: DocxPackage) => IdentityKeyIndex =
  keyIndex(
    censusWalker(ownParagraphKeys, "shared"),
    ownParagraphKeys,
    storiesOf,
  );
