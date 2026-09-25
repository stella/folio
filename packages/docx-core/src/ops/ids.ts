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
type Visit = (entries: readonly [string, unknown][], heldBy: string | undefined) => void;

const walk = (value: unknown, visit: Visit, heldBy?: string): void => {
  if (typeof value !== "object" || value === null) {
    return;
  }
  if (Array.isArray(value)) {
    for (const item of value) walk(item, visit, heldBy);
    return;
  }
  if (value instanceof Map) {
    for (const item of value.values()) walk(item, visit, heldBy);
    return;
  }
  // Bytes, dates and other built-ins hold no records.
  if (!isPlainRecord(value)) {
    return;
  }
  const entries = Object.entries(value);
  visit(entries, heldBy);
  for (const [key, field] of entries) walk(field, visit, key);
};

/** The package's stories, each once: the body without its derived section view. */
const storiesOf = (pkg: DocxPackage): unknown => ({
  ...pkg,
  document: { ...pkg.document, sections: undefined },
});

const fieldOf = (entries: readonly [string, unknown][], name: string): unknown =>
  entries.find(([key]) => key === name)?.[1];

const paragraphIdOf = (entries: readonly [string, unknown][]): string | undefined => {
  if (fieldOf(entries, "type") !== "paragraph") return undefined;
  const id = fieldOf(entries, "paraId");
  return typeof id === "string" ? id : undefined;
};

/** Every `paraId` of every paragraph in a value, nested ones included, in document order. */
export const paragraphIdsIn = (value: unknown): string[] => {
  const out: string[] = [];
  walk(value, (entries) => {
    const id = paragraphIdOf(entries);
    if (id !== undefined) out.push(id);
  });
  return out;
};

/** Every paragraph id of a package, each story walked once. */
export const packageParagraphIds = (pkg: DocxPackage): string[] => paragraphIdsIn(storiesOf(pkg));

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
export const countIds = (ids: readonly string[]): Map<string, number> => countKeys(ids.map(idKey));

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
export const IDENTITY_SPACES = Object.freeze({ REVISION: "revision", CONTROL: "control" } as const);

export type IdentitySpace = (typeof IDENTITY_SPACES)[keyof typeof IDENTITY_SPACES];

/** One id a record carries, in its space. */
export type IdentitySlot = { space: IdentitySpace; id: number };

export const slotKey = ({ space, id }: IdentitySlot): string => `${space}:${id}`;

const revisionIdOf = (entries: readonly [string, unknown][]): number | undefined => {
  const info = fieldOf(entries, "info");
  if (typeof info !== "object" || info === null) return undefined;
  const infoEntries = Object.entries(info);
  const id = fieldOf(infoEntries, "id");
  return typeof id === "number" && typeof fieldOf(infoEntries, "author") === "string"
    ? id
    : undefined;
};

const controlIdOf = (entries: readonly [string, unknown][]): number | undefined => {
  const id = fieldOf(entries, "id");
  return typeof fieldOf(entries, "sdtType") === "string" && typeof id === "number" ? id : undefined;
};

/**
 * Every revision and content-control id in a value, as slot keys: tracked
 * changes and property changes by their `info`, content controls by their
 * properties.
 */
export const identityKeysIn = (value: unknown): string[] => {
  const out: string[] = [];
  // A row- or cell-level content control holding several rows or cells is
  // recorded on each of them (`TableRow.contentControls`): equal records there
  // are one control. A different record with the same id is a second one.
  const stacked = new Map<number, Record<string, unknown>[]>();
  walk(value, (entries, heldBy) => {
    const revision = revisionIdOf(entries);
    if (revision !== undefined) {
      out.push(slotKey({ space: IDENTITY_SPACES.REVISION, id: revision }));
    }
    const control = controlIdOf(entries);
    if (control === undefined) return;
    if (heldBy === CONTENT_CONTROL_STACK) {
      const record = Object.fromEntries(entries);
      const seen = stacked.get(control) ?? [];
      if (seen.some((other) => structurallyEqual(other, record))) return;
      seen.push(record);
      stacked.set(control, seen);
    }
    out.push(slotKey({ space: IDENTITY_SPACES.CONTROL, id: control }));
  });
  return out;
};

/** The field holding the stack of content controls a table row or cell sits inside. */
const CONTENT_CONTROL_STACK = "contentControls";

/** Every revision and content-control id of a package, each story walked once. */
export const packageIdentityKeys = (pkg: DocxPackage): string[] => identityKeysIn(storiesOf(pkg));
