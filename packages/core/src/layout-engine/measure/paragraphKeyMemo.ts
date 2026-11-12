/**
 * The measurement key of a paragraph that an edit only moved, without
 * serializing it again.
 *
 * Every layout pass converts the document into fresh blocks, and a paragraph
 * after an edit comes out as a copy of the block kept for it with shifted
 * positions. Its key is the key of what it was moved from whenever every
 * value the key reads is the one it read then: the key never reads
 * positions, and the values a moved block shares are never written after
 * conversion. A block whose values differ, or whose origin is unknown, is
 * serialized in full.
 */

import { documentFontFactsCacheIdentity } from "../../fonts/documentFontFacts";
import type { DocumentFontFacts } from "../../fonts/documentFontFacts";
import type { ParagraphBlock } from "../types";

/** The run fields the key leaves out: they say where a run stands. */
const POSITION_FIELDS: ReadonlySet<string> = new Set(["pmStart", "pmEnd"]);
const NO_FIELDS: ReadonlySet<string> = new Set();

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" &&
  value !== null &&
  Object.getPrototypeOf(value) === Object.prototype;

/** A plain object field as it was, one level down. */
type NestedValue = { readonly fields: readonly (readonly [string, unknown])[] };

type FieldValue = { readonly value: unknown } | NestedValue;

type Fields = readonly (readonly [string, FieldValue])[];

const fieldsOf = (
  record: Record<string, unknown>,
  skip: ReadonlySet<string>,
): Fields =>
  Object.entries(record)
    .filter(([name]) => !skip.has(name))
    .map(([name, value]) => [
      name,
      isPlainObject(value) ? { fields: Object.entries(value) } : { value },
    ]);

const countFields = (
  record: Record<string, unknown>,
  skip: ReadonlySet<string>,
): number => {
  let count = 0;
  for (const name of Object.keys(record)) if (!skip.has(name)) count += 1;
  return count;
};

/** Whether `record` holds the values `fields` recorded, nested ones included. */
const holds = (
  record: Record<string, unknown>,
  fields: Fields,
  skip: ReadonlySet<string>,
): boolean => {
  if (countFields(record, skip) !== fields.length) return false;
  for (const [name, recorded] of fields) {
    const value = record[name];
    if (!("fields" in recorded)) {
      if (value !== recorded.value) return false;
      continue;
    }
    if (!isPlainObject(value)) return false;
    if (Object.keys(value).length !== recorded.fields.length) return false;
    for (const [nestedName, nestedValue] of recorded.fields)
      if (value[nestedName] !== nestedValue) return false;
  }
  return true;
};

type Recorded = {
  key: string;
  generation: number;
  attrs: Fields | undefined;
  runs: readonly Fields[];
  /** Each font facts value's identity, which may change on the same object. */
  factIdentities: readonly (string | undefined)[];
};

const factIdentitiesOf = (block: ParagraphBlock): (string | undefined)[] => {
  const identities = [factIdentity(block.attrs?.documentFontFacts)];
  for (const run of block.runs)
    if ("documentFontFacts" in run)
      identities.push(factIdentity(run.documentFontFacts));
  return identities;
};

const factIdentity = (facts: DocumentFontFacts | undefined) =>
  documentFontFactsCacheIdentity(facts);

const record = (
  block: ParagraphBlock,
  key: string,
  generation: number,
): Recorded => ({
  key,
  generation,
  attrs:
    block.attrs === undefined
      ? undefined
      : fieldsOf(block.attrs as Record<string, unknown>, NO_FIELDS),
  runs: block.runs.map((run) =>
    fieldsOf(run as unknown as Record<string, unknown>, POSITION_FIELDS),
  ),
  factIdentities: factIdentitiesOf(block),
});

const stillHolds = (
  recorded: Recorded,
  block: ParagraphBlock,
  generation: number,
): boolean => {
  if (recorded.generation !== generation) return false;
  if (recorded.runs.length !== block.runs.length) return false;
  if ((recorded.attrs === undefined) !== (block.attrs === undefined))
    return false;
  if (
    recorded.attrs !== undefined &&
    !holds(block.attrs as Record<string, unknown>, recorded.attrs, NO_FIELDS)
  )
    return false;
  for (const [index, run] of block.runs.entries()) {
    const fields = recorded.runs[index];
    if (
      !fields ||
      !holds(run as unknown as Record<string, unknown>, fields, POSITION_FIELDS)
    )
      return false;
  }
  const identities = factIdentitiesOf(block);
  return (
    identities.length === recorded.factIdentities.length &&
    identities.every(
      (identity, index) => identity === recorded.factIdentities[index],
    )
  );
};

/** The block each block was moved from, as the block cache keeps it. */
const origins = new WeakMap<ParagraphBlock, ParagraphBlock>();
/** The key last given to a block or to blocks moved from an origin. */
const keys = new WeakMap<ParagraphBlock, Recorded>();

/**
 * Record that `block` is `origin`, the block kept for its paragraph, moved:
 * the two show the same, positions aside.
 */
export const noteParagraphOrigin = (
  block: ParagraphBlock,
  origin: ParagraphBlock,
): void => {
  origins.set(block, origin);
};

/** `block`'s key: a recorded one while its values hold, else `compute`'s. */
export const memoizedParagraphKey = (
  block: ParagraphBlock,
  generation: number,
  compute: (block: ParagraphBlock) => string,
): string => {
  const own = keys.get(block);
  if (own && stillHolds(own, block, generation)) return own.key;
  const origin = origins.get(block);
  const moved = origin === undefined ? undefined : keys.get(origin);
  if (moved && stillHolds(moved, block, generation)) {
    keys.set(block, moved);
    return moved.key;
  }
  const key = compute(block);
  const recorded = record(block, key, generation);
  keys.set(block, recorded);
  if (origin !== undefined) keys.set(origin, recorded);
  return key;
};
