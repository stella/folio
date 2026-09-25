/**
 * Model equality for the invariants that save a document and parse it back.
 *
 * Deep equality on two parsed packages is unusable raw: both save paths refresh
 * `dcterms:modified` and `lastModifiedBy`, `originalBuffer` is the input bytes
 * by definition, media are large binaries whose identity is their length, and a
 * `Map`'s iteration order is not content. The projection below erases exactly
 * those, and nothing else, so a surviving difference is a difference in the
 * document.
 *
 * It also erases the capture slots (`sourceXml`, `verbatimXml`, `rawXml` and
 * their fingerprints). Those are not content: they are the serializer's licence
 * to replay bytes it has already seen, and the invariant that makes the real
 * serializers run removes them on purpose. Comparing them would report the
 * removal rather than its consequences.
 *
 * The declared normalisations are the ones
 * `packages/core/src/docx/saveEquivalence.property.test.ts` defines over the
 * committed fixtures, extended by the capture slots. They are restated here
 * rather than imported because that projection is test-local; see
 * `corpus/README.md` for the standing proposal to give it one owner.
 *
 * A difference is reported as a path with every index erased, so two files that
 * lose the same field under different paragraphs share one signature.
 *
 * The walk reports EVERY distinct difference a file exhibits, not the first.
 * Reporting only the first made the ratchet punish fixes: a file with three
 * losses contributed one signature, so removing that loss revealed the next
 * one, which the baseline had never seen and the gate failed as a new defect.
 * With the whole set reported, fixing a loss can only make a count go down.
 */

import type { Document } from "@stll/folio-core/types/document";

import {
  type CorpusFailure,
  type CorpusInvariant,
  MODEL_TYPE_DISCRIMINATORS,
  failureFromAssertion,
  normalizeFailureMessage,
} from "../corpus-signature";
import { isVocabularyToken } from "./value-vocabulary";

/** Refreshed by every save, or the input bytes themselves: never document content. */
const VOLATILE_KEYS: ReadonlySet<string> = new Set([
  "originalBuffer",
  "modified",
  "lastModifiedBy",
]);

/** The serializer's licence to replay captured bytes, not the content it replays. */
const CAPTURE_KEYS: ReadonlySet<string> = new Set([
  "sourceXml",
  "gridSourceXml",
  "verbatimXml",
  "verbatimFingerprint",
  "rawXml",
  "rawXmlMode",
  "rawImageFingerprint",
  "rawPropertiesXml",
  "rawEndPropertiesXml",
  "rawWatermarkXml",
  "alternateContent",
]);

const isErased = (key: string): boolean => VOLATILE_KEYS.has(key) || CAPTURE_KEYS.has(key);

const normalizeValue = (value: unknown): unknown => {
  if (value instanceof Uint8Array) {
    return { byteLength: value.length };
  }
  if (value instanceof ArrayBuffer) {
    return { byteLength: value.byteLength };
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (value instanceof Map) {
    return [...value.entries()]
      .map(([entryKey, entryValue]) => [String(entryKey), normalizeValue(entryValue)] as const)
      .sort(([left], [right]) => (left < right ? -1 : 1));
  }
  if (Array.isArray(value)) {
    return value.map((item) => normalizeValue(item));
  }
  if (typeof value === "object" && value !== null) {
    const record: Record<string, unknown> = {};
    for (const entryKey of Object.keys(value).sort()) {
      // Erased keys are dropped, never replaced by a sentinel. A sentinel is a
      // value, so it differs from the key being absent, and the comparison
      // would report the erasure itself on every package that carries the slot
      // on one side only.
      if (isErased(entryKey)) {
        continue;
      }
      record[entryKey] = normalizeValue((value as Record<string, unknown>)[entryKey]);
    }
    return record;
  }
  return value;
};

export const normalizeDocumentPackage = (document: Document): unknown =>
  normalizeValue(document.package);

/**
 * The shapes a string is reported as when it is not a token of the format.
 *
 * Quoted, because the slot holds a string and the shape is all that is said
 * about it. The spellings are the ones `normalizeFailureMessage` already erases
 * a message down to, so a reader meets one vocabulary rather than two.
 */
const STRING_SHAPES = {
  url: '"<url>"',
  path: '"<path>"',
  guid: '"<guid>"',
  hex: '"<hex>"',
  id: '"<id>"',
  opaque: '"<string>"',
} as const;

type StringShape = (typeof STRING_SHAPES)[keyof typeof STRING_SHAPES];

/**
 * A token is quoted verbatim; a string longer than the longest token either
 * vocabulary source declares is not one, so the bound caps every residual
 * literal a signature can carry.
 */
const MAX_QUOTED_VALUE_LENGTH = 32;

const URL_RE = /^[a-z][a-z0-9+.-]*:(?:\/\/|[^/])/iu;
const BRACED_GUID_RE = /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/iu;
/**
 * Bounded above as well as below: a `w14:paraId` is eight hex digits and a
 * checksum is not much longer, while a paragraph written entirely in the
 * letters `a` to `f` is text, not an identifier.
 */
const HEX_RUN_RE = /^(?:0x)?[0-9a-f]{8,64}$/iu;
const RELATIONSHIP_ID_RE = /^rId\d+$/u;

/**
 * What a string is, when it is not a token the format defines.
 *
 * Shape, never content. A signature is committed to a public repository, and a
 * corpus document's text, its authors and its file names have no business in
 * one; a shape is also a row that a fixture's wording cannot move.
 */
const shapeOfString = (value: string): StringShape => {
  if (URL_RE.test(value)) {
    return STRING_SHAPES.url;
  }
  if (value.includes("/") || value.includes("\\")) {
    return STRING_SHAPES.path;
  }
  if (BRACED_GUID_RE.test(value)) {
    return STRING_SHAPES.guid;
  }
  if (HEX_RUN_RE.test(value)) {
    return STRING_SHAPES.hex;
  }
  if (RELATIONSHIP_ID_RE.test(value)) {
    return STRING_SHAPES.id;
  }
  return STRING_SHAPES.opaque;
};

const describeString = (value: string): string =>
  isVocabularyToken(value) && value.length <= MAX_QUOTED_VALUE_LENGTH
    ? JSON.stringify(value)
    : shapeOfString(value);

const describeValue = (value: unknown): string => {
  if (value === undefined) {
    return "absent";
  }
  if (value === null) {
    return "null";
  }
  if (typeof value === "string") {
    return describeString(value);
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return Array.isArray(value) ? "array" : "object";
};

/** The side a shape token belongs to, when both sides reduce to the same shape. */
const sided = (shape: string, side: "A" | "B"): string => shape.replace(/>"$/u, ` ${side}>"`);

/**
 * How one value became another.
 *
 * Pair-aware on purpose: two different strings that both reduce to
 * `"<string>"` would otherwise read as a value that did not change. Labelling
 * the sides says the slot holds two different strings without saying what
 * either of them is.
 */
export const describeChange = (left: unknown, right: unknown): string => {
  const before = describeValue(left);
  const after = describeValue(right);
  if (before !== after || typeof left !== "string" || typeof right !== "string") {
    return `${before} became ${after}`;
  }
  return `${sided(before, "A")} became ${sided(after, "B")}`;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** The model discriminator a value carries, or `undefined` when it carries none. */
const discriminatorOf = (value: unknown): string | undefined => {
  if (!isRecord(value)) {
    return undefined;
  }
  const { type } = value;
  return typeof type === "string" && MODEL_TYPE_DISCRIMINATORS.has(type) ? type : undefined;
};

/**
 * The segment that steps into one element of an array.
 *
 * `content[run]` rather than `content[]`: the same field lost under a run and
 * under the paragraph that holds it are two defects with two owners, and a
 * path that names only the field reports them as one. The discriminator is
 * read from the side that has one — the two agree except where the element's
 * kind is itself the difference, and that difference gets its own row under
 * `.type`.
 */
const elementSegment = (key: string, left: unknown, right: unknown): string =>
  `${key}[${discriminatorOf(left) ?? discriminatorOf(right) ?? ""}]`;

/**
 * How many distinct differences one file may report for one invariant.
 *
 * A pathological package can differ in thousands of places, and a census that
 * carried them all would be a transcript of that one file rather than a list of
 * defects. Past the bound the file reports {@link omittedDifferencesMessage}
 * instead, which is a signature of its own: a file that exceeds the cap is a
 * fact worth ratcheting, and it cannot be mistaken for a defect.
 */
export const MAX_REPORTED_DIFFERENCES = 64;

export const omittedDifferencesMessage = (omitted: number): string =>
  `…and ${omitted} more differences past the reporting cap`;

/**
 * Differences are deduplicated by the signature they will become, not by the
 * text they are now: `paraId: "<hex>" became "<hex>"` under two hundred
 * comments is one defect, and counting it two hundred times towards the cap
 * would spend the whole budget on one row.
 */
type DifferenceCollector = {
  readonly seen: Set<string>;
  readonly messages: string[];
  omitted: number;
};

const record = (collector: DifferenceCollector, message: string): void => {
  const key = normalizeFailureMessage(message);
  if (collector.seen.has(key)) {
    return;
  }
  collector.seen.add(key);
  if (collector.messages.length < MAX_REPORTED_DIFFERENCES) {
    collector.messages.push(message);
    return;
  }
  collector.omitted += 1;
};

/**
 * Every difference between two normalised packages.
 *
 * Array positions collapse to the element's kind: a field lost under the
 * twelfth paragraph and the same field lost under the third are one defect,
 * and the file that shows it is in the census example. Which kind held it is
 * not an index but an owner, so it stays.
 *
 * An array whose length changed is reported and not descended into. Comparing
 * two arrays of different lengths index by index reports the shift rather than
 * the loss, and inventing that noise is a worse answer than the one row. Which
 * index diverged is a separate gap, owned by the array comparison itself. The
 * row is untyped for the same reason: a length is the array's, not any one
 * element's.
 */
const collectDifferences = (
  left: unknown,
  right: unknown,
  path: string,
  collector: DifferenceCollector,
): void => {
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) {
      record(collector, `${path}: ${describeChange(left, right)}`);
      return;
    }
    if (left.length !== right.length) {
      record(collector, `${path}[]: length changed`);
      return;
    }
    for (const [index, item] of left.entries()) {
      const other = right[index];
      collectDifferences(item, other, elementSegment(path, item, other), collector);
    }
    return;
  }

  if (isRecord(left) && isRecord(right)) {
    const keys = [...new Set([...Object.keys(left), ...Object.keys(right)])].sort();
    for (const key of keys) {
      collectDifferences(left[key], right[key], `${path}.${key}`, collector);
    }
    return;
  }

  if (left === right) {
    return;
  }
  record(collector, `${path}: ${describeChange(left, right)}`);
};

export type PackageDifferences = {
  /** Distinct difference messages, in walk order, at most {@link MAX_REPORTED_DIFFERENCES}. */
  messages: readonly string[];
  /** Distinct differences the walk found past the cap and did not report. */
  omitted: number;
};

/**
 * What changed between two parsed packages, as messages with no per-file
 * particulars. Empty when they agree.
 */
export const describePackageDifferences = (
  before: Document,
  after: Document,
): PackageDifferences => {
  const collector: DifferenceCollector = { seen: new Set(), messages: [], omitted: 0 };
  collectDifferences(
    normalizeDocumentPackage(before),
    normalizeDocumentPackage(after),
    "package",
    collector,
  );
  return { messages: collector.messages, omitted: collector.omitted };
};

/**
 * One comparison's failures, the overflow marker included.
 *
 * The three invariants that compare packages phrase a difference differently —
 * `reserialize` even phrases two differences of one file differently — so the
 * wording is the caller's. Whether a capped file says so is not: one place
 * decides that, or the three would drift about it.
 */
export const differenceFailures = (
  invariant: CorpusInvariant,
  { messages, omitted }: PackageDifferences,
  describe: (message: string) => string,
): CorpusFailure[] => {
  const failures = messages.map((message) => failureFromAssertion(invariant, describe(message)));
  if (omitted > 0) {
    failures.push(failureFromAssertion(invariant, omittedDifferencesMessage(omitted)));
  }
  return failures;
};
