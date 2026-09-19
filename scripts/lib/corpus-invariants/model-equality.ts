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
 */

import type { Document } from "@stll/folio-core/types/document";

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
 * Values small enough to name in a signature.
 *
 * A boolean or a short token is the defect (`"start"` became `"left"`); a
 * paragraph's text is the file. Anything longer than this is reported by type
 * alone.
 */
const MAX_QUOTED_VALUE_LENGTH = 24;

const describeValue = (value: unknown): string => {
  if (value === undefined) {
    return "absent";
  }
  if (value === null) {
    return "null";
  }
  if (typeof value === "string") {
    return value.length <= MAX_QUOTED_VALUE_LENGTH ? JSON.stringify(value) : "string";
  }
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  return Array.isArray(value) ? "array" : "object";
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/**
 * The first difference between two normalised packages, or null.
 *
 * Array positions collapse to `[]`: a field lost under the twelfth paragraph
 * and the same field lost under the third are one defect, and the file that
 * shows it is in the census example.
 */
const findDifference = (left: unknown, right: unknown, path: string): string | null => {
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) {
      return `${path}: ${describeValue(left)} became ${describeValue(right)}`;
    }
    if (left.length !== right.length) {
      return `${path}[]: length changed`;
    }
    for (const [index, item] of left.entries()) {
      const difference = findDifference(item, right[index], `${path}[]`);
      if (difference !== null) {
        return difference;
      }
    }
    return null;
  }

  if (isRecord(left) && isRecord(right)) {
    const keys = [...new Set([...Object.keys(left), ...Object.keys(right)])].sort();
    for (const key of keys) {
      const difference = findDifference(left[key], right[key], `${path}.${key}`);
      if (difference !== null) {
        return difference;
      }
    }
    return null;
  }

  if (left === right) {
    return null;
  }
  return `${path}: ${describeValue(left)} became ${describeValue(right)}`;
};

/**
 * What changed between two parsed packages, as a message with no per-file
 * particulars, or null when they agree.
 */
export const describePackageDifference = (before: Document, after: Document): string | null =>
  findDifference(normalizeDocumentPackage(before), normalizeDocumentPackage(after), "package");
