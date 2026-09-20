/**
 * Losses the container contract already decided to accept.
 *
 * A baseline row says "this many files fail this way". It does not say whether
 * anybody looked. Most rows are defects nobody has reached yet; a few are the
 * recorded consequence of a decision `docs/container-contract.md` states in
 * prose, and reading them as unexamined defects makes the list longer than it
 * is and buries a real regression among rows nobody intends to fix.
 *
 * `corpus/expected-refusals.json` is the wrong home for them. Its contract is
 * "packages folio deliberately declines" — a typed error, no model, nothing
 * measured — and its check treats a signature as either a refusal or a defect,
 * never both. A disposition is the opposite case: folio parsed the file, built
 * the model, and lost a field on a leg the contract says has no carrier yet.
 *
 * So dispositions get their own committed file. Each entry names what it
 * claims, why, where the contract says so, and what would retire it. The check
 * reports them under their own heading with counts, and ratchets them the same
 * way everything else here ratchets: growth fails, because growth means the
 * class widened past the decision; a shrink has to be written down; and an
 * entry nothing matches any more is a decision about markup folio no longer
 * produces, so it fails until it is deleted.
 *
 * The file is hand-written except for `fileHits`, exactly as the refusal list
 * is hand-written except for `files`. A reason is a decision a person makes,
 * and no rewrite promotes a signature on its own.
 */

import type { CensusSignature, CorpusCensus } from "./corpus-census";
import type { ExpectedRefusals } from "./corpus-refusals";
import {
  type CorpusInvariant,
  ELISION,
  MODEL_PATH_RE,
  MODEL_TYPE_DISCRIMINATORS,
} from "./corpus-signature";

/**
 * How an entry says which signatures it claims.
 *
 * `signature` is the whole string and claims exactly one row. `path` claims a
 * class: the model path a difference message carries, with `**` standing for
 * any run of segments, restricted to the invariants whose leg the decision is
 * about. A disposition is a statement about an owner in the model, and an
 * owner appears under many paths, so a class needs a pattern rather than the
 * sixty-five rows one wave of the corpus happened to produce.
 *
 * An array segment names the kind it steps into, so a pattern says which kind
 * it means: `content[run]` claims runs, `content[*]` claims any element of a
 * `content` array, and `content[]` claims only the elements the model gives no
 * discriminator. An owner is what a disposition is about, so naming it is what
 * makes the claim exact rather than a guess from the shape of the path.
 */
export type DispositionMatch =
  | { kind: "signature"; signature: string }
  | { kind: "path"; invariants: readonly CorpusInvariant[]; paths: readonly string[] };

export type ExpectedDispositionEntry = {
  /** A stable handle, so a report and a review can name the same decision. */
  id: string;
  match: DispositionMatch;
  /** Why this loss is the contract working, in a sentence a reviewer can weigh. */
  reason: string;
  /** Where `docs/` states it, as a path and anchor. */
  contract: string;
  /** What has to exist before this entry is deleted rather than carried. */
  removalCondition: string;
  /**
   * Matched rows' file counts, added up.
   *
   * File hits, not files: one file can exhibit the same decision under a
   * paragraph and under a table cell, which is two rows and two counts. The
   * ratchet only needs a number that cannot grow without the class widening,
   * and a sum of per-signature counts is that number.
   */
  fileHits: number;
};

export type ExpectedDispositions = {
  schemaVersion: 1;
  entries: ExpectedDispositionEntry[];
};

/**
 * The model path a difference message carries, or `undefined`.
 *
 * The pattern is the normaliser's, so the path a disposition is matched
 * against is the one the normaliser shortened.
 */
export const differencePath = (message: string): string | undefined =>
  MODEL_PATH_RE.exec(message)?.[0];

/** The discriminator a pattern segment writes when it means "whatever kind". */
const ANY_DISCRIMINATOR = "*";

/** An array segment split into the field it reads and the kind it steps into. */
const ARRAY_SEGMENT_RE = /^(?<field>[^[\]]*)\[(?<kind>[^[\]]*)\]$/u;

/** The field an `array[*]` pattern reads, or `undefined` for every other segment. */
const anyKindField = (patternSegment: string): string | undefined => {
  const groups = ARRAY_SEGMENT_RE.exec(patternSegment)?.groups;
  return groups?.["kind"] === ANY_DISCRIMINATOR ? groups["field"] : undefined;
};

const matchesWholeSegment = (patternSegment: string, pathSegment: string): boolean => {
  const field = anyKindField(patternSegment);
  if (field === undefined) {
    return patternSegment === pathSegment;
  }
  return ARRAY_SEGMENT_RE.exec(pathSegment)?.groups?.["field"] === field;
};

const matchesCutSegment = (patternSegment: string, prefix: string): boolean => {
  const field = anyKindField(patternSegment);
  if (field === undefined) {
    return patternSegment.startsWith(prefix);
  }
  const opener = `${field}[`;
  return prefix.startsWith(opener) || opener.startsWith(prefix);
};

/**
 * One segment against one pattern segment.
 *
 * A message is capped, so the last segment of a long path can arrive cut.
 * `…content[run].preservedAttribu…` is the same decision as the row that fit,
 * and refusing to match it would split one class by message length. A cut
 * segment matches a pattern segment it is a prefix of, which is as much as the
 * evidence supports: a path cut before its leaf matches nothing, and the bare
 * `…` a shortened path leaves in place of its middle matches only `**`.
 */
const segmentMatches = (patternSegment: string, pathSegment: string): boolean => {
  if (!pathSegment.endsWith(ELISION)) {
    return matchesWholeSegment(patternSegment, pathSegment);
  }
  const prefix = pathSegment.slice(0, -ELISION.length);
  return prefix.length > 0 && matchesCutSegment(patternSegment, prefix);
};

/** `**` against a run of segments, `*` never: a path has no partial segments. */
const matchSegments = (
  pattern: readonly string[],
  segments: readonly string[],
  patternIndex: number,
  segmentIndex: number,
): boolean => {
  const head = pattern[patternIndex];
  if (head === undefined) {
    return segmentIndex === segments.length;
  }
  if (head === "**") {
    for (let skip = segmentIndex; skip <= segments.length; skip += 1) {
      if (matchSegments(pattern, segments, patternIndex + 1, skip)) {
        return true;
      }
    }
    return false;
  }
  const segment = segments[segmentIndex];
  if (segment === undefined || !segmentMatches(head, segment)) {
    return false;
  }
  return matchSegments(pattern, segments, patternIndex + 1, segmentIndex + 1);
};

export const pathMatchesPattern = (pattern: string, modelPath: string): boolean =>
  matchSegments(pattern.split("."), modelPath.split("."), 0, 0);

const assertNever = (value: never): never => {
  throw new Error(`unhandled disposition match: ${JSON.stringify(value)}`);
};

/**
 * What a match is decided on.
 *
 * Every row the gate compares carries all three: a census signature, a core
 * baseline row and a family baseline row alike.
 */
export type DispositionCandidate = {
  signature: string;
  /** Widened: a candidate may come from a stored signature nobody re-typed. */
  invariant: string;
  message: string;
};

const claims = (match: DispositionMatch, candidate: DispositionCandidate): boolean => {
  switch (match.kind) {
    case "signature": {
      return candidate.signature === match.signature;
    }
    case "path": {
      if (!match.invariants.some((invariant) => invariant === candidate.invariant)) {
        return false;
      }
      const modelPath = differencePath(candidate.message);
      return (
        modelPath !== undefined &&
        match.paths.some((pattern) => pathMatchesPattern(pattern, modelPath))
      );
    }
    default: {
      return assertNever(match);
    }
  }
};

const SIGNATURE_SEPARATOR = " | ";

/**
 * A candidate rebuilt from a signature alone, for a list that stores no more.
 *
 * `failureSignature` joins three fields with the same separator and only the
 * frame may be empty, so the invariant is everything before the first join and
 * the message everything between the first and the last.
 */
export const candidateFromSignature = (signature: string): DispositionCandidate => {
  const first = signature.indexOf(SIGNATURE_SEPARATOR);
  const last = signature.lastIndexOf(SIGNATURE_SEPARATOR);
  if (first === -1 || last === first) {
    return { signature, invariant: "", message: "" };
  }
  return {
    signature,
    invariant: signature.slice(0, first),
    message: signature.slice(first + SIGNATURE_SEPARATOR.length, last),
  };
};

/** The entry that claims a row, or `undefined` when it is an ordinary defect. */
export const dispositionOf = (
  dispositions: ExpectedDispositions,
  candidate: DispositionCandidate,
): ExpectedDispositionEntry | undefined =>
  dispositions.entries.find((entry) => claims(entry.match, candidate));

export type PartitionedDispositions = {
  /** The census the baseline ratchets against: rows no entry claims. */
  defects: CorpusCensus;
  dispositions: CensusSignature[];
};

export const partitionExpectedDispositions = (
  census: CorpusCensus,
  dispositions: ExpectedDispositions,
): PartitionedDispositions => {
  const defects: CensusSignature[] = [];
  const claimed: CensusSignature[] = [];
  for (const signature of census.signatures) {
    (dispositionOf(dispositions, signature) === undefined ? defects : claimed).push(signature);
  }
  return { defects: { ...census, signatures: defects }, dispositions: claimed };
};

/**
 * Rows a disposition claims, removed from a baseline or a census.
 *
 * Both sides of every comparison go through this, so a claimed row is missing
 * from the recorded side and the observed side at once. That is what lets the
 * entries land before the baselines are re-measured: the rows stay in the
 * committed files, nothing reads them as resolved, and the next
 * `write-baseline` drops them because the census it is written from no longer
 * carries them.
 */
export const withoutDispositions = <T extends DispositionCandidate>(
  rows: readonly T[],
  dispositions: ExpectedDispositions,
): T[] => rows.filter((row) => dispositionOf(dispositions, row) === undefined);

const hitsByEntry = (
  dispositions: ExpectedDispositions,
  observed: readonly CensusSignature[],
): Map<string, number> => {
  const hits = new Map(dispositions.entries.map((entry) => [entry.id, 0]));
  for (const signature of observed) {
    const entry = dispositionOf(dispositions, signature);
    if (entry !== undefined) {
      hits.set(entry.id, (hits.get(entry.id) ?? 0) + signature.files);
    }
  }
  return hits;
};

/**
 * The list with every count refreshed from what the run observed.
 *
 * Reasons, contracts, removal conditions and the set of entries are preserved
 * exactly: this refreshes a ratchet, it does not decide what belongs on it.
 */
export const refreshedExpectedDispositions = (
  dispositions: ExpectedDispositions,
  observed: readonly CensusSignature[],
): ExpectedDispositions => {
  const hits = hitsByEntry(dispositions, observed);
  return {
    schemaVersion: 1,
    entries: dispositions.entries
      .map((entry) => ({ ...entry, fileHits: hits.get(entry.id) ?? 0 }))
      .sort((left, right) => (left.id < right.id ? -1 : 1)),
  };
};

export type DispositionViolation = {
  kind: "more-files" | "fewer-files" | "resolved-disposition";
  signature: string;
  detail: string;
};

export const compareToExpectedDispositions = (
  dispositions: ExpectedDispositions,
  observed: readonly CensusSignature[],
): DispositionViolation[] => {
  const hits = hitsByEntry(dispositions, observed);
  const violations: DispositionViolation[] = [];
  for (const entry of dispositions.entries) {
    const fileHits = hits.get(entry.id) ?? 0;
    if (fileHits > entry.fileHits) {
      violations.push({
        kind: "more-files",
        signature: entry.id,
        detail: `${fileHits} file hits carry this disposition, the entry allows ${entry.fileHits}; the class widened past ${entry.contract}`,
      });
      continue;
    }
    if (fileHits === 0) {
      violations.push({
        kind: "resolved-disposition",
        signature: entry.id,
        detail:
          "nothing matches it any more; delete it from corpus/expected-dispositions.json rather than carrying a decision about markup folio no longer produces",
      });
      continue;
    }
    if (fileHits < entry.fileHits) {
      violations.push({
        kind: "fewer-files",
        signature: entry.id,
        detail: `${fileHits} file hits carry this disposition, down from ${entry.fileHits}; rerun with \`write-baseline\``,
      });
    }
  }
  return violations;
};

export const renderExpectedDispositions = (
  dispositions: ExpectedDispositions,
  observed: readonly CensusSignature[],
): string => {
  if (dispositions.entries.length === 0) {
    return "";
  }
  const hits = hitsByEntry(dispositions, observed);
  const rowsByEntry = new Map<string, number>();
  for (const signature of observed) {
    const entry = dispositionOf(dispositions, signature);
    if (entry !== undefined) {
      rowsByEntry.set(entry.id, (rowsByEntry.get(entry.id) ?? 0) + 1);
    }
  }
  return [
    `  known dispositions ${dispositions.entries.length} (reported, not ratcheted as defects):`,
    ...dispositions.entries.flatMap((entry) => [
      `  ${String(hits.get(entry.id) ?? 0).padStart(5)} file hits across ${rowsByEntry.get(entry.id) ?? 0} signature(s)  ${entry.id}`,
      `         ${entry.reason}`,
      `         contract: ${entry.contract}`,
      `         retires when: ${entry.removalCondition}`,
    ]),
  ].join("\n");
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const ENTRY_KEYS = new Set(["contract", "fileHits", "id", "match", "reason", "removalCondition"]);
const LIST_KEYS = new Set(["entries", "schemaVersion"]);
const HAND_WRITTEN_KEYS = ["reason", "contract", "removalCondition"] as const;

/**
 * The kinds a pattern names that the model does not declare.
 *
 * A mistyped kind claims nothing, and a disposition that claims nothing only
 * surfaces a nightly later, as an entry to delete. The vocabulary is closed
 * and committed, so the typo is catchable where the file is read.
 */
const unknownKindsIn = (pattern: string): string[] =>
  pattern
    .split(".")
    .map((segment) => ARRAY_SEGMENT_RE.exec(segment)?.groups?.["kind"])
    .filter(
      (kind): kind is string =>
        kind !== undefined &&
        kind.length > 0 &&
        kind !== ANY_DISCRIMINATOR &&
        !MODEL_TYPE_DISCRIMINATORS.has(kind),
    );

const validateMatch = (value: unknown, location: string, issues: string[]): void => {
  if (!isRecord(value)) {
    issues.push(`${location}: expected an object`);
    return;
  }
  if (value["kind"] === "signature") {
    const { signature } = value;
    if (typeof signature !== "string" || signature.length === 0) {
      issues.push(`${location}.signature: expected a non-empty string`);
    }
    return;
  }
  if (value["kind"] !== "path") {
    issues.push(`${location}.kind: expected \`signature\` or \`path\``);
    return;
  }
  for (const key of ["invariants", "paths"]) {
    const list = value[key];
    if (!Array.isArray(list) || list.length === 0) {
      issues.push(`${location}.${key}: expected a non-empty array`);
      continue;
    }
    for (const [index, item] of list.entries()) {
      if (typeof item !== "string" || item.length === 0) {
        issues.push(`${location}.${key}[${index}]: expected a non-empty string`);
      }
    }
  }
  const { paths } = value;
  if (Array.isArray(paths)) {
    for (const [index, pattern] of paths.entries()) {
      if (typeof pattern !== "string") {
        continue;
      }
      if (!pattern.startsWith("package.")) {
        issues.push(`${location}.paths[${index}]: a model path starts at \`package\``);
      }
      for (const kind of unknownKindsIn(pattern)) {
        issues.push(
          `${location}.paths[${index}]: \`${kind}\` is not a kind the model declares; a pattern that names none claims nothing`,
        );
      }
    }
  }
};

/**
 * Everything wrong with the committed file, plus every overlap with the refusal
 * list.
 *
 * The two lists answer different questions and a signature in both would be
 * excused twice and ratcheted under two rules at once. Which one is right is a
 * decision, so the check refuses rather than picking.
 */
export const validateExpectedDispositions = (
  value: unknown,
  refusals: ExpectedRefusals,
): string[] => {
  const issues: string[] = [];
  if (!isRecord(value)) {
    return ["dispositions: expected an object"];
  }
  for (const key of Object.keys(value)) {
    if (!LIST_KEYS.has(key)) {
      issues.push(`${key}: unknown field`);
    }
  }
  if (value["schemaVersion"] !== 1) {
    issues.push("schemaVersion: expected 1");
  }
  const { entries } = value;
  if (!Array.isArray(entries)) {
    issues.push("entries: expected an array");
    return issues;
  }
  const seen = new Set<string>();
  for (const [index, entry] of entries.entries()) {
    const location = `entries[${index}]`;
    if (!isRecord(entry)) {
      issues.push(`${location}: expected an object`);
      continue;
    }
    for (const key of Object.keys(entry)) {
      if (!ENTRY_KEYS.has(key)) {
        issues.push(`${location}.${key}: unknown field`);
      }
    }
    const { id, fileHits } = entry;
    if (typeof id !== "string" || id.length === 0) {
      issues.push(`${location}.id: expected a non-empty string`);
    } else {
      if (seen.has(id)) {
        issues.push(`${location}.id: duplicate id \`${id}\``);
      }
      seen.add(id);
    }
    for (const key of HAND_WRITTEN_KEYS) {
      const field = entry[key];
      if (typeof field !== "string" || field.trim().length === 0) {
        issues.push(`${location}.${key}: expected a non-empty ${key}`);
      }
    }
    if (!Number.isInteger(fileHits) || Number(fileHits) < 1) {
      issues.push(`${location}.fileHits: expected a positive integer`);
    }
    validateMatch(entry["match"], `${location}.match`, issues);
  }

  if (issues.length > 0) {
    return issues;
  }
  // SAFETY: every field was just checked, field by field, against this shape.
  const checked = value as unknown as ExpectedDispositions;
  for (const refusal of refusals.entries) {
    const entry = dispositionOf(checked, candidateFromSignature(refusal.signature));
    if (entry !== undefined) {
      issues.push(
        `${entry.id} also claims the expected refusal ${refusal.signature}; a signature is a refusal or a disposition, never both`,
      );
    }
  }
  return issues;
};
