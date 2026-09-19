/**
 * Saving must reach a fixed point, and reach it after one save.
 *
 * The first save is never a no-op, by design: every repack runs
 * `normalizeRevisionIdsInXmlParts` and `normalizeParaIdRangeInXmlParts` over
 * every `word/*.xml` part and `normalizeAppVersionInZip` over
 * `docProps/app.xml`, including parts the save itself did not touch, because a
 * revision id and a paragraph id are package-wide facts no single serializer
 * can bound. Comparing the input package against its first save would therefore
 * report that normalisation rather than any defect. What the format does owe is
 * that the normalisation converges: once folio has written a package, writing
 * it again must produce the same parts.
 *
 * The comparison is at the part level rather than over the model, so it catches
 * what a model comparison cannot: a part folio does not model drifting on every
 * save, an id space that renumbers itself each pass, a relationship or content
 * type reconciled differently the second time. Whole-package bytes are not
 * comparable (the archive carries timestamps and compression state), so parts
 * are compared by content.
 */

import { parseDocx } from "@stll/folio-core/docx/parser";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import { unzipDocx } from "@stll/folio-core/docx/unzip";
import type { Document } from "@stll/folio-core/types/document";
import { Result } from "better-result";

import { type CorpusFailure, failureFromAssertion, failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";

const DIGIT_RUN_RE = /\d+/gu;

/**
 * A part path with its per-package particulars erased.
 *
 * Two files that lose `word/header2.xml` and `word/header7.xml` are one defect
 * about headers, not two about numbers, so every digit run collapses to `N`.
 * A path that carries no number (`docProps/app.xml`) is already general.
 */
export const generalizePartPath = (path: string): string => path.replaceAll(DIGIT_RUN_RE, "N");

const PART_KINDS = {
  xml: "xml",
  binary: "binary",
} as const;

type PackagePart =
  | { kind: typeof PART_KINDS.xml; text: string }
  | { kind: typeof PART_KINDS.binary; bytes: Uint8Array };

/**
 * Every entry of a saved package, keyed by path.
 *
 * `allXml` holds the decoded text of the parts the reader interprets; anything
 * else comes off the archive as bytes, so a macro project or an embedded
 * workbook is compared too.
 */
const readPackageParts = async (buffer: ArrayBuffer): Promise<Map<string, PackagePart>> => {
  const raw = await unzipDocx(buffer, { extractAllXml: true });
  const entries = Object.entries(raw.originalZip.files).filter(([, file]) => !file.dir);
  return new Map(
    await Promise.all(
      entries.map(async ([path, file]): Promise<readonly [string, PackagePart]> => {
        const xml = raw.allXml.get(path);
        if (xml !== undefined) {
          return [path, { kind: PART_KINDS.xml, text: xml }];
        }
        return [path, { kind: PART_KINDS.binary, bytes: await file.async("uint8array") }];
      }),
    ),
  );
};

const sameBytes = (first: Uint8Array, second: Uint8Array): boolean => {
  if (first.length !== second.length) {
    return false;
  }
  return first.every((byte, index) => byte === second[index]);
};

const partsDiffer = (first: PackagePart, second: PackagePart): boolean => {
  switch (first.kind) {
    case PART_KINDS.xml:
      return second.kind !== PART_KINDS.xml || first.text !== second.text;
    case PART_KINDS.binary:
      return second.kind !== PART_KINDS.binary || !sameBytes(first.bytes, second.bytes);
    default: {
      const unreachable: never = first;
      return unreachable;
    }
  }
};

const PART_DIFFERENCE_KINDS = {
  missing: "missing",
  added: "added",
  changed: "changed",
} as const;

type PartDifferenceKind = (typeof PART_DIFFERENCE_KINDS)[keyof typeof PART_DIFFERENCE_KINDS];

/** What each difference says, with the part path appended and nothing else. */
const PART_DIFFERENCE_MESSAGES = {
  [PART_DIFFERENCE_KINDS.missing]:
    "a part present after the first save is missing after the second",
  [PART_DIFFERENCE_KINDS.added]: "a part appears only after the second save",
  [PART_DIFFERENCE_KINDS.changed]: "a part changed between the first and the second save",
} as const satisfies Record<PartDifferenceKind, string>;

type PartDifference = { kind: PartDifferenceKind; path: string };

/**
 * How many differing parts one file may report.
 *
 * A package whose every part drifts is one defect; letting it name fifty parts
 * would spend the baseline's whole budget on a single file.
 */
const MAX_REPORTED_PART_DIFFERENCES = 5;

const findPartDifferences = (
  first: Map<string, PackagePart>,
  second: Map<string, PackagePart>,
): PartDifference[] => {
  const differences: PartDifference[] = [];
  for (const [path, part] of first) {
    const counterpart = second.get(path);
    if (counterpart === undefined) {
      differences.push({ kind: PART_DIFFERENCE_KINDS.missing, path });
      continue;
    }
    if (partsDiffer(part, counterpart)) {
      differences.push({ kind: PART_DIFFERENCE_KINDS.changed, path });
    }
  }
  for (const path of second.keys()) {
    if (!first.has(path)) {
      differences.push({ kind: PART_DIFFERENCE_KINDS.added, path });
    }
  }
  // Archive order is the package's, not the defect's: sorting makes the capped
  // report the same on every run over the same file.
  return differences.sort((left, right) => (left.path < right.path ? -1 : 1));
};

const failuresForDifferences = (differences: PartDifference[]): CorpusFailure[] => {
  const messages = new Set(
    differences
      .slice(0, MAX_REPORTED_PART_DIFFERENCES)
      .map(({ kind, path }) => `${PART_DIFFERENCE_MESSAGES[kind]}: ${generalizePartPath(path)}`),
  );
  return [...messages].map((message) =>
    failureFromAssertion(EXTENDED_CORPUS_INVARIANTS.saveIdempotence, message),
  );
};

const save = (document: Document): Promise<ArrayBuffer> =>
  repackDocx(document, { updateModifiedDate: false });

export const runSaveIdempotenceInvariant = async ({
  parsed,
}: CorpusInvariantInput): Promise<CorpusInvariantOutcome> => {
  const timings: Record<string, number> = {};

  const first = await timeStage(timings, "first-save", () =>
    Result.tryPromise({ try: () => save(parsed), catch: (cause: unknown) => cause }),
  );
  if (first.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.saveIdempotence, first.error)],
      timings,
    };
  }

  const reparsed = await timeStage(timings, "reparse", () =>
    Result.tryPromise({
      try: () => parseDocx(first.value, { preloadFonts: false }),
      catch: (cause: unknown) => cause,
    }),
  );
  if (reparsed.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.saveIdempotence, reparsed.error)],
      timings,
    };
  }

  const second = await timeStage(timings, "second-save", () =>
    Result.tryPromise({ try: () => save(reparsed.value), catch: (cause: unknown) => cause }),
  );
  if (second.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.saveIdempotence, second.error)],
      timings,
    };
  }

  const parts = await timeStage(timings, "unzip", () =>
    Result.tryPromise({
      try: () => Promise.all([readPackageParts(first.value), readPackageParts(second.value)]),
      catch: (cause: unknown) => cause,
    }),
  );
  if (parts.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.saveIdempotence, parts.error)],
      timings,
    };
  }

  const [firstParts, secondParts] = parts.value;
  const differences = await timeStage(timings, "compare", () =>
    Promise.resolve(findPartDifferences(firstParts, secondParts)),
  );
  return { failures: failuresForDifferences(differences), timings };
};
