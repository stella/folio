/**
 * The invariants the corpus gate asserts on one public `.docx`.
 *
 * Every one of them is something Word does without complaint, so a failure is a
 * folio defect, never a verdict on the file:
 *
 * 1. `parse` — a package a reader can open must parse.
 * 2. `fixed-point` — parse → ProseMirror → repack → parse must preserve the
 *    visible text and the text-block count. This is the equality
 *    `packages/core/src/docx/__tests__/corpusFixedPoint.test.ts` defines over the
 *    committed fixtures; the corpus gate runs the same one over public files.
 * 3. `repack-validates` — the package folio just wrote must satisfy folio's own
 *    package validator. A gate that only round-trips would accept writing a
 *    package folio itself rejects.
 * 4. `style-set-rebuild` — extracting a document's style set and building a new
 *    package from it must not panic. This is the narrow path that turns a
 *    document's styles into a reusable template, and it reaches serializer
 *    assertions no round trip touches.
 */

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx, repackDocx, validateDocx } from "@stll/folio-core/docx/rezip";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { extractDocumentStyleSet } from "@stll/folio-core/style-sets/extract";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { Result } from "better-result";

import { classifyCorpusFile, type NotADocxReason } from "./corpus-classify";
import { runExtendedChecks } from "./corpus-extended";
import { PRODUCER_FAMILIES } from "./corpus-producer";
import {
  CORPUS_INVARIANTS,
  type CorpusFailure,
  type CorpusInvariant,
  failureFromAssertion,
  failureFromError,
} from "./corpus-signature";

/** What a checked file cost, for the performance census. */
export type CorpusCheckCost = { bytes: number; parseMs: number; peakRssBytes: number };

export type CorpusCheckResult =
  | { kind: "not-a-docx"; reason: NotADocxReason; detail: string }
  | {
      kind: "checked";
      failures: CorpusFailure[];
      producer: string;
      cost: CorpusCheckCost;
      timings: Record<string, number>;
    };

const STYLE_SET_NAME = "corpus-gate";

/** A file that did not parse never reached the part that names its producer. */
const UNKNOWN_PRODUCER: string = PRODUCER_FAMILIES.unknown;

/** The ProseMirror document `toProseDoc` produces, without naming its package from here. */
type ProseDocument = ReturnType<typeof toProseDoc>;

const countTextBlocks = (doc: ProseDocument): number => {
  let count = 0;
  doc.descendants((node) => {
    if (node.isTextblock) {
      count += 1;
    }
    return true;
  });
  return count;
};

const toArrayBuffer = (bytes: Uint8Array): ArrayBuffer =>
  bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;

type FixedPointOutcome =
  | { kind: "failed"; failures: CorpusFailure[] }
  | { kind: "repacked"; repacked: ArrayBuffer; failures: CorpusFailure[] };

const checkFixedPoint = async (parsed: Document): Promise<FixedPointOutcome> => {
  const original = Result.try(() => toProseDoc(parsed));
  if (original.isErr()) {
    return {
      kind: "failed",
      failures: [failureFromError(CORPUS_INVARIANTS.fixedPoint, original.error)],
    };
  }

  const repacked = await Result.tryPromise({
    try: async () => {
      const back = fromProseDoc(original.value, parsed);
      return await repackDocx(back, { updateModifiedDate: false });
    },
    catch: (cause: unknown) => cause,
  });
  if (repacked.isErr()) {
    return {
      kind: "failed",
      failures: [failureFromError(CORPUS_INVARIANTS.fixedPoint, repacked.error)],
    };
  }

  const reparsed = await Result.tryPromise({
    try: async () => toProseDoc(await parseDocx(repacked.value, { preloadFonts: false })),
    catch: (cause: unknown) => cause,
  });
  if (reparsed.isErr()) {
    return {
      kind: "repacked",
      repacked: repacked.value,
      failures: [failureFromError(CORPUS_INVARIANTS.fixedPoint, reparsed.error)],
    };
  }

  const failures: CorpusFailure[] = [];
  if (reparsed.value.textContent !== original.value.textContent) {
    failures.push(
      failureFromAssertion(
        CORPUS_INVARIANTS.fixedPoint,
        `visible text changed (${original.value.textContent.length} characters in, ${reparsed.value.textContent.length} out)`,
      ),
    );
  }
  const blocksIn = countTextBlocks(original.value);
  const blocksOut = countTextBlocks(reparsed.value);
  if (blocksIn !== blocksOut) {
    failures.push(
      failureFromAssertion(
        CORPUS_INVARIANTS.fixedPoint,
        `text-block count changed (${blocksIn} in, ${blocksOut} out)`,
      ),
    );
  }
  return { kind: "repacked", repacked: repacked.value, failures };
};

const checkRepackValidates = async (repacked: ArrayBuffer): Promise<CorpusFailure[]> => {
  const validated = await Result.tryPromise({
    try: () => validateDocx(repacked),
    catch: (cause: unknown) => cause,
  });
  if (validated.isErr()) {
    return [failureFromError(CORPUS_INVARIANTS.repackValidates, validated.error)];
  }
  if (validated.value.valid) {
    return [];
  }
  return [
    failureFromAssertion(
      CORPUS_INVARIANTS.repackValidates,
      validated.value.errors.join("; ") || "the validator reported no reason",
    ),
  ];
};

const checkStyleSetRebuild = async (parsed: Document): Promise<CorpusFailure[]> => {
  const rebuilt = await Result.tryPromise({
    try: async () => {
      const styleSet = extractDocumentStyleSet(parsed, { name: STYLE_SET_NAME });
      return await createDocx(createEmptyDocument({ styleSet }));
    },
    catch: (cause: unknown) => cause,
  });
  return rebuilt.isErr()
    ? [failureFromError(CORPUS_INVARIANTS.styleSetRebuild, rebuilt.error)]
    : [];
};

/**
 * Run every invariant that the file reaches.
 *
 * A package that does not parse yields one failure and stops: the later
 * invariants have no model to run against, and reporting them would inflate
 * every signature a parse defect causes.
 */
export type RunCorpusChecksOptions = {
  invariantBudgetMs: number;
  fileBudgetMs: number;
  /** Run only these invariants; `parse` always runs, because nothing else can without it. */
  only?: ReadonlySet<CorpusInvariant>;
};

export const runCorpusChecks = async (
  bytes: Uint8Array,
  { invariantBudgetMs, fileBudgetMs, only }: RunCorpusChecksOptions,
): Promise<CorpusCheckResult> => {
  const classification = await classifyCorpusFile(bytes);
  if (classification.kind === "not-a-docx") {
    return classification;
  }

  const buffer = toArrayBuffer(bytes);
  const parseStarted = Bun.nanoseconds();
  const parsed = await Result.tryPromise({
    try: () => parseDocx(buffer, { preloadFonts: false }),
    catch: (cause: unknown) => cause,
  });
  const cost = {
    bytes: bytes.byteLength,
    parseMs: (Bun.nanoseconds() - parseStarted) / 1e6,
    peakRssBytes: process.memoryUsage.rss(),
  };
  if (parsed.isErr()) {
    return {
      kind: "checked",
      failures: [failureFromError(CORPUS_INVARIANTS.parse, parsed.error)],
      producer: UNKNOWN_PRODUCER,
      cost,
      timings: {},
    };
  }

  const wanted = (invariant: CorpusInvariant): boolean => only === undefined || only.has(invariant);
  const fixedPoint =
    wanted(CORPUS_INVARIANTS.fixedPoint) || wanted(CORPUS_INVARIANTS.repackValidates)
      ? await checkFixedPoint(parsed.value)
      : { kind: "failed" as const, failures: [] };
  const validation =
    fixedPoint.kind === "repacked" && wanted(CORPUS_INVARIANTS.repackValidates)
      ? await checkRepackValidates(fixedPoint.repacked)
      : [];
  const styleSet = wanted(CORPUS_INVARIANTS.styleSetRebuild)
    ? await checkStyleSetRebuild(parsed.value)
    : [];
  const extended = await runExtendedChecks({
    bytes,
    buffer,
    parsed: parsed.value,
    documentPart: classification.documentPart,
    invariantBudgetMs,
    fileBudgetMs,
    ...(only === undefined ? {} : { only }),
  });
  return {
    kind: "checked",
    failures: [...fixedPoint.failures, ...validation, ...styleSet, ...extended.failures],
    producer: extended.producer.label,
    cost,
    timings: extended.timings,
  };
};
