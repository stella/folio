/**
 * Measure the ProseMirror projection itself, with reuse declined.
 *
 * `editor-round-trip` runs the same pipeline and asks a different question.
 * Once `fromProseDoc` may return a record the editor did not change by
 * reference from the base document, a round trip over an untouched package
 * stops exercising the projection at all: every block comes back `===` its
 * base, and what is left is a plain repack. The number would collapse and read
 * as a fix, while a field `toProseDoc` cannot carry would still be lost for
 * every edited paragraph.
 *
 * So this leg declines reuse, exactly as `reserialize` strips the capture slots
 * so the serializers must run. The explicit option keeps the instrument stable
 * if ordinary editor round trips later begin reusing matched base records.
 */

import { parseDocx } from "@stll/folio-core/docx/parser";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import type { Document } from "@stll/folio-core/types/document";
import { Result } from "better-result";

import { failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";
import { describePackageDifferences, differenceFailures } from "./model-equality";

const DIFFERENCE_PREFIX = "editor projection changed";

/**
 * The forcing, named so the gate and its test read the same call.
 *
 * The analogue of `withoutSerializerCaptures`: that one removes what replay
 * would otherwise answer from, this one removes what the merge would otherwise
 * answer from.
 */
export const projectedWithoutReuse = (
  pmDoc: ReturnType<typeof toProseDoc>,
  baseDocument: Document,
): Document => fromProseDoc(pmDoc, baseDocument, { reuse: "none" });

export const runEditorProjectionInvariant = async ({
  parsed,
}: CorpusInvariantInput): Promise<CorpusInvariantOutcome> => {
  const timings: Record<string, number> = {};

  const proseDoc = await timeStage(timings, "to-prose", () =>
    Promise.resolve(Result.try(() => toProseDoc(parsed))),
  );
  if (proseDoc.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.editorProjection, proseDoc.error)],
      timings,
    };
  }

  const back = await timeStage(timings, "from-prose", () =>
    Promise.resolve(Result.try(() => projectedWithoutReuse(proseDoc.value, parsed))),
  );
  if (back.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.editorProjection, back.error)],
      timings,
    };
  }

  const saved = await timeStage(timings, "save", () =>
    Result.tryPromise({
      try: () => repackDocx(back.value, { updateModifiedDate: false }),
      catch: (cause: unknown) => cause,
    }),
  );
  if (saved.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.editorProjection, saved.error)],
      timings,
    };
  }

  const reparsed = await timeStage(timings, "parse", () =>
    Result.tryPromise({
      try: () => parseDocx(saved.value, { preloadFonts: false }),
      catch: (cause: unknown) => cause,
    }),
  );
  if (reparsed.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.editorProjection, reparsed.error)],
      timings,
    };
  }

  const differences = await timeStage(timings, "compare", () =>
    Promise.resolve(describePackageDifferences(parsed, reparsed.value)),
  );
  return {
    failures: differenceFailures(
      EXTENDED_CORPUS_INVARIANTS.editorProjection,
      differences,
      (message) => `${DIFFERENCE_PREFIX} ${message}`,
    ),
    timings,
  };
};
