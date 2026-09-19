/**
 * Open a document in the editor's model, save it, and demand the whole package
 * back.
 *
 * The gate's `fixed-point` invariant runs this exact pipeline already, but it
 * compares only the visible text and the text-block count: a round trip that
 * drops a table's cell margins, a run's language tag or a section's page
 * borders passes it unchanged. This one compares the normalised package, so a
 * field the ProseMirror conversion cannot carry is a failure rather than a
 * silent loss.
 *
 * The comparison is against the input model, not against the ProseMirror
 * document, because the loss can happen on either leg: `toProseDoc` may never
 * represent a property, and `fromProseDoc` may fail to write back one it did.
 */

import { parseDocx } from "@stll/folio-core/docx/parser";
import { repackDocx } from "@stll/folio-core/docx/rezip";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "@stll/folio-core/prosemirror/conversion/toProseDoc";
import { Result } from "better-result";

import { failureFromAssertion, failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";
import { describePackageDifference } from "./model-equality";

const DIFFERENCE_PREFIX = "editor round trip changed";

export const runEditorRoundTripInvariant = async ({
  parsed,
}: CorpusInvariantInput): Promise<CorpusInvariantOutcome> => {
  const timings: Record<string, number> = {};

  const proseDoc = await timeStage(timings, "to-prose", () =>
    Promise.resolve(Result.try(() => toProseDoc(parsed))),
  );
  if (proseDoc.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.editorRoundTrip, proseDoc.error)],
      timings,
    };
  }

  const back = await timeStage(timings, "from-prose", () =>
    Promise.resolve(Result.try(() => fromProseDoc(proseDoc.value, parsed))),
  );
  if (back.isErr()) {
    return {
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.editorRoundTrip, back.error)],
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
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.editorRoundTrip, saved.error)],
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
      failures: [failureFromError(EXTENDED_CORPUS_INVARIANTS.editorRoundTrip, reparsed.error)],
      timings,
    };
  }

  const difference = await timeStage(timings, "compare", () =>
    Promise.resolve(describePackageDifference(parsed, reparsed.value)),
  );
  if (difference === null) {
    return { failures: [], timings };
  }
  return {
    failures: [
      failureFromAssertion(
        EXTENDED_CORPUS_INVARIANTS.editorRoundTrip,
        `${DIFFERENCE_PREFIX} ${difference}`,
      ),
    ],
    timings,
  };
};
