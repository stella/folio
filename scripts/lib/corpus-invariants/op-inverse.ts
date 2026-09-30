import { isDeepStrictEqual } from "node:util";
import { applyDocumentOps } from "../../../packages/docx-core/src/ops/documentOps";
import { Result } from "better-result";
import { failureFromAssertion, failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";
import {
  exactOpModel,
  generateOpSequence,
  OP_SEQUENCE_SEEDS,
  prepareOpDocument,
  sameOpModel,
  seedFromBytes,
  serializeOpDocument,
  serializedOpParts,
  type OpSequence,
} from "./op-sequences";

const INVARIANT = EXTENDED_CORPUS_INVARIANTS.opInverse;

/** Check individual inverses and the reverse composition, without erasing captures. */
export const inverseSequenceFailures = (sequence: OpSequence): string[] => {
  const failures = sequence.mutations.map((type) => `${type} mutated its input document`);
  for (const { before, op, edit } of sequence.steps) {
    const restored = applyDocumentOps(edit.document, edit.inverse);
    if (restored.isErr()) {
      failures.push(`${op.type} inverse refused: ${restored.error.reason}`);
      continue;
    }
    if (!sameOpModel(before, restored.value.document))
      failures.push(`${op.type} inverse changed the original records`);
    if (serializeOpDocument(before) !== serializeOpDocument(restored.value.document))
      failures.push(`${op.type} inverse changed the original serialized blocks`);
  }
  const restored = applyDocumentOps(sequence.document, sequence.inverse);
  if (restored.isErr()) {
    failures.push(`sequence inverse refused: ${restored.error.reason}`);
    return failures;
  }
  if (!isDeepStrictEqual(sequence.originalModel, exactOpModel(restored.value.document)))
    failures.push("sequence inverse changed the original records");
  if (sequence.originalXml !== serializeOpDocument(restored.value.document))
    failures.push("sequence inverse changed the original serialized blocks");
  return [...new Set(failures)];
};

export const runOpInverseInvariant = async (
  input: CorpusInvariantInput,
): Promise<CorpusInvariantOutcome> => {
  const timings = {};
  const outcome = await timeStage(timings, "sequences", () =>
    Result.tryPromise({
      try: async () => {
        const document = await prepareOpDocument(input);
        const seed = seedFromBytes(input.bytes);
        const failures: string[] = [];
        const control = await serializedOpParts(document);
        for (const salt of OP_SEQUENCE_SEEDS) {
          const sequence = generateOpSequence(document, seed ^ salt);
          failures.push(...inverseSequenceFailures(sequence));
          const restored = applyDocumentOps(sequence.document, sequence.inverse);
          if (restored.isErr()) continue;
          // oxlint-disable-next-line no-await-in-loop -- compare each independently restored sequence with the shared control
          const restoredParts = await serializedOpParts(restored.value.document);
          if (!sameOpModel(control, restoredParts))
            failures.push("sequence inverse changed the original serialized package parts");
        }
        return [...new Set(failures)];
      },
      catch: (cause: unknown) => cause,
    }),
  );
  return {
    timings,
    failures: outcome.isErr()
      ? [failureFromError(INVARIANT, outcome.error)]
      : outcome.value.map((message) => failureFromAssertion(INVARIANT, message)),
  };
};
