import { isDeepStrictEqual } from "node:util";
import { applyDocumentOps } from "../../../packages/docx-core/src/ops/documentOps";
import { Result } from "better-result";
import { failureFromAssertion } from "../corpus-signature";
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
  type OpSequenceStep,
} from "./op-sequences";
import { firstDifferingOpPart } from "./op-part-difference";

import { opErrorOutcome } from "./op-outcome";

const INVARIANT = EXTENDED_CORPUS_INVARIANTS.opInverse;

type SerializedInverseStepOptions = {
  step: OpSequenceStep;
  control: ReadonlyMap<string, Uint8Array>;
  restored: ReadonlyMap<string, Uint8Array>;
};

/** The same package-byte law, classified by the operation whose inverse ran. */
export const serializedInverseStepFailures = ({
  step,
  control,
  restored,
}: SerializedInverseStepOptions): string[] => {
  const part = firstDifferingOpPart({ control, edited: restored });
  return part === undefined
    ? []
    : [`${step.op.type} inverse changed the original serialized package parts: ${part}`];
};

type SerializedInverseSequenceOptions = {
  sequence: OpSequence;
  control: ReadonlyMap<string, Uint8Array>;
  restored: ReadonlyMap<string, Uint8Array>;
  serializeParts?: (document: OpSequence["document"]) => Promise<Map<string, Uint8Array>>;
};

/** Diagnose only failing compound saves, stopping before one defect cascades into later undos. */
export const serializedInverseSequenceFailures = async ({
  sequence,
  control,
  restored,
  serializeParts = serializedOpParts,
}: SerializedInverseSequenceOptions): Promise<string[]> => {
  const part = firstDifferingOpPart({ control, edited: restored });
  if (part === undefined) return [];
  let current = sequence.document;
  for (const step of sequence.steps.toReversed()) {
    const undone = applyDocumentOps(current, step.edit.inverse);
    if (undone.isErr())
      return [`${step.op.type} compound inverse refused: ${undone.error.reason}; part: ${part}`];
    // oxlint-disable-next-line no-await-in-loop -- diagnose each reverse-composed undo against its own original state
    const [beforeParts, restoredParts] = await Promise.all([
      serializeParts(step.before),
      serializeParts(undone.value.document),
    ]);
    const failures = serializedInverseStepFailures({
      step,
      control: beforeParts,
      restored: restoredParts,
    });
    if (failures.length > 0) return failures;
    current = undone.value.document;
  }
  // Each grouped undo passed: retain composition context without claiming an individual culprit.
  const kinds = [...new Set(sequence.steps.map(({ op }) => op.type))];
  if (kinds.length === 0)
    return [`empty sequence composition changed the original serialized package parts: ${part}`];
  return kinds.map(
    (type) => `${type} sequence composition changed the original serialized package parts: ${part}`,
  );
};

/** Check individual inverses and the reverse composition, without erasing captures. */
export const inverseStepFailures = ({ before, op, edit }: OpSequenceStep): string[] => {
  const restored = applyDocumentOps(edit.document, edit.inverse);
  if (restored.isErr()) return [`${op.type} inverse refused: ${restored.error.reason}`];
  const failures: string[] = [];
  if (!sameOpModel(before, restored.value.document))
    failures.push(`${op.type} inverse changed the original records`);
  if (serializeOpDocument(before) !== serializeOpDocument(restored.value.document))
    failures.push(`${op.type} inverse changed the original serialized blocks`);
  return failures;
};

export const inverseSequenceFailures = (sequence: OpSequence): string[] => {
  const failures = sequence.mutations.map((type) => `${type} mutated its input document`);
  for (const step of sequence.steps) failures.push(...inverseStepFailures(step));
  const restored = applyDocumentOps(sequence.document, sequence.inverse);
  failures.push(...inverseRestorationFailures(sequence, restored));
  return [...new Set(failures)];
};

const inverseRestorationFailures = (
  sequence: OpSequence,
  restored: ReturnType<typeof applyDocumentOps>,
): string[] => {
  const failures: string[] = [];
  if (restored.isErr()) {
    failures.push(`sequence inverse refused: ${restored.error.reason}`);
    return failures;
  }
  if (!isDeepStrictEqual(sequence.originalModel, exactOpModel(restored.value.document)))
    failures.push("sequence inverse changed the original records");
  if (sequence.originalXml !== serializeOpDocument(restored.value.document))
    failures.push("sequence inverse changed the original serialized blocks");
  return failures;
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
          failures.push(
            ...sequence.mutations.map(
              (type) => `${type} mutated its input document; part: model-only`,
            ),
          );
          for (const step of sequence.steps) {
            const stepFailures = inverseStepFailures(step);
            if (stepFailures.length === 0) continue;
            const undone = applyDocumentOps(step.edit.document, step.edit.inverse);
            if (undone.isErr()) {
              failures.push(...stepFailures.map((message) => `${message}; part: not-serialized`));
              continue;
            }
            // oxlint-disable-next-line no-await-in-loop -- save only a failed individual inverse to classify its first differing part
            const [beforeParts, undoneParts] = await Promise.all([
              serializedOpParts(step.before),
              serializedOpParts(undone.value.document),
            ]);
            const stepPart =
              firstDifferingOpPart({ control: beforeParts, edited: undoneParts }) ?? "model-only";
            failures.push(...stepFailures.map((message) => `${message}; part: ${stepPart}`));
          }
          const restored = applyDocumentOps(sequence.document, sequence.inverse);
          const modelFailures = inverseRestorationFailures(sequence, restored);
          if (restored.isErr()) {
            for (const message of modelFailures) {
              for (const type of new Set(sequence.steps.map(({ op }) => op.type)))
                failures.push(`${type} composition context: ${message}; part: not-serialized`);
            }
            continue;
          }
          // oxlint-disable-next-line no-await-in-loop -- compare each independently restored sequence with the shared control
          const restoredParts = await serializedOpParts(restored.value.document);
          // oxlint-disable-next-line no-await-in-loop -- only a failed compound package save triggers per-operation diagnosis
          const packageFailures = await serializedInverseSequenceFailures({
            sequence,
            control,
            restored: restoredParts,
          });
          failures.push(...packageFailures);
          const part = firstDifferingOpPart({ control, edited: restoredParts }) ?? "model-only";
          for (const message of modelFailures) {
            // A classified reverse undo replaces an unclassified compound assertion.
            if (packageFailures.length > 0) continue;
            for (const type of new Set(sequence.steps.map(({ op }) => op.type)))
              failures.push(`${type} composition context: ${message}; part: ${part}`);
          }
        }
        return [...new Set(failures)];
      },
      catch: (cause: unknown) => cause,
    }),
  );
  if (outcome.isErr())
    return opErrorOutcome({ invariant: INVARIANT, error: outcome.error, timings });
  return {
    status: "evaluated",
    timings,
    failures: outcome.value.map((message) => failureFromAssertion(INVARIANT, message)),
  };
};
