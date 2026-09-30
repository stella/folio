import type { BlockContent, Document } from "../../../packages/docx-core/src/model/document";
import { Result } from "better-result";
import { storyParagraphs } from "../../../packages/docx-core/src/ops/blocks";
import { idKey } from "../../../packages/docx-core/src/ops/ids";
import { failureFromAssertion, failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  timeStage,
} from "./contract";
import {
  generateOpSequence,
  OP_SEQUENCE_SEEDS,
  prepareOpDocument,
  sameOpModel,
  seedFromBytes,
  serializeOpDocument,
  serializedOpParts,
  type OpSequenceStep,
} from "./op-sequences";
import { generalizePartPath } from "./save-idempotence";

const INVARIANT = EXTENDED_CORPUS_INVARIANTS.opLocality;

/** Remove declared paragraphs and their empty ancestors, preserving every other field. */
const untouchedBlocks = (
  blocks: readonly BlockContent[],
  touched: ReadonlySet<string>,
): BlockContent[] => {
  const out: BlockContent[] = [];
  for (const block of blocks) {
    switch (block.type) {
      case "paragraph":
        if (block.paraId === undefined || !touched.has(idKey(block.paraId))) out.push(block);
        break;
      case "table": {
        const rows = block.rows.flatMap((row) => {
          const cells = row.cells.flatMap((cell) => {
            const content = untouchedBlocks(cell.content, touched);
            return content.length === 0 ? [] : [{ ...cell, content }];
          });
          return cells.length === 0 ? [] : [{ ...row, cells }];
        });
        if (rows.length > 0) out.push({ ...block, rows });
        break;
      }
      case "blockSdt":
      case "blockCustomXml": {
        const content = untouchedBlocks(block.content, touched);
        if (content.length > 0) out.push({ ...block, content });
        break;
      }
      case "preservedBlock":
      case "bookmarkStart":
      case "bookmarkEnd":
        out.push(block);
        break;
      default: {
        const unreachable: never = block;
        return unreachable;
      }
    }
  }
  return out;
};

const unrelatedModel = (document: Document): unknown => ({
  ...document,
  package: {
    ...document.package,
    document: { ...document.package.document, content: undefined, sections: undefined },
  },
});

const projected = (document: Document, touched: ReadonlySet<string>): Document => {
  const body = {
    ...document.package.document,
    content: untouchedBlocks(document.package.document.content, touched),
  };
  delete body.sections;
  return { ...document, package: { ...document.package, document: body } };
};

/** Check the producer's declared touched set, never infer it from observed differences. */
export const localityStepFailures = ({ before, op, edit }: OpSequenceStep): string[] => {
  const failures: string[] = [];
  const modified = new Set(edit.touched.modified.map(idKey));
  const inserted = new Set(edit.touched.inserted.map(idKey));
  const removed = new Set(edit.touched.removed.map(idKey));
  const touched = new Set([...modified, ...inserted, ...removed]);
  const beforeParagraphs = new Map(
    storyParagraphs(before.package.document).flatMap(({ paragraph }) =>
      paragraph.paraId === undefined ? [] : [[idKey(paragraph.paraId), paragraph] as const],
    ),
  );
  const afterParagraphs = new Map(
    storyParagraphs(edit.document.package.document).flatMap(({ paragraph }) =>
      paragraph.paraId === undefined ? [] : [[idKey(paragraph.paraId), paragraph] as const],
    ),
  );
  for (const [id, paragraph] of beforeParagraphs) {
    const after = afterParagraphs.get(id);
    if (after === undefined) {
      if (!removed.has(id)) failures.push(`${op.type} removed an undeclared block`);
    } else if (!touched.has(id) && !sameOpModel(paragraph, after)) {
      failures.push(`${op.type} changed an untouched paragraph`);
    }
  }
  for (const id of afterParagraphs.keys()) {
    if (!beforeParagraphs.has(id) && !inserted.has(id))
      failures.push(`${op.type} inserted an undeclared block`);
  }
  if (!sameOpModel(unrelatedModel(before), unrelatedModel(edit.document)))
    failures.push(`${op.type} changed records outside the main-story blocks`);
  const originalUntouched = projected(before, touched);
  const editedUntouched = projected(edit.document, touched);
  if (
    !sameOpModel(
      originalUntouched.package.document.content,
      editedUntouched.package.document.content,
    )
  )
    failures.push(`${op.type} changed an untouched block or container`);
  if (serializeOpDocument(originalUntouched) !== serializeOpDocument(editedUntouched))
    failures.push(`${op.type} changed serialization outside its touched blocks`);
  return [...new Set(failures)];
};

export const runOpLocalityInvariant = async (
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
          failures.push(...sequence.mutations.map((type) => `${type} mutated its input document`));
          for (const step of sequence.steps) failures.push(...localityStepFailures(step));
          // oxlint-disable-next-line no-await-in-loop -- each sequence is checked against one shared control save
          const edited = await serializedOpParts(sequence.document);
          for (const path of new Set([...control.keys(), ...edited.keys()])) {
            if (path === input.documentPart) continue;
            if (!sameOpModel(control.get(path), edited.get(path)))
              failures.push(
                `sequence changed unrelated serialized part: ${generalizePartPath(path)}`,
              );
          }
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
