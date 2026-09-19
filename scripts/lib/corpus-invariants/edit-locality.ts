/**
 * One tiny edit must stay tiny.
 *
 * A reviewer opens a package, changes a sentence, and saves. Everything the
 * edit did not touch has to come back unchanged: the other paragraphs, the
 * styles the document never renamed, its numbering, its theme, its headers and
 * footers, its images, its fonts and its settings. folio's save path rewrites
 * whole parts from the model whenever a capture no longer matches, so an edit
 * that reaches one paragraph can still cost a part nobody asked it to touch,
 * and nothing in the round-trip invariants would see it: they never edit.
 *
 * This invariant inserts exactly one character into the first non-empty body
 * paragraph through the public headless editing API, saves, and compares.
 * `"direct"` mode is deliberate: tracked changes write `w:ins` markup into the
 * neighbouring runs, so "every other block is unchanged" would be false by
 * construction and prove nothing.
 *
 * The baseline it compares against is a CONTROL SAVE, not the input bytes.
 * Every save re-normalises `w:id` revision ids and `w14:paraId` ranges over
 * every `word/*.xml` part and rewrites `AppVersion` in `docProps/app.xml`,
 * including parts the save did not otherwise touch, because those id spaces are
 * package-wide facts no single serializer can bound. Comparing an edited save
 * against the original bytes would therefore report that normalisation as a
 * locality defect on every file. Saving the same document twice, once with the
 * edit and once without, pays the normalisation on both sides, so what survives
 * is the edit's own reach and nothing else. A no-edit save can still come back
 * byte for byte, when the package needs no normalisation at all; the residue is
 * then whatever the edit forced folio to rewrite, which is the finding.
 *
 * The edited paragraph is located after the fact, by the character it gained:
 * the block ids the editor works in and the body indices the model uses are
 * different id spaces, and inferring one from the other would make a mapping
 * bug look like a locality bug. A paragraph that already carried the character
 * carries it in the control save too, so only the one that gained it is
 * excluded from the comparison.
 */

import { FolioDocxReviewer } from "@stll/folio-core/ai-edits/headless";
import type { FolioAIBlock, FolioAIEditOperation } from "@stll/folio-core/ai-edits/types";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { unzipDocx } from "@stll/folio-core/docx/unzip";
import type { Document } from "@stll/folio-core/types/document";
import { Result } from "better-result";

import { type CorpusFailure, failureFromAssertion, failureFromError } from "../corpus-signature";
import {
  type CorpusInvariantInput,
  type CorpusInvariantOutcome,
  EXTENDED_CORPUS_INVARIANTS,
  type StageTimings,
  timeStage,
} from "./contract";
import { describePackageDifference } from "./model-equality";
import { generalizePartPath } from "./save-idempotence";

const INVARIANT = EXTENDED_CORPUS_INVARIANTS.editLocality;

/**
 * The one character the edit inserts: U+2038 CARET.
 *
 * A character no legal document writes, so a paragraph that carries it after
 * the edit and not before is the paragraph the edit reached.
 */
export const INSERTED_CHARACTER = "‸";

const MESSAGES = {
  missingInsertion: "the inserted character is absent from the edited paragraph",
  changedBlock: "an unedited block changed",
  changedPart: "an unrelated part changed",
} as const;

/**
 * The first paragraph an edit can be aimed at and still be checked.
 *
 * Blank paragraphs have no text to replace. A block inside a table or a
 * content control is addressable, but it is not a direct child of the body, so
 * excluding it from the model comparison would mean walking into the container
 * that holds it; the first top-level paragraph asks the same question without
 * that walk.
 */
export const firstEditableBlock = (blocks: readonly FolioAIBlock[]): FolioAIBlock | undefined =>
  blocks.find(
    (block) =>
      block.text.length > 0 &&
      block.table === undefined &&
      (block.containerPath === undefined || block.containerPath.length === 0),
  );

/**
 * Append the character to the block's whole text.
 *
 * `find` is the entire text so the applier can never call it ambiguous: a
 * string contains itself exactly once.
 */
export const insertOneCharacterOperation = (block: FolioAIBlock): FolioAIEditOperation =>
  ({
    id: "edit-locality-insert",
    type: "replaceInBlock",
    blockId: block.id,
    find: block.text,
    replace: `${block.text}${INSERTED_CHARACTER}`,
  }) as const satisfies FolioAIEditOperation;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Whether the character appears anywhere under a block, text or captured markup alike. */
const carriesInsertedCharacter = (value: unknown, seen: WeakSet<object>): boolean => {
  if (typeof value === "string") {
    return value.includes(INSERTED_CHARACTER);
  }
  if (Array.isArray(value)) {
    return value.some((item) => carriesInsertedCharacter(item, seen));
  }
  if (value instanceof Map) {
    return [...value.values()].some((entry) => carriesInsertedCharacter(entry, seen));
  }
  if (!isRecord(value) || seen.has(value)) {
    return false;
  }
  seen.add(value);
  return Object.values(value).some((entry) => carriesInsertedCharacter(entry, seen));
};

const markedBlockIndices = (document: Document): Set<number> => {
  const marked = new Set<number>();
  for (const [index, block] of document.package.document.content.entries()) {
    if (carriesInsertedCharacter(block, new WeakSet())) {
      marked.add(index);
    }
  }
  return marked;
};

/** No body index, so the comparison keeps every block. */
const NO_EDITED_BLOCK = -1;

/**
 * The body's blocks minus the edited one, as a document the package comparison
 * accepts.
 *
 * `describePackageDifference` walks a whole package, and this invariant owes
 * two separate answers about it: the blocks are compared here and the parts are
 * compared by bytes below. Restricting the comparison to `body.content` keeps a
 * styles or header difference out of the block message, which the part message
 * owns, and still reports one normalised path when a block does change.
 */
const uneditedBlocks = (document: Document, editedIndex: number): Document => ({
  package: {
    document: {
      content: document.package.document.content.filter((_, index) => index !== editedIndex),
    },
  },
});

/** Parts an edit to one body paragraph has no business rewriting. */
const UNRELATED_PART_PATHS: ReadonlySet<string> = new Set([
  "word/styles.xml",
  "word/numbering.xml",
  "word/fonttable.xml",
  "word/settings.xml",
]);

const UNRELATED_PART_PREFIXES = ["word/theme/", "word/header", "word/footer", "word/media/"];

const isUnrelatedPart = (path: string): boolean => {
  const lower = path.toLowerCase();
  return (
    UNRELATED_PART_PATHS.has(lower) ||
    UNRELATED_PART_PREFIXES.some((prefix) => lower.startsWith(prefix))
  );
};

const PART_KINDS = {
  xml: "xml",
  binary: "binary",
} as const;

type PackagePart =
  | { kind: typeof PART_KINDS.xml; text: string }
  | { kind: typeof PART_KINDS.binary; bytes: Uint8Array };

/** The unrelated parts of one saved package, keyed by path. */
const readUnrelatedParts = async (buffer: ArrayBuffer): Promise<Map<string, PackagePart>> => {
  const raw = await unzipDocx(buffer, { extractAllXml: true });
  const entries = Object.entries(raw.originalZip.files).filter(
    ([path, file]) => !file.dir && isUnrelatedPart(path),
  );
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

const sameBytes = (first: Uint8Array, second: Uint8Array): boolean =>
  first.length === second.length && first.every((byte, index) => byte === second[index]);

const partsDiffer = (first: PackagePart, second: PackagePart | undefined): boolean => {
  if (second === undefined) {
    return true;
  }
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

/**
 * Every unrelated part the edit did not leave alone, as generalised paths.
 *
 * A part that only the edited save holds is a difference too, and generalising
 * collapses `word/header2.xml` and `word/header7.xml` into the one defect they
 * are, so the set needs no separate cap.
 */
const changedPartPaths = (
  control: Map<string, PackagePart>,
  edited: Map<string, PackagePart>,
): string[] => {
  const changed = new Set<string>();
  for (const [path, part] of control) {
    if (partsDiffer(part, edited.get(path))) {
      changed.add(generalizePartPath(path));
    }
  }
  for (const path of edited.keys()) {
    if (!control.has(path)) {
      changed.add(generalizePartPath(path));
    }
  }
  // Archive order is the package's, not the defect's: sorting makes one file's
  // report the same on every run.
  return [...changed].toSorted((left, right) => (left < right ? -1 : 1));
};

const parse = (buffer: ArrayBuffer): Promise<Document> =>
  parseDocx(buffer, { preloadFonts: false });

export const runEditLocalityInvariant = async ({
  buffer,
}: CorpusInvariantInput): Promise<CorpusInvariantOutcome> => {
  const timings: StageTimings = {};

  const opened = await timeStage(timings, "open", () =>
    Result.tryPromise({
      try: () => FolioDocxReviewer.fromBuffer(buffer),
      catch: (cause: unknown) => cause,
    }),
  );
  if (opened.isErr()) {
    return { failures: [failureFromError(INVARIANT, opened.error)], timings };
  }
  const reviewer = opened.value;

  const selected = await timeStage(timings, "snapshot", () =>
    Result.try(() => firstEditableBlock(reviewer.snapshot().blocks)),
  );
  if (selected.isErr()) {
    return { failures: [failureFromError(INVARIANT, selected.error)], timings };
  }
  const target = selected.value;
  // A package with nothing to edit has nothing to keep local.
  if (target === undefined) {
    return { failures: [], timings };
  }

  const application = await timeStage(timings, "apply", () =>
    Result.try(() =>
      reviewer.applyOperations([insertOneCharacterOperation(target)], { mode: "direct" }),
    ),
  );
  if (application.isErr()) {
    return { failures: [failureFromError(INVARIANT, application.error)], timings };
  }
  if (application.value.applied.length === 0) {
    return {
      failures: [failureFromAssertion(INVARIANT, MESSAGES.missingInsertion)],
      timings,
    };
  }

  const editedSave = await timeStage(timings, "save", () =>
    Result.tryPromise({ try: () => reviewer.toBuffer(), catch: (cause: unknown) => cause }),
  );
  if (editedSave.isErr()) {
    return { failures: [failureFromError(INVARIANT, editedSave.error)], timings };
  }

  const controlSave = await timeStage(timings, "control-save", () =>
    Result.tryPromise({
      try: async () => (await FolioDocxReviewer.fromBuffer(buffer)).toBuffer(),
      catch: (cause: unknown) => cause,
    }),
  );
  if (controlSave.isErr()) {
    return { failures: [failureFromError(INVARIANT, controlSave.error)], timings };
  }

  const read = await timeStage(timings, "parse", () =>
    Result.tryPromise({
      try: () =>
        Promise.all([
          parse(editedSave.value),
          parse(controlSave.value),
          readUnrelatedParts(editedSave.value),
          readUnrelatedParts(controlSave.value),
        ]),
      catch: (cause: unknown) => cause,
    }),
  );
  if (read.isErr()) {
    return { failures: [failureFromError(INVARIANT, read.error)], timings };
  }
  const [edited, control, editedParts, controlParts] = read.value;

  const compared = await timeStage(timings, "compare", () => {
    const failures: CorpusFailure[] = [];

    const marked = markedBlockIndices(control);
    const gained = [...markedBlockIndices(edited)].filter((index) => !marked.has(index));
    const editedIndex = gained.at(0) ?? NO_EDITED_BLOCK;
    if (editedIndex === NO_EDITED_BLOCK) {
      failures.push(failureFromAssertion(INVARIANT, MESSAGES.missingInsertion));
    }

    const difference = describePackageDifference(
      uneditedBlocks(control, editedIndex),
      uneditedBlocks(edited, editedIndex),
    );
    if (difference !== null) {
      failures.push(failureFromAssertion(INVARIANT, `${MESSAGES.changedBlock}: ${difference}`));
    }

    for (const path of changedPartPaths(controlParts, editedParts)) {
      failures.push(failureFromAssertion(INVARIANT, `${MESSAGES.changedPart}: ${path}`));
    }

    return failures;
  });

  return { failures: compared, timings };
};
