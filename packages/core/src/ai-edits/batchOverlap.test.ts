/**
 * Operations of one batch all address the document as it was read. The
 * applier runs them from the end of the document backwards, which keeps
 * disjoint targets valid; two operations on the same target are not disjoint.
 * The batch refuses the later of two operations whose targets conflict
 * (`overlappingOperation`) and applies the rest exactly as they would apply
 * one at a time — so no operation ever edits a block it did not name.
 *
 * The oracle applies the operations the batch applied, each in a batch of its
 * own and re-resolved against the document as it then stands: insertions
 * first, then comments and formatting (which move no position), then the rest
 * from the end of the document backwards, as the batch runs them.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { EditorState } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";

import {
  blockTexts,
  OperationSession,
  openReviewer,
  paragraphsDocx,
  reopened,
  reopenedAccepted,
  textRun,
} from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationResult,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown";
import type { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIEditSnapshot } from "./types";

setDefaultTimeout(propertyTestTimeout(60_000));

type Mode = "direct" | "tracked-changes";
const MODES: readonly Mode[] = ["direct", "tracked-changes"];

const apply = (
  reviewer: FolioDocxReviewer,
  mode: Mode,
  operations: FolioDocumentOperation[],
): FolioDocumentOperationResult =>
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations,
  });

describe("an operation inside a block another operation of the batch deletes", () => {
  const threeClauses = () =>
    paragraphsDocx([
      [textRun("Alpha beta gamma.")],
      [textRun("Second clause.")],
      [textRun("Third clause.")],
    ]);

  test.each(MODES)(
    "is refused and never edits the next paragraph (%s, saved and reopened)",
    async (mode) => {
      const reviewer = await openReviewer(await threeClauses());
      const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
      const result = apply(reviewer, mode, [
        { id: "replace", type: "replaceInBlock", blockId, find: "Alpha", replace: "First" },
        { id: "delete", type: "deleteBlock", blockId },
      ]);

      expect(result.applied.map(({ id }) => id)).toEqual(["replace"]);
      const message = 'operation "replace", earlier in this batch, already claims its target.';
      expect(result.skipped).toEqual([{ id: "delete", reason: "overlappingOperation", message }]);
      expect(result.issues).toEqual([
        {
          operationId: "delete",
          operationIndex: 1,
          path: "$.operations[1]",
          code: "overlappingOperation",
          retryable: true,
          recovery: "refreshDocument",
          message,
        },
      ]);
      const final = mode === "direct" ? await reopened(reviewer) : await reopenedAccepted(reviewer);
      expect(blockTexts(final)).toEqual(["First beta gamma.", "Second clause.", "Third clause."]);
    },
  );

  test.each(MODES)("refuses the later one whichever comes first (%s)", async (mode) => {
    const reviewer = await openReviewer(await threeClauses());
    const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
    const result = apply(reviewer, mode, [
      { id: "delete", type: "deleteBlock", blockId },
      { id: "replace", type: "replaceInBlock", blockId, find: "Alpha", replace: "First" },
    ]);

    expect(result.skipped).toEqual([
      {
        id: "replace",
        reason: "overlappingOperation",
        message: 'operation "delete", earlier in this batch, already claims its target.',
      },
    ]);
    const final = mode === "direct" ? await reopened(reviewer) : await reopenedAccepted(reviewer);
    expect(blockTexts(final)).toEqual(["Second clause.", "Third clause."]);
  });
});

describe("a row inserted in a batch that changes blocks around its table", () => {
  test.each(MODES)("lands where it was resolved (%s, saved and reopened)", async (mode) => {
    const source = [
      "Lead paragraph.",
      "",
      "| A1 | B1 |",
      "| --- | --- |",
      "| A2 | B2 |",
      "| A3 | B3 |",
      "",
      "Tail paragraph.",
    ].join("\n");
    const reviewer = await openReviewer(
      (await ensureParaIds(new Uint8Array(await createDocx(fromMarkdown(source))))).docx.slice()
        .buffer,
    );
    const idOf = (text: string) =>
      reviewer.snapshot().blocks.find((block) => block.text === text)?.id ?? "";
    const result = apply(reviewer, mode, [
      { id: "row", type: "insertTableRow", blockId: idOf("A2"), cellTexts: ["X", "Y"] },
      {
        id: "lead",
        type: "insertAfterBlock",
        blockId: idOf("Lead paragraph."),
        text: "Inserted before the table.",
      },
      { id: "delete", type: "deleteBlock", blockId: idOf("Lead paragraph.") },
      { id: "cell", type: "replaceInBlock", blockId: idOf("A1"), find: "A1", replace: "A one" },
      {
        id: "tail",
        type: "splitBlock",
        blockId: idOf("Tail paragraph."),
        offset: 4,
        separator: " ",
      },
    ]);
    expect(result.skipped).toEqual([]);
    const final = mode === "direct" ? await reopened(reviewer) : await reopenedAccepted(reviewer);
    expect(blockTexts(final)).toEqual([
      "Inserted before the table.",
      "A one",
      "B1",
      "A2",
      "B2",
      "X",
      "Y",
      "A3",
      "B3",
      "Tail",
      "paragraph.",
    ]);
  });
});

// ---------------------------------------------------------------------------
// The class: every pair and random batches against the one-at-a-time oracle.
// ---------------------------------------------------------------------------

const BLOCK_COUNT = 4;
const TOKEN_COUNT = 4;

const token = (block: number, index: number): string => `t${String(block)}${String(index)}x`;
const blockText = (block: number): string =>
  Array.from({ length: TOKEN_COUNT }, (_, index) => token(block, index)).join(" ");

type Span = { block: number; first: number; last: number };

type GeneratedOperation =
  | ({ kind: "replaceInBlock" | "replaceRange" | "formatRange" | "commentOnRange" } & Span)
  | { kind: "splitBlock"; block: number; before: number }
  | {
      kind:
        | "commentOnBlock"
        | "mergeBlockWithNext"
        | "deleteBlock"
        | "replaceBlock"
        | "setBlockParagraphProperties"
        | "insertAfterBlock"
        | "insertBeforeBlock";
      block: number;
    };

const INSERTIONS: ReadonlySet<GeneratedOperation["kind"]> = new Set([
  "insertAfterBlock",
  "insertBeforeBlock",
]);
const ANNOTATIONS: ReadonlySet<GeneratedOperation["kind"]> = new Set([
  "formatRange",
  "commentOnRange",
  "commentOnBlock",
]);

const spanText = ({ block, first, last }: Span): string =>
  Array.from({ length: last - first + 1 }, (_, offset) => token(block, first + offset)).join(" ");

/**
 * Where the batch places an operation, in the coordinates it resolved against:
 * `[block, 0, 0]` is the position before the block, `[block, 1, offset]` a
 * position in its text. The position after a block is the one before the next.
 */
const placement = (operation: GeneratedOperation): [number, number, number] => {
  const text = blockText(operation.block);
  switch (operation.kind) {
    case "replaceInBlock":
    case "replaceRange":
    case "formatRange":
    case "commentOnRange":
      return [operation.block, 1, text.indexOf(token(operation.block, operation.first))];
    case "splitBlock":
      return [operation.block, 1, text.indexOf(token(operation.block, operation.before)) - 1];
    case "mergeBlockWithNext":
    case "insertAfterBlock":
      return [operation.block + 1, 0, 0];
    case "setBlockParagraphProperties":
    case "insertBeforeBlock":
      return [operation.block, 0, 0];
    default:
      return [operation.block, 1, 0];
  }
};

const comparePlacement = (
  left: readonly [number, number, number],
  right: readonly [number, number, number],
): number => left[0] - right[0] || left[1] - right[1] || left[2] - right[2];

/**
 * The operation `generated` stands for, resolved against the reviewer's
 * document as it stands: token text is looked up where it is now. `null`
 * when what it names is no longer there.
 */
const materialize = (
  reviewer: { snapshot(): FolioAIEditSnapshot },
  blockIds: readonly string[],
  generated: GeneratedOperation,
  index: number,
): FolioDocumentOperation | null => {
  const id = `op${String(index)}`;
  const blockId = blockIds[generated.block] ?? "";
  const current = reviewer.snapshot().blocks.find((block) => block.id === blockId);
  if (!current) {
    return null;
  }
  const rangeOf = (span: Span) => {
    const text = spanText(span);
    const start = current.text.indexOf(text);
    return start < 0
      ? null
      : createFolioAITextRangeHandle({
          blockId,
          text: current.text,
          startOffset: start,
          endOffset: start + text.length,
        });
  };
  switch (generated.kind) {
    case "replaceInBlock": {
      const find = spanText(generated);
      return current.text.includes(find)
        ? { id, type: "replaceInBlock", blockId, find, replace: `new${String(index)}` }
        : null;
    }
    case "replaceRange": {
      const range = rangeOf(generated);
      return range && { id, type: "replaceRange", range, replace: `rng${String(index)}` };
    }
    case "formatRange": {
      const range = rangeOf(generated);
      return range && { id, type: "formatRange", range, formatting: { bold: true } };
    }
    case "commentOnRange": {
      const range = rangeOf(generated);
      return range && { id, type: "commentOnRange", range, comment: { text: `c${String(index)}` } };
    }
    case "commentOnBlock":
      return { id, type: "commentOnBlock", blockId, comment: { text: `c${String(index)}` } };
    case "splitBlock": {
      // After the token before the break, which an edit of the text after
      // the break leaves where it is.
      const previous = token(generated.block, generated.before - 1);
      const at = current.text.indexOf(previous);
      return at < 0
        ? null
        : { id, type: "splitBlock", blockId, offset: at + previous.length, separator: " " };
    }
    case "mergeBlockWithNext":
    case "deleteBlock":
      return { id, type: generated.kind, blockId };
    case "replaceBlock":
      return { id, type: "replaceBlock", blockId, text: `whole${String(index)}` };
    case "setBlockParagraphProperties":
      return {
        id,
        type: "setBlockParagraphProperties",
        blockId,
        properties: { alignment: "center" },
      };
    case "insertAfterBlock":
    case "insertBeforeBlock":
      return { id, type: generated.kind, blockId, text: `ins${String(index)}` };
  }
};

/** Reasons a generated batch may skip an operation for, besides a conflict. */
const INCIDENTAL_SKIPS: ReadonlySet<string> = new Set([
  "noopOperation",
  "pendingParagraphPropertyChange",
  "pendingRunPropertyChange",
  "unsupportedBlock",
]);

let tokenDocument: Promise<EditorState> | undefined;
/** The token paragraphs, parsed once; every batch starts from the same state. */
const freshSession = async (): Promise<OperationSession> => {
  tokenDocument ??= paragraphsDocx(
    Array.from({ length: BLOCK_COUNT }, (_, block) => [textRun(blockText(block))]),
  )
    .then(openReviewer)
    .then((reviewer) => reviewer.state);
  return new OperationSession(await tokenDocument);
};

/**
 * Apply `generated` as one batch, then the applied ones one at a time; both
 * must leave the same document, and a block no operation names must keep its
 * text. Returns what went wrong, if anything, naming the batch.
 */
const batchAgainstOneAtATime = async (
  generated: readonly GeneratedOperation[],
  mode: Mode,
): Promise<string[]> => {
  const problems: string[] = [];
  const report = (problem: string) => {
    problems.push(`${mode} ${JSON.stringify(generated)}: ${problem}`);
  };
  const batch = await freshSession();
  const blockIds = batch.snapshot().blocks.map((block) => block.id);
  const operations = generated.flatMap((operation, index) => {
    const materialized = materialize(batch, blockIds, operation, index);
    return materialized === null ? [] : [materialized];
  });
  if (operations.length !== generated.length) {
    report("a generated operation did not resolve against the document it was made for");
    return problems;
  }
  const result = batch.apply(mode, operations);
  for (const { id, reason } of result.skipped) {
    if (reason !== "overlappingOperation" && !INCIDENTAL_SKIPS.has(reason)) {
      report(`${id} skipped as ${reason}`);
    }
  }
  const appliedIds = new Set(result.applied.map(({ id }) => id));
  const applied = generated.flatMap((operation, index) =>
    appliedIds.has(`op${String(index)}`) ? [{ operation, index }] : [],
  );

  const oracle = await freshSession();
  const materializeAll = (entries: readonly { operation: GeneratedOperation; index: number }[]) =>
    entries.flatMap(({ operation, index }) => {
      const materialized = materialize(oracle, blockIds, operation, index);
      if (materialized === null) {
        report(`op${String(index)} no longer resolves one at a time`);
        return [];
      }
      return [materialized];
    });
  const applyAlone = (operationsToApply: FolioDocumentOperation[]) => {
    if (operationsToApply.length === 0) {
      return;
    }
    // An operation the ones before it already carried out has nothing left
    // to do; the documents are compared below either way.
    for (const { id, reason } of oracle.apply(mode, operationsToApply).skipped) {
      if (reason !== "noopOperation") {
        report(`${id} skipped one at a time as ${reason}`);
      }
    }
  };
  // Insertions sharing a gap keep their input order: that order is the
  // batch's own contract, so the oracle states them together the same way.
  applyAlone(materializeAll(applied.filter(({ operation }) => INSERTIONS.has(operation.kind))));
  const rest = [
    ...applied.filter(({ operation }) => ANNOTATIONS.has(operation.kind)),
    ...applied
      .filter(
        ({ operation }) => !INSERTIONS.has(operation.kind) && !ANNOTATIONS.has(operation.kind),
      )
      .toSorted(
        (left, right) =>
          comparePlacement(placement(right.operation), placement(left.operation)) ||
          right.index - left.index,
      ),
  ];
  for (const entry of rest) {
    applyAlone(materializeAll([entry]));
  }

  // A comment's reference anchor is placed after its range, where a second
  // comment on the same text may or may not cover it depending on which came
  // first; anchors are left out, the comments' ranges are compared.
  const compareDocuments = (stage: string) => {
    const batchDocument = JSON.stringify(batch.presentation({ withAnchors: false }));
    const oracleDocument = JSON.stringify(oracle.presentation({ withAnchors: false }));
    if (batchDocument !== oracleDocument) {
      report(`${stage}, the batch left ${batchDocument}, one at a time left ${oracleDocument}`);
    }
  };
  // A tracked batch that adds a paragraph after the story's last one gives it
  // the story's last paragraph mark; one at a time, the insertion has
  // already moved that mark before the deletion of the old last paragraph
  // could claim it. The redlines differ by design; what they accept to does
  // not.
  const lastBlock = BLOCK_COUNT - 1;
  const rotatesFinalBreak =
    mode !== "direct" &&
    applied.some(
      ({ operation }) => operation.kind === "insertAfterBlock" && operation.block === lastBlock,
    ) &&
    applied.some(
      ({ operation }) => operation.kind === "deleteBlock" && operation.block === lastBlock,
    );
  // The redline itself first: the same revisions on the same text.
  if (!rotatesFinalBreak) {
    compareDocuments("as applied");
  }

  const named = new Set(
    generated.flatMap((operation) =>
      operation.kind === "mergeBlockWithNext"
        ? [operation.block, operation.block + 1]
        : [operation.block],
    ),
  );
  // A merge into a block the batch deletes joins across it, into the block
  // after it — and across every deleted block that follows, to the first one
  // the batch keeps — in either mode, as one at a time.
  const deletes = (block: number) =>
    applied.some(({ operation }) => operation.kind === "deleteBlock" && operation.block === block);
  for (const { operation } of applied) {
    if (operation.kind !== "mergeBlockWithNext") {
      continue;
    }
    for (let joined = operation.block + 1; deletes(joined); joined++) {
      named.add(joined + 1);
    }
  }
  if (mode !== "direct") {
    // Both redlines accept, and to the same document.
    const accepts = (session: OperationSession, whose: string): boolean => {
      try {
        session.acceptAll();
        return true;
      } catch (error) {
        report(`${whose} redline cannot be accepted: ${String(error).slice(0, 120)}`);
        return false;
      }
    };
    if (accepts(oracle, "the one-at-a-time") && accepts(batch, "the batch's")) {
      compareDocuments("accepted");
    }
  }
  const finalBlocks = batch.snapshot().blocks;
  for (const [block, blockId] of blockIds.entries()) {
    const text = finalBlocks.find((candidate) => candidate.id === blockId)?.text;
    if (!named.has(block) && text !== blockText(block)) {
      report(`block ${String(block)}, which no operation names, now reads ${String(text)}`);
    }
  }
  return problems;
};

/** One of every kind on block 1, covering its tokens 1 and 2. */
const KIND_SAMPLES: readonly GeneratedOperation[] = [
  { kind: "replaceInBlock", block: 1, first: 1, last: 2 },
  { kind: "replaceRange", block: 1, first: 1, last: 2 },
  { kind: "formatRange", block: 1, first: 1, last: 2 },
  { kind: "commentOnRange", block: 1, first: 1, last: 2 },
  { kind: "commentOnBlock", block: 1 },
  { kind: "splitBlock", block: 1, before: 2 },
  { kind: "mergeBlockWithNext", block: 1 },
  { kind: "deleteBlock", block: 1 },
  { kind: "replaceBlock", block: 1 },
  { kind: "setBlockParagraphProperties", block: 1 },
  { kind: "insertAfterBlock", block: 1 },
  { kind: "insertBeforeBlock", block: 1 },
];

/** The same kinds aimed elsewhere: the rest of block 1, and its neighbours. */
const NEIGHBOUR_SAMPLES: readonly GeneratedOperation[] = [
  { kind: "replaceInBlock", block: 1, first: 2, last: 3 },
  { kind: "replaceInBlock", block: 1, first: 3, last: 3 },
  { kind: "formatRange", block: 1, first: 0, last: 3 },
  { kind: "splitBlock", block: 1, before: 3 },
  { kind: "replaceInBlock", block: 2, first: 0, last: 0 },
  { kind: "deleteBlock", block: 2 },
  { kind: "setBlockParagraphProperties", block: 2 },
  { kind: "insertBeforeBlock", block: 2 },
  { kind: "mergeBlockWithNext", block: 0 },
  { kind: "mergeBlockWithNext", block: 2 },
  { kind: "deleteBlock", block: 0 },
];

describe("a batch of two operations on one block or its neighbours", () => {
  for (const mode of MODES) {
    test(`either refuses the conflict or applies them as one at a time would (${mode})`, async () => {
      const problems: string[] = [];
      for (const first of KIND_SAMPLES) {
        for (const second of [...KIND_SAMPLES, ...NEIGHBOUR_SAMPLES]) {
          problems.push(...(await batchAgainstOneAtATime([first, second], mode)));
          problems.push(...(await batchAgainstOneAtATime([second, first], mode)));
        }
      }
      expect(problems).toEqual([]);
    }, 240_000);
  }
});

describe("a batch that deletes the story's last paragraph and inserts after it", () => {
  // Deleted alone, the last paragraph keeps its mark and stays a paragraph to
  // format. The insertion after it lands first, so the deletion then takes
  // its mark as well, and its properties with it: formatting it is refused,
  // whichever of the three comes last.
  const last = BLOCK_COUNT - 1;
  const trio: readonly GeneratedOperation[] = [
    { kind: "deleteBlock", block: last },
    { kind: "insertAfterBlock", block: last },
    { kind: "setBlockParagraphProperties", block: last },
  ];
  const orders = [
    [0, 1, 2],
    [0, 2, 1],
    [1, 0, 2],
    [1, 2, 0],
    [2, 0, 1],
    [2, 1, 0],
  ].map((order) => order.flatMap((index) => trio[index] ?? []));

  for (const mode of MODES) {
    test(`refuses the conflict and applies the rest as one at a time would (${mode})`, async () => {
      const problems: string[] = [];
      // The counterexample the random batch below found.
      const found: GeneratedOperation[] = [
        { kind: "deleteBlock", block: last },
        { kind: "insertAfterBlock", block: last },
        { kind: "splitBlock", block: 0, before: 1 },
        { kind: "setBlockParagraphProperties", block: last },
      ];
      for (const generated of [...orders, found]) {
        problems.push(...(await batchAgainstOneAtATime(generated, mode)));
      }
      expect(problems).toEqual([]);
    });
  }

  test("refuses the paragraph properties once the insertion follows the deletion", async () => {
    const session = await freshSession();
    const blockId = session.snapshot().blocks[last]?.id ?? "";
    const result = session.apply("tracked-changes", [
      { id: "delete", type: "deleteBlock", blockId },
      { id: "insert", type: "insertAfterBlock", blockId, text: "Added." },
      {
        id: "center",
        type: "setBlockParagraphProperties",
        blockId,
        properties: { alignment: "center" },
      },
    ]);
    expect(result.skipped).toEqual([
      {
        id: "center",
        reason: "overlappingOperation",
        message: 'operation "delete", earlier in this batch, already claims its target.',
      },
    ]);
  });
});

describe("a merge into blocks the batch deletes", () => {
  const last = BLOCK_COUNT - 1;
  const cases: Record<string, GeneratedOperation[]> = {
    "joins across a deleted block": [
      { kind: "deleteBlock", block: 1 },
      { kind: "mergeBlockWithNext", block: 0 },
    ],
    "joins across a run of deleted blocks": [
      { kind: "deleteBlock", block: 1 },
      { kind: "mergeBlockWithNext", block: 0 },
      { kind: "deleteBlock", block: 2 },
    ],
    // The counterexample the random batch below found: nothing follows the
    // deleted block to join into.
    "deleted to the story's end": [
      { kind: "deleteBlock", block: last },
      { kind: "mergeBlockWithNext", block: last - 1 },
      { kind: "replaceInBlock", block: 0, first: 0, last: 0 },
      { kind: "splitBlock", block: 0, before: 1 },
    ],
  };
  for (const mode of MODES) {
    for (const [name, generated] of Object.entries(cases)) {
      test(`${name}: as one at a time would (${mode})`, async () => {
        expect(await batchAgainstOneAtATime(generated, mode)).toEqual([]);
      });
    }
  }

  test("applied directly, a merge with nothing left to join is refused", async () => {
    const session = await freshSession();
    const blocks = session.snapshot().blocks;
    const result = session.apply("direct", [
      { id: "delete", type: "deleteBlock", blockId: blocks[last]?.id ?? "" },
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: blocks[last - 1]?.id ?? "",
        separator: " ",
      },
    ]);
    expect(result.skipped.map(({ id, reason }) => ({ id, reason }))).toEqual([
      { id: "merge", reason: "unsupportedBlock" },
    ]);
  });

  test("tracked, a merge into deletions that run to the end writes no separator", async () => {
    const session = await freshSession();
    const blocks = session.snapshot().blocks;
    const result = session.apply("tracked-changes", [
      { id: "delete", type: "deleteBlock", blockId: blocks[last]?.id ?? "" },
      {
        id: "merge",
        type: "mergeBlockWithNext",
        blockId: blocks[last - 1]?.id ?? "",
        separator: " ",
      },
    ]);
    expect(result.skipped).toEqual([]);
    session.acceptAll();
    const accepted = session.snapshot().blocks.map(({ text }) => text);
    expect(accepted.at(-1)).toBe(blockText(last - 1));
  });
});

const spanArbitrary = fc
  .record({
    block: fc.nat({ max: BLOCK_COUNT - 1 }),
    first: fc.nat({ max: TOKEN_COUNT - 1 }),
    length: fc.nat({ max: 1 }),
  })
  .map(({ block, first, length }) => ({
    block,
    first,
    last: Math.min(TOKEN_COUNT - 1, first + length),
  }));

const operationArbitrary: fc.Arbitrary<GeneratedOperation> = fc.oneof(
  fc
    .tuple(
      fc.constantFrom(
        "replaceInBlock" as const,
        "replaceRange" as const,
        "formatRange" as const,
        "commentOnRange" as const,
      ),
      spanArbitrary,
    )
    .map(([kind, { block, first, last }]) => ({ kind, block, first, last })),
  fc
    .record({
      block: fc.nat({ max: BLOCK_COUNT - 1 }),
      before: fc.integer({ min: 1, max: TOKEN_COUNT - 1 }),
    })
    .map(({ block, before }) => ({ kind: "splitBlock" as const, block, before })),
  fc
    .tuple(
      fc.constantFrom(
        "commentOnBlock" as const,
        "deleteBlock" as const,
        "replaceBlock" as const,
        "setBlockParagraphProperties" as const,
        "insertAfterBlock" as const,
        "insertBeforeBlock" as const,
      ),
      fc.nat({ max: BLOCK_COUNT - 1 }),
    )
    .map(([kind, block]) => ({ kind, block })),
  fc.nat({ max: BLOCK_COUNT - 2 }).map((block) => ({ kind: "mergeBlockWithNext" as const, block })),
);

describe("two paragraph-property operations on one block", () => {
  test.each(MODES)("refuse the later one and keep the earlier one's values (%s)", async (mode) => {
    const session = await freshSession();
    const blockId = session.snapshot().blocks[1]?.id ?? "";
    const align = (id: string, alignment: "center" | "right"): FolioDocumentOperation => ({
      id,
      type: "setBlockParagraphProperties",
      blockId,
      properties: { alignment },
    });
    const result = session.apply(mode, [align("center", "center"), align("right", "right")]);
    expect(result.applied.map(({ id }) => id)).toEqual(["center"]);
    expect(result.skipped).toEqual([
      {
        id: "right",
        reason: "overlappingOperation",
        message: 'operation "center", earlier in this batch, already claims its target.',
      },
    ]);
    expect(session.state.doc.child(1).attrs["alignment"]).toBe("center");
  });
});

describe("a random batch with overlapping, nested and duplicate targets", () => {
  test("refuses each conflict and applies the rest as one at a time would", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(operationArbitrary, { minLength: 2, maxLength: 7 }),
        fc.constantFrom(...MODES),
        async (generated, mode) => {
          expect(await batchAgainstOneAtATime(generated, mode)).toEqual([]);
        },
      ),
      { numRuns: 150 },
    );
  }, 300_000);
});
