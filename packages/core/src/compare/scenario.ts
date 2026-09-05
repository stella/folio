/**
 * Edit scripts: ground truth for the compare engine, by construction.
 *
 * A property test cannot assert "this redline is right" against a document it
 * only guessed at. It can assert it against a document it BUILT: take a real
 * base, apply a scripted edit directly (no tracked changes), and the result is
 * a target whose difference from the base is known exactly. `compareDocx` then
 * has to rediscover that difference, and the script says what it should have
 * found.
 *
 * Every step is applied through the same headless applier `compareDocx` writes
 * its own operations through, in `"direct"` mode — so the target is an ordinary
 * document, not one carrying revisions the comparison would have to resolve
 * first.
 *
 * Steps address blocks by their index in the base story's block list rather
 * than by id, so a generator can pick targets without first parsing the
 * document. A step that does not resolve is reported in `unresolved`, never
 * dropped: a script silently doing less than it says would weaken every
 * property built on it.
 */

import { panic, Result } from "better-result";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAITextRangeHandle } from "../ai-edits/snapshot";
import type {
  FolioAIBlock,
  FolioAIEditOperation,
  FolioAIInlineFormatting,
} from "../ai-edits/types";
import { CompareDocxParseError, CompareDocxSerializeError } from "./types";

export type EditScriptStep =
  | { type: "insertParagraphAfter"; blockIndex: number; text: string }
  | { type: "deleteParagraph"; blockIndex: number }
  | { type: "replaceWords"; blockIndex: number; find: string; replace: string }
  /** Delete the block at `blockIndex` and reinsert its text before `beforeBlockIndex`. */
  | { type: "moveParagraph"; blockIndex: number; beforeBlockIndex: number }
  | {
      type: "formatRange";
      blockIndex: number;
      startOffset: number;
      endOffset: number;
      formatting: FolioAIInlineFormatting;
    }
  | { type: "insertTableRow"; blockIndex: number; cellTexts: readonly string[] }
  | { type: "deleteTableRow"; blockIndex: number }
  | { type: "editTableCell"; blockIndex: number; text: string };

export type EditScript = readonly EditScriptStep[];

export const EDIT_SCRIPT_UNRESOLVED_REASONS = Object.freeze([
  "block-out-of-range",
  "block-not-in-table",
  "range-out-of-bounds",
  "text-not-found",
  /** The applier refused the operation the step produced. */
  "refused",
] as const);

export type EditScriptUnresolvedReason = (typeof EDIT_SCRIPT_UNRESOLVED_REASONS)[number];

export type EditScriptUnresolvedStep = {
  step: EditScriptStep;
  reason: EditScriptUnresolvedReason;
};

export type EditScriptResult = {
  /** The base document with every applied step written in directly. */
  buffer: ArrayBuffer;
  /** The steps that reached the document, in script order. */
  applied: readonly EditScriptStep[];
  unresolved: readonly EditScriptUnresolvedStep[];
};

type PlannedStep = {
  step: EditScriptStep;
  operations: readonly FolioAIEditOperation[];
};

type StepPlan =
  | { status: "planned"; operations: readonly FolioAIEditOperation[] }
  | { status: "unresolved"; reason: EditScriptUnresolvedReason };

const blockAt = (blocks: readonly FolioAIBlock[], index: number): FolioAIBlock | undefined =>
  Number.isInteger(index) && index >= 0 ? blocks.at(index) : undefined;

type PlanStepOptions = {
  step: EditScriptStep;
  blocks: readonly FolioAIBlock[];
  nextOperationId: () => string;
};

const planStep = ({ step, blocks, nextOperationId }: PlanStepOptions): StepPlan => {
  const block = blockAt(blocks, step.blockIndex);
  if (!block) {
    return { status: "unresolved", reason: "block-out-of-range" };
  }

  switch (step.type) {
    case "insertParagraphAfter":
      return {
        status: "planned",
        operations: [
          { id: nextOperationId(), type: "insertAfterBlock", blockId: block.id, text: step.text },
        ],
      };
    case "deleteParagraph":
      return {
        status: "planned",
        operations: [{ id: nextOperationId(), type: "deleteBlock", blockId: block.id }],
      };
    case "replaceWords":
      return block.text.includes(step.find)
        ? {
            status: "planned",
            operations: [
              {
                id: nextOperationId(),
                type: "replaceInBlock",
                blockId: block.id,
                find: step.find,
                replace: step.replace,
              },
            ],
          }
        : { status: "unresolved", reason: "text-not-found" };
    case "moveParagraph": {
      const destination = blockAt(blocks, step.beforeBlockIndex);
      if (!destination || destination.id === block.id) {
        return { status: "unresolved", reason: "block-out-of-range" };
      }
      return {
        status: "planned",
        operations: [
          { id: nextOperationId(), type: "deleteBlock", blockId: block.id },
          {
            id: nextOperationId(),
            type: "insertBeforeBlock",
            blockId: destination.id,
            text: block.text,
          },
        ],
      };
    }
    case "formatRange": {
      const range = createFolioAITextRangeHandle({
        blockId: block.id,
        text: block.text,
        startOffset: step.startOffset,
        endOffset: step.endOffset,
      });
      return range
        ? {
            status: "planned",
            operations: [
              { id: nextOperationId(), type: "formatRange", range, formatting: step.formatting },
            ],
          }
        : { status: "unresolved", reason: "range-out-of-bounds" };
    }
    case "insertTableRow":
      return block.table
        ? {
            status: "planned",
            operations: [
              {
                id: nextOperationId(),
                type: "insertTableRow",
                blockId: block.id,
                position: "after",
                cellTexts: [...step.cellTexts],
              },
            ],
          }
        : { status: "unresolved", reason: "block-not-in-table" };
    case "deleteTableRow":
      return block.table
        ? {
            status: "planned",
            operations: [
              { id: nextOperationId(), type: "deleteTableRow", blockId: block.id },
            ],
          }
        : { status: "unresolved", reason: "block-not-in-table" };
    case "editTableCell":
      return block.table
        ? {
            status: "planned",
            operations: [
              { id: nextOperationId(), type: "replaceBlock", blockId: block.id, text: step.text },
            ],
          }
        : { status: "unresolved", reason: "block-not-in-table" };
    default: {
      const unreachable: never = step;
      return panic("Unhandled edit script step", { step: unreachable });
    }
  }
};

/**
 * Apply `script` to `base` directly and return the resulting target document.
 */
export const applyEditScript = async (
  base: ArrayBuffer,
  script: EditScript,
): Promise<Result<EditScriptResult, CompareDocxParseError | CompareDocxSerializeError>> => {
  const parsed = await Result.tryPromise({
    try: async () => await FolioDocxReviewer.fromBuffer(base),
    catch: (cause) =>
      new CompareDocxParseError({
        message: "The edit script's base document could not be parsed.",
        side: "base",
        cause,
      }),
  });
  if (parsed.isErr()) {
    return Result.err(parsed.error);
  }

  const reviewer = parsed.value;
  const snapshot = reviewer.snapshot();
  const unresolved: EditScriptUnresolvedStep[] = [];
  const planned: PlannedStep[] = [];
  let operationSequence = 0;
  const nextOperationId = (): string => `scenario-${++operationSequence}`;

  for (const step of script) {
    const plan = planStep({ step, blocks: snapshot.blocks, nextOperationId });
    if (plan.status === "unresolved") {
      unresolved.push({ step, reason: plan.reason });
      continue;
    }
    planned.push({ step, operations: plan.operations });
  }

  const { skipped } = reviewer.applyOperations(
    planned.flatMap(({ operations }) => [...operations]),
    { mode: "direct", snapshot },
  );
  const skippedIds = new Set(skipped.map(({ id }) => id));
  const applied: EditScriptStep[] = [];
  for (const { step, operations } of planned) {
    if (operations.some(({ id }) => skippedIds.has(id))) {
      unresolved.push({ step, reason: "refused" });
      continue;
    }
    applied.push(step);
  }

  const serialized = await Result.tryPromise({
    try: async () => await reviewer.toBuffer(),
    catch: (cause) =>
      new CompareDocxSerializeError({
        message: "The edit script's target document could not be serialized.",
        cause,
      }),
  });
  return serialized.isErr()
    ? Result.err(serialized.error)
    : Result.ok({ buffer: serialized.value, applied, unresolved });
};
