/**
 * Applies a document-operation batch to a ProseMirror story: runs the AI edit
 * applier on a held view, checks the result, and commits it or refuses it.
 * The batch contract lives in `./document-operations`.
 */
import type { Node as PMNode } from "prosemirror-model";
import type { Transaction } from "prosemirror-state";

import type { ValidateDocumentModelIssue } from "@stll/docx-core";

import {
  applyFolioAIEditOperationsInPackage,
  previewFolioAIEditOperationsWithResult,
} from "./ai-edits/apply";
import type { FolioAIEditApplyOutcome } from "./ai-edits/applyTypes";
import {
  describeModelError,
  findIntroducedModelErrors,
} from "./ai-edits/result-validation";
import { trackedBatchFidelityError } from "./ai-edits/tracked-batch-fidelity";
import type { FolioAIEditSkippedOperation } from "./ai-edits/types";
import {
  type ApplyFolioDocumentOperationsOptions,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationResult,
  getFolioDocumentOperationIssues,
  getFolioDocumentOperationReceiptsForStory,
  parseFolioDocumentOperationBatch,
  withDistinctSuggestionIds,
} from "./document-operations";
import {
  NO_NUMBERING_IDS_IN_USE,
  type NumberingIdsInUse,
} from "./docx/numberingIds";
import { getSuggestions } from "./prosemirror/commands/comments";
import { completeNumberingForDoc } from "./prosemirror/listInstanceReferences";
import { getStatedDocumentNumbering } from "./prosemirror/plugins/documentNumbering";

/** One run of the applier whose result has not reached the caller's view yet. */
type BatchAttempt = {
  outcome: FolioAIEditApplyOutcome;
  /** The story the run would leave, or `null` when it changed nothing. */
  doc: PMNode | null;
  /** Comment ids the run allocated, provisional ones included. */
  commentIds: readonly number[];
  /** Hand the run's transaction to the caller's view; a no-op for a preview. */
  commit: () => void;
};

/** An attempt after the result check: what to report and what to commit. */
type GuardedBatch = {
  fidelityRefused?: true;
  outcome: FolioAIEditApplyOutcome;
  commit: () => void;
};

const NO_COMMIT = (): void => {};

/** Skips reported in the order of the operations they belong to. */
const inOperationOrder = (
  operations: readonly FolioDocumentOperation[],
  skipped: readonly FolioAIEditSkippedOperation[],
): FolioAIEditSkippedOperation[] => {
  const indexById = new Map(operations.map(({ id }, index) => [id, index]));
  return [...skipped].sort(
    (left, right) =>
      (indexById.get(left.id) ?? -1) - (indexById.get(right.id) ?? -1),
  );
};

export const applyFolioDocumentOperations = (
  input: ApplyFolioDocumentOperationsOptions,
): FolioDocumentOperationResult =>
  applyFolioDocumentOperationsInPackage({
    ...input,
    numberingIdsElsewhere: NO_NUMBERING_IDS_IN_USE,
  });

/**
 * Options of a caller that holds every story of the package: the numbering
 * ids (`w:num`, `w:abstractNum`) the other stories define. A list the batch
 * starts takes none of them, so its id names it in every story and survives a
 * save.
 */
type ApplyFolioDocumentOperationsInPackageOptions =
  ApplyFolioDocumentOperationsOptions & {
    numberingIdsElsewhere: NumberingIdsInUse;
  };

/** {@link applyFolioDocumentOperations} for a caller holding every story. Internal. */
export const applyFolioDocumentOperationsInPackage = ({
  view,
  snapshot,
  batch,
  story = "main",
  author,
  createCommentId,
  createUndoHandle,
  revisionStamp,
  wordDiff,
  tableTemplates,
  replacementBackground,
  undefinedReferences,
  pendingSuggestionIds,
  numberingIdsElsewhere,
  restagesSuggestions,
}: ApplyFolioDocumentOperationsInPackageOptions): FolioDocumentOperationResult => {
  const validatedBatch = parseFolioDocumentOperationBatch(batch);
  const parsedBatch =
    restagesSuggestions === true
      ? validatedBatch
      : withDistinctSuggestionIds(
          validatedBatch,
          new Set([
            ...getSuggestions(view.state).map(
              ({ suggestionId }) => suggestionId,
            ),
            ...(pendingSuggestionIds ?? []),
          ]),
        );
  const isAtomic = parsedBatch.atomic === true;
  const beforeDoc = view.state.doc;
  // The package's own definitions, before the batch. A run's result can carry
  // a reference to an instance only its own paragraphs define (a "start new
  // list" request mints one), so the check completes this against each run's
  // own doc rather than holding one map for every run.
  const packageNumbering = getStatedDocumentNumbering(view.state);

  const attempt = (
    operations: readonly FolioDocumentOperation[],
    preview: boolean,
  ): BatchAttempt => {
    const common = {
      snapshot,
      operations,
      mode: parsedBatch.mode ?? "tracked-changes",
      ...(parsedBatch.existingRevisions === "refuse-cross-author" && {
        refuseCrossAuthorRevisions: true,
      }),
      ...(author !== undefined && { author }),
      ...(revisionStamp !== undefined && { revisionStamp }),
      ...(wordDiff !== undefined && { wordDiff }),
      ...(tableTemplates !== undefined && { tableTemplates }),
      ...(replacementBackground !== undefined && { replacementBackground }),
      numberingIdsElsewhere,
      undefinedReferences,
    } as const;
    if (preview) {
      const previewed = previewFolioAIEditOperationsWithResult(
        {
          ...common,
          view,
          ...(createCommentId !== undefined && { createCommentId }),
        },
        parsedBatch.dryRun && common.mode === "tracked-changes"
          ? "bounded"
          : "coarse",
      );
      return { ...previewed, commit: NO_COMMIT };
    }
    // The run dispatches into a holding view, not the caller's: its result is
    // checked first, and a refused result must never reach a live editor.
    const held: { transaction?: Transaction } = {};
    const commentIds: number[] = [];
    // canonical-gap: publicOps.comments
    // canonical-gap: publicOps.unsupportedInline
    // canonical-gap: pm-document-operations
    const outcome = applyFolioAIEditOperationsInPackage({
      ...common,
      view: {
        state: view.state,
        dispatch: (transaction) => {
          held.transaction = transaction;
        },
      },
      ...(createCommentId !== undefined && {
        createCommentId: (text: string) => {
          const id = createCommentId(text);
          commentIds.push(id);
          return id;
        },
      }),
    });
    return {
      outcome,
      doc: held.transaction?.doc ?? null,
      commentIds,
      commit: () => {
        if (held.transaction !== undefined) {
          view.dispatch(held.transaction);
        }
      },
    };
  };

  const introducedErrors = (run: BatchAttempt) => {
    if (run.doc === null || run.outcome.applied.length === 0) {
      return [];
    }
    // `undefined` still means the state cannot say (no numbering plugin);
    // otherwise complete the package's definitions against this run's own
    // doc, so an instance a "start new list" request just minted validates.
    const numbering =
      packageNumbering === undefined
        ? undefined
        : (completeNumberingForDoc(packageNumbering ?? undefined, run.doc) ??
          null);
    return findIntroducedModelErrors(beforeDoc, run.doc, {
      numbering,
      createdCommentIds: run.commentIds,
      ...(author !== undefined && { commentAuthor: author }),
    });
  };

  /**
   * An attempt on the refusal path, where the applier runs on subsets of the
   * batch it was not asked to apply: a throw there is one more reason to
   * refuse, never a reason to fail a batch the first run could apply.
   */
  const checkedAttempt = (
    operations: readonly FolioDocumentOperation[],
    preview: boolean,
  ): {
    run: BatchAttempt | null;
    error: ValidateDocumentModelIssue | undefined;
  } => {
    try {
      const run = attempt(operations, preview);
      return { run, error: introducedErrors(run)[0] };
    } catch (error) {
      return {
        run: null,
        error: {
          path: "package.document",
          message: `the operation could not be applied (${
            error instanceof Error ? error.message : String(error)
          }).`,
          severity: "error",
        },
      };
    }
  };

  /**
   * Which applied operations produced the unsaveable result: each one is
   * previewed alone against the unchanged story. When none fails alone, the
   * failure is their combination, and every one of them is refused.
   */
  const refuseInvalidResult = (
    run: BatchAttempt,
    firstError: ValidateDocumentModelIssue,
  ): FolioAIEditSkippedOperation[] => {
    const operationsById = new Map(
      parsedBatch.operations.map((operation) => [operation.id, operation]),
    );
    const refused: FolioAIEditSkippedOperation[] = [];
    for (const { id } of run.outcome.applied) {
      const operation = operationsById.get(id);
      if (operation === undefined) {
        continue;
      }
      const alone = checkedAttempt([operation], true).error;
      if (alone !== undefined) {
        refused.push({
          id,
          reason: "invalidResult",
          message: describeModelError(alone),
        });
      }
    }
    if (refused.length > 0) {
      return refused;
    }
    return run.outcome.applied.map(({ id }) => ({
      id,
      reason: "invalidResult",
      message: describeModelError(firstError),
    }));
  };

  /**
   * A fidelity failure invalidates the batch. Preserve the refusal details
   * for operations already skipped by the applier or save-time validator.
   */
  const refuseFidelityLoss = (
    run: BatchAttempt,
    skipped: readonly FolioAIEditSkippedOperation[] = [],
  ): GuardedBatch | null => {
    if (
      run.doc === null ||
      parsedBatch.mode === "direct" ||
      parsedBatch.mode === "suggested"
    ) {
      return null;
    }
    const message = trackedBatchFidelityError({
      before: view.state,
      after: run.doc,
      createdCommentIds: run.commentIds,
      retiredCommentIds: run.outcome.applied.flatMap(
        ({ retiredCommentIds }) => retiredCommentIds ?? [],
      ),
    });
    if (message === null) return null;
    return {
      fidelityRefused: true,
      outcome: {
        ...run.outcome,
        applied: [],
        skipped: inOperationOrder(parsedBatch.operations, [
          ...run.outcome.skipped,
          ...skipped,
          ...run.outcome.applied.map(({ id }): FolioAIEditSkippedOperation => ({
            id,
            reason: "invalidResult",
            message,
          })),
        ]),
      },
      commit: NO_COMMIT,
    };
  };

  /** Check the held result before dispatching any operation to the caller. */
  const guardedAttempt = (preview: boolean): GuardedBatch => {
    const first = attempt(parsedBatch.operations, preview);
    const [firstError] = introducedErrors(first);
    if (firstError === undefined) {
      // Atomic preflight uses a coarse diff. Verify the held commit instead.
      const fidelity =
        preview && !parsedBatch.dryRun ? null : refuseFidelityLoss(first);
      if (fidelity !== null) return fidelity;
      return {
        outcome: {
          ...first.outcome,
          skipped: inOperationOrder(
            parsedBatch.operations,
            first.outcome.skipped,
          ),
        },
        commit: first.commit,
      };
    }
    const refused = refuseInvalidResult(first, firstError);
    const refusedIds = new Set(refused.map(({ id }) => id));
    if (isAtomic) {
      // Only the refused operations name a reason of their own; the atomic
      // result reports every other one as rejected with the batch.
      return {
        outcome: {
          ...first.outcome,
          applied: [],
          skipped: inOperationOrder(parsedBatch.operations, [
            ...first.outcome.skipped,
            ...refused,
          ]),
        },
        commit: NO_COMMIT,
      };
    }
    const second = checkedAttempt(
      parsedBatch.operations.filter(({ id }) => !refusedIds.has(id)),
      preview,
    );
    if (second.run === null || second.error !== undefined) {
      // The rest still fails without the operations that failed alone: an
      // interaction the per-operation check cannot pin down. Refuse it all.
      const reported = second.run?.outcome ?? first.outcome;
      return {
        outcome: {
          ...reported,
          applied: [],
          skipped: inOperationOrder(parsedBatch.operations, [
            ...reported.skipped.filter(({ id }) => !refusedIds.has(id)),
            ...refused,
            ...reported.applied
              .filter(({ id }) => !refusedIds.has(id))
              .map(({ id }): FolioAIEditSkippedOperation => ({
                id,
                reason: "invalidResult",
                message: describeModelError(second.error ?? firstError),
              })),
          ]),
        },
        commit: NO_COMMIT,
      };
    }
    const fidelity = refuseFidelityLoss(second.run, refused);
    if (fidelity !== null) return fidelity;
    return {
      outcome: {
        ...second.run.outcome,
        skipped: inOperationOrder(parsedBatch.operations, [
          ...second.run.outcome.skipped,
          ...refused,
        ]),
      },
      commit: second.run.commit,
    };
  };

  const preview = () => guardedAttempt(true).outcome;

  const atomicResult = (
    previewResult: FolioAIEditApplyOutcome,
    status: "previewed" | "rejected",
  ): FolioDocumentOperationResult => {
    const skippedById = new Map(
      previewResult.skipped.map((operation) => [operation.id, operation]),
    );
    const skipped = parsedBatch.operations.map(
      ({ id }): FolioAIEditSkippedOperation =>
        skippedById.get(id) ?? { id, reason: "atomicBatchRejected" },
    );
    return {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      status,
      applied: [],
      skipped,
      issues: getFolioDocumentOperationIssues(parsedBatch.operations, skipped),
      receipts: [],
      ...(previewResult.normalizations !== undefined && {
        normalizations: previewResult.normalizations,
      }),
      undoHandle: null,
      nextRevisionId: previewResult.nextRevisionId,
    };
  };

  if (parsedBatch.dryRun === true) {
    const previewResult = preview();
    if (isAtomic && previewResult.skipped.length > 0) {
      return atomicResult(previewResult, "previewed");
    }
    return {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      status: "previewed",
      applied: previewResult.applied.map(
        ({ id, paragraphIdentity, retiredCommentIds }) => ({
          id,
          ...(paragraphIdentity === undefined ? {} : { paragraphIdentity }),
          ...(retiredCommentIds === undefined ? {} : { retiredCommentIds }),
        }),
      ),
      skipped: previewResult.skipped,
      issues: getFolioDocumentOperationIssues(
        parsedBatch.operations,
        previewResult.skipped,
      ),
      receipts: getFolioDocumentOperationReceiptsForStory({
        operations: parsedBatch.operations,
        applied: previewResult.applied,
        story,
      }),
      ...(previewResult.normalizations !== undefined && {
        normalizations: previewResult.normalizations,
      }),
      undoHandle: null,
      nextRevisionId: previewResult.nextRevisionId,
    };
  }

  if (isAtomic) {
    const previewResult = preview();
    if (previewResult.skipped.length > 0) {
      return atomicResult(previewResult, "rejected");
    }
  }

  const guarded = guardedAttempt(false);
  const { heldRevisionConflict, ...result } = guarded.outcome;
  if (heldRevisionConflict) return atomicResult(result, "rejected");
  if (
    guarded.fidelityRefused === true ||
    (isAtomic &&
      result.skipped.some(({ reason }) => reason === "invalidResult"))
  ) {
    return atomicResult(result, "rejected");
  }
  guarded.commit();
  return {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    status: "committed",
    ...result,
    issues: getFolioDocumentOperationIssues(
      parsedBatch.operations,
      result.skipped,
    ),
    receipts: getFolioDocumentOperationReceiptsForStory({
      operations: parsedBatch.operations,
      applied: result.applied,
      story,
    }),
    undoHandle:
      result.applied.length > 0 && createUndoHandle !== undefined
        ? createUndoHandle()
        : null,
  };
};
