/**
 * Compare two `.docx` buffers and produce a third buffer whose text
 * differences are represented as tracked changes.
 *
 * Each matched editable story is processed independently through the shared
 * block alignment and document-operation paths. Package parts that exist on
 * only one side are reported because creating or removing those parts is a
 * distinct package-level operation.
 */

import { panic, TaggedError } from "better-result";

import {
  FolioDocxReviewer,
  isFolioResolvedReviewedView,
  type FolioDocumentStoryHandle,
  type FolioResolvedReviewedView,
} from "./ai-edits/headless";
import type {
  FolioAIBlock,
  FolioAIEditAppliedOperation,
  FolioAIEditOperation,
  FolioAIEditSkippedOperation,
  FolioAIEditSnapshot,
} from "./ai-edits/types";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "./document-operations";
import { pairFolioDocumentStories } from "./document-stories";
import {
  resolveFolioDocumentPrivacyTransforms,
  rewriteDocxMetadataPrivacy,
  type FolioDocumentPrivacyOptions,
  type FolioDocumentPrivacyReport,
} from "./docx/metadataPrivacy";
import { alignFolioBlocks, type FolioAlignedBlockEvent } from "./version-comparison";

/** Options for {@link generateRedlineDocx}. */
export type GenerateRedlineDocxOptions = {
  /** Author recorded on the generated tracked changes. (default: `"folio compare"`) */
  author?: string;
  /** Resolved base input state. (default: `"final"`) */
  baseView?: FolioResolvedReviewedView;
  /** Resolved revised input state. (default: `"final"`) */
  revisedView?: FolioResolvedReviewedView;
  /** Optional output-only package-metadata privacy transforms. */
  privacy?: FolioDocumentPrivacyOptions;
};

export class InvalidGenerateRedlineDocxOptionsError extends TaggedError(
  "InvalidGenerateRedlineDocxOptionsError",
)<{
  message: string;
  option: "baseView" | "revisedView";
  receivedValue: unknown;
}>() {}

export type GenerateRedlineUnprocessedStory = {
  baseStory: FolioDocumentStoryHandle | null;
  revisedStory: FolioDocumentStoryHandle | null;
  reason: "missing-base-story" | "missing-revised-story";
};

/** Result of {@link generateRedlineDocx}. */
export type GenerateRedlineDocxResult = {
  /** The base package with generated tracked changes. */
  buffer: ArrayBuffer;
  /** Operations applied across every matched story. */
  applied: FolioAIEditAppliedOperation[];
  /** Block operations that could not be applied. */
  skipped: FolioAIEditSkippedOperation[];
  /** Package parts that could not be represented as story-scoped text edits. */
  unprocessedStories: GenerateRedlineUnprocessedStory[];
  /** Privacy transforms applied to the generated package. */
  privacyReport: FolioDocumentPrivacyReport;
};

const nextBaseBlockIdByIndex = (events: readonly FolioAlignedBlockEvent[]): (string | null)[] => {
  const nextIds = Array.from<string | null>({ length: events.length });
  let nextId: string | null = null;
  for (let index = events.length - 1; index >= 0; index--) {
    nextIds[index] = nextId;
    const event = events[index];
    if (event?.type === "pair") {
      nextId = event.baseBlock.id;
    } else if (event?.type === "baseOnly") {
      nextId = event.block.id;
    }
  }
  return nextIds;
};

type BuildRedlineOperationsOptions = {
  baseSnapshot: FolioAIEditSnapshot;
  revisedBlocks: readonly FolioAIBlock[];
  nextOperationId: () => string;
};

const buildRedlineOperations = ({
  baseSnapshot,
  revisedBlocks,
  nextOperationId,
}: BuildRedlineOperationsOptions): FolioAIEditOperation[] => {
  const events = alignFolioBlocks(baseSnapshot.blocks, revisedBlocks);
  const anchorIds = nextBaseBlockIdByIndex(events);
  const operations: FolioAIEditOperation[] = [];
  const trailingAdditions: { text: string; styleId?: string }[] = [];
  const lastBaseBlockId = baseSnapshot.blocks.at(-1)?.id ?? null;

  events.forEach((event, eventIndex) => {
    if (event.type === "pair") {
      if (event.baseBlock.text !== event.revisedBlock.text) {
        operations.push({
          id: nextOperationId(),
          type: "replaceBlock",
          blockId: event.baseBlock.id,
          text: event.revisedBlock.text,
        });
      }
      return;
    }
    if (event.type === "baseOnly") {
      operations.push({
        id: nextOperationId(),
        type: "deleteBlock",
        blockId: event.block.id,
      });
      return;
    }
    const anchorId = anchorIds[eventIndex] ?? null;
    if (anchorId === null) {
      trailingAdditions.push({
        text: event.block.text,
        ...(event.block.styleId !== undefined && { styleId: event.block.styleId }),
      });
      return;
    }
    operations.push({
      id: nextOperationId(),
      type: "insertBeforeBlock",
      blockId: anchorId,
      text: event.block.text,
      ...(event.block.styleId !== undefined && { styleId: event.block.styleId }),
    });
  });

  if (lastBaseBlockId === null && baseSnapshot.emptyDocumentAnchorId !== undefined) {
    const firstAddition = trailingAdditions.shift();
    if (firstAddition !== undefined) {
      operations.push({
        id: nextOperationId(),
        type: "replaceBlock",
        blockId: baseSnapshot.emptyDocumentAnchorId,
        text: firstAddition.text,
        ...(firstAddition.styleId !== undefined && { styleId: firstAddition.styleId }),
      });
    }
  }

  for (const addition of trailingAdditions) {
    operations.push({
      id: nextOperationId(),
      type: "insertAfterBlock",
      blockId: lastBaseBlockId ?? baseSnapshot.emptyDocumentAnchorId ?? "redline-unanchored",
      text: addition.text,
      ...(addition.styleId !== undefined && { styleId: addition.styleId }),
    });
  }

  return operations;
};

const resolveInputView = (
  value: unknown,
  option: "baseView" | "revisedView",
): FolioResolvedReviewedView => {
  if (value === undefined) {
    return "final";
  }
  if (!isFolioResolvedReviewedView(value)) {
    throw new InvalidGenerateRedlineDocxOptionsError({
      message: `${option} must be original or final.`,
      option,
      receivedValue: value,
    });
  }
  return value;
};

/** Compare two buffers and return tracked changes for every matched editable story. */
export const generateRedlineDocx = async (
  base: ArrayBuffer,
  revised: ArrayBuffer,
  options: GenerateRedlineDocxOptions = {},
): Promise<GenerateRedlineDocxResult> => {
  const baseView = resolveInputView(options.baseView, "baseView");
  const revisedView = resolveInputView(options.revisedView, "revisedView");
  const privacyTransforms = resolveFolioDocumentPrivacyTransforms(
    options.privacy?.transforms ?? [],
  );
  const [baseReviewer, revisedReviewer] = await Promise.all([
    FolioDocxReviewer.fromBuffer(base, { author: options.author ?? "folio compare" }),
    FolioDocxReviewer.fromBuffer(revised),
  ]);
  const baseStories = baseReviewer.listStories().map(({ handle }) => handle);
  const revisedStories = revisedReviewer.listStories().map(({ handle }) => handle);
  for (const story of baseStories) {
    if (!baseReviewer.resolveReviewedStory({ story, view: baseView })) {
      panic("A listed base story could not be resolved");
    }
  }

  const applied: FolioAIEditAppliedOperation[] = [];
  const skipped: FolioAIEditSkippedOperation[] = [];
  const unprocessedStories: GenerateRedlineUnprocessedStory[] = [];
  let operationSequence = 0;
  const nextOperationId = () => `redline-${++operationSequence}`;

  for (const pair of pairFolioDocumentStories(baseStories, revisedStories)) {
    if (!pair.baseStory) {
      unprocessedStories.push({
        ...pair,
        reason: "missing-base-story",
      });
      continue;
    }
    if (!pair.revisedStory) {
      unprocessedStories.push({
        ...pair,
        reason: "missing-revised-story",
      });
      continue;
    }
    const baseSnapshot = baseReviewer.snapshotStory(pair.baseStory);
    const revisedSnapshot = revisedReviewer.readReviewedStory({
      story: pair.revisedStory,
      view: revisedView,
    })?.snapshot;
    if (!baseSnapshot || !revisedSnapshot) {
      panic("A matched document story could not be read");
    }
    const operations = buildRedlineOperations({
      baseSnapshot,
      revisedBlocks: revisedSnapshot.blocks,
      nextOperationId,
    });
    if (operations.length === 0) {
      continue;
    }
    const result = baseReviewer.applyDocumentOperationsToStory({
      story: pair.baseStory,
      snapshot: baseSnapshot,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations,
      },
    });
    applied.push(...result.applied);
    skipped.push(...result.skipped);
  }

  const redlineBuffer = await baseReviewer.toBuffer();
  const privacyResult =
    privacyTransforms.length === 0
      ? {
          buffer: redlineBuffer,
          privacyReport: { appliedTransforms: [], removedMetadataProperties: [] },
        }
      : await rewriteDocxMetadataPrivacy(redlineBuffer, { transforms: privacyTransforms });
  return {
    buffer: privacyResult.buffer,
    applied,
    skipped,
    unprocessedStories,
    privacyReport: privacyResult.privacyReport,
  };
};
