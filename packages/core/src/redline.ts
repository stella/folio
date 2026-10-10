/**
 * Compare two `.docx` buffers and produce a third buffer whose text
 * and supported inline-formatting differences are represented as tracked changes.
 *
 * Each matched editable story is processed independently through the shared
 * block alignment and document-operation paths. Package parts that exist on
 * only one side are reported because creating or removing those parts is a
 * distinct package-level operation.
 *
 * @packageDocumentation
 */

import { panic, TaggedError } from "better-result";

import {
  FolioDocxReviewer,
  getFolioDocxComparisonAccess,
  isFolioResolvedReviewedView,
  type FolioDocumentStoryHandle,
  type FolioResolvedReviewedView,
} from "./ai-edits/headless";
import { createFolioAITextRangeHandle, trailingBodyBlockId } from "./ai-edits/snapshot";
import type {
  FolioAIBlock,
  FolioAIBlockParagraphProperties,
  FolioAIParagraphBlock,
  FolioAIEditAppliedOperation,
  FolioAIEditOperation,
  FolioAIEditSkippedOperation,
  FolioAIEditSnapshot,
} from "./ai-edits/types";
import { createScopedWordDiffOptions } from "./ai-edits/word-diff";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "./document-operations";
import { pairFolioDocumentStories } from "./document-stories";
import {
  resolveFolioDocumentPrivacyTransforms,
  rewriteDocxMetadataPrivacy,
  type FolioDocumentPrivacyOptions,
  type FolioDocumentPrivacyReport,
} from "./docx/metadataPrivacy";
import { FolioContentInlinePresentationProjectionError } from "./compare/content";
import { importStyleClosureWithNumbering } from "./compare/import-style-closure";
import {
  createInsertedResourceProjection,
  rebindInsertedResourceBlocks,
} from "./compare/inserted-resource-snapshot";
import { inlineFormattingSegments } from "./compare/formatting";
import {
  GenerateRedlineDocxOperationLimitError,
  MAX_GENERATED_REDLINE_OPERATIONS,
} from "./redlineOperationLimit";
import type { NumberingDefinitions } from "./types/document";
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

/** Raised when a resolved input view is not `"original"` or `"final"`. */
export class InvalidGenerateRedlineDocxOptionsError extends TaggedError(
  "InvalidGenerateRedlineDocxOptionsError",
)<{
  message: string;
  option: "baseView" | "revisedView";
  receivedValue: unknown;
}> {}

/** A document story that could not be paired across the two input packages. */
export type GenerateRedlineUnprocessedStory = {
  /** Story in the base package, or `null` when it exists only in the revision. */
  baseStory: FolioDocumentStoryHandle | null;
  /** Story in the revised package, or `null` when it exists only in the base. */
  revisedStory: FolioDocumentStoryHandle | null;
  /** Why the story could not be represented in the generated redline. */
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

type BuildFormattingRedlineOperationsOptions = {
  baseBlock: FolioAIBlock;
  revisedBlock: FolioAIBlock;
  nextOperationId: () => string;
};

const buildFormattingRedlineOperations = ({
  baseBlock,
  revisedBlock,
  nextOperationId,
}: BuildFormattingRedlineOperationsOptions): FolioAIEditOperation[] => {
  const formattingComparison = inlineFormattingSegments({
    baseBlock,
    targetBlock: revisedBlock,
    maxSegments: MAX_GENERATED_REDLINE_OPERATIONS,
  });
  const segments = (() => {
    switch (formattingComparison.status) {
      case "compared":
        return formattingComparison.segments;
      case "budget-exceeded":
        throw new GenerateRedlineDocxOperationLimitError({
          message: "The document comparison exceeds the generated operation limit.",
        });
      case "unalignable":
        throw new FolioContentInlinePresentationProjectionError({
          message: "The inline formatting runs could not be aligned for redline generation.",
          side: formattingComparison.side,
          baseBlockId: baseBlock.id,
          revisedBlockId: revisedBlock.id,
        });
      default: {
        const unreachable: never = formattingComparison;
        return panic("Unhandled inline formatting comparison result", { result: unreachable });
      }
    }
  })();
  const operations: FolioAIEditOperation[] = [];
  for (const { startOffset, endOffset, formatting } of segments) {
    const range = createFolioAITextRangeHandle({
      blockId: baseBlock.id,
      text: baseBlock.text,
      startOffset,
      endOffset,
    });
    if (!range) {
      panic("An aligned formatting range could not be represented");
    }
    operations.push({
      id: nextOperationId(),
      type: "formatRange",
      range,
      formatting,
    });
  }
  return operations;
};

const buildRedlineOperations = ({
  baseSnapshot,
  revisedBlocks,
  nextOperationId,
}: BuildRedlineOperationsOptions): FolioAIEditOperation[] => {
  const events = alignFolioBlocks(baseSnapshot.blocks, revisedBlocks);
  const anchorIds = nextBaseBlockIdByIndex(events);
  const operations: FolioAIEditOperation[] = [];
  const trailingAdditions: FolioAIBlock[] = [];
  // The last BODY-LEVEL paragraph, which the format guarantees exists: a table
  // may not be the last child of a body. Anchoring to the last block put the
  // anchor inside a table whenever the story ended with one, and an insertion
  // anchored there escapes to the table's boundary with no mark able to
  // express the break it added.
  const lastBaseBlockId = trailingBodyBlockId(baseSnapshot);
  const baseBlocksById = new Map(baseSnapshot.blocks.map((block) => [block.id, block]));

  events.forEach((event, eventIndex) => {
    if (event.type === "pair") {
      if (event.baseBlock.text !== event.revisedBlock.text) {
        operations.push({
          id: nextOperationId(),
          type: "replaceBlock",
          blockId: event.baseBlock.id,
          text: event.revisedBlock.text,
        });
      } else {
        operations.push(
          ...buildFormattingRedlineOperations({
            baseBlock: event.baseBlock,
            revisedBlock: event.revisedBlock,
            nextOperationId,
          }),
        );
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
      trailingAdditions.push(event.block);
      return;
    }
    operations.push({
      id: nextOperationId(),
      type: "insertBeforeBlock",
      blockId: anchorId,
      text: event.block.text,
      ...insertedParagraphProperties(event.block, baseBlocksById.get(anchorId)),
    });
  });

  // An empty document is one blank paragraph, and the alignment pairs it with
  // the first addition like any other block: it is replaced, and the rest
  // follow it. There is no hidden anchor to special-case any more.
  for (const addition of trailingAdditions) {
    operations.push({
      id: nextOperationId(),
      type: "insertAfterBlock",
      blockId: lastBaseBlockId ?? "redline-unanchored",
      text: addition.text,
      ...insertedParagraphProperties(
        addition,
        lastBaseBlockId === null ? undefined : baseBlocksById.get(lastBaseBlockId),
      ),
    });
  }

  return operations;
};

/**
 * The revised value, or `null` to clear one the anchor would pass on, or
 * `undefined` when neither has one: an explicit `null` costs a restyle pass.
 */
const statedOrCleared = <Value>(
  revised: Value | undefined,
  anchor: Value | undefined,
): Value | null | undefined => revised ?? (anchor === undefined ? undefined : null);

/**
 * The paragraph properties an inserted block states. An insertion that says
 * nothing takes the properties of its anchor, which is whichever base block
 * happens to follow it, so whatever the anchor states and the revised block
 * does not is cleared: a list item keeps its numbering, and a plain paragraph
 * beside a list item stays plain.
 */
const insertedParagraphProperties = (
  block: FolioAIBlock,
  anchor: FolioAIBlock | undefined,
): FolioAIBlockParagraphProperties => {
  const paragraph = paragraphBlockOf(block);
  if (paragraph === undefined) {
    return {};
  }
  const anchorParagraph = anchor === undefined ? undefined : paragraphBlockOf(anchor);
  const properties: FolioAIBlockParagraphProperties = {};
  const styleId = statedOrCleared(paragraph.styleId, anchorParagraph?.styleId);
  if (styleId !== undefined) properties.styleId = styleId;
  const alignment = statedOrCleared(paragraph.directAlignment, anchorParagraph?.directAlignment);
  if (alignment !== undefined) properties.alignment = alignment;
  const spacing = statedOrCleared(paragraph.directSpacing, anchorParagraph?.directSpacing);
  if (spacing !== undefined) properties.spacing = spacing;
  const indentation = statedOrCleared(
    paragraph.directIndentation,
    anchorParagraph?.directIndentation,
  );
  if (indentation !== undefined) properties.indentation = indentation;
  properties.numbering = paragraph.statedNumbering;
  return properties;
};

const paragraphBlockOf = (block: FolioAIBlock): FolioAIParagraphBlock | undefined => {
  switch (block.kind) {
    case "paragraph":
    case "heading":
    case "listItem":
      return block;
    case "diagnostic":
      return undefined;
    default: {
      const unreachable: never = block;
      return panic("Unhandled redline block kind", { block: unreachable });
    }
  }
};

/** The `w:numId`s whose `w:num` and abstract definition both exist. */
const definedNumIds = (numbering: NumberingDefinitions | null | undefined): Set<number> => {
  const abstractIds = new Set(numbering?.abstractNums.map((entry) => entry.abstractNumId));
  return new Set(
    (numbering?.nums ?? [])
      .filter((entry) => abstractIds.has(entry.abstractNumId))
      .map((entry) => entry.numId),
  );
};

/** Unresolvable source references have no visible list; retain only importable resources. */
const insertedNumberingReferences = (
  revisedReviewer: FolioDocxReviewer,
  snapshots: readonly FolioAIEditSnapshot[],
) => {
  const revisedNumbering = getFolioDocxComparisonAccess(revisedReviewer).numberingDefinitions();
  const resolvable = definedNumIds(revisedNumbering);
  return snapshots.flatMap((snapshot) =>
    snapshot.blocks.flatMap((block) => {
      const paragraph = paragraphBlockOf(block);
      const reference = paragraph?.listReference;
      return reference && resolvable.has(reference.numId) ? [reference] : [];
    }),
  );
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
  const wordDiff = createScopedWordDiffOptions({});
  let operationSequence = 0;
  const nextOperationId = () => {
    if (operationSequence >= MAX_GENERATED_REDLINE_OPERATIONS) {
      throw new GenerateRedlineDocxOperationLimitError({
        message: "The document comparison exceeds the generated operation limit.",
      });
    }
    return `redline-${++operationSequence}`;
  };

  const plannedStories: {
    story: FolioDocumentStoryHandle;
    snapshot: FolioAIEditSnapshot;
    revisedSnapshot: FolioAIEditSnapshot;
    inserted: ReturnType<typeof createInsertedResourceProjection>;
  }[] = [];
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
    plannedStories.push({
      story: pair.baseStory,
      snapshot: baseSnapshot,
      revisedSnapshot,
      inserted: createInsertedResourceProjection({ base: baseSnapshot, revised: revisedSnapshot }),
    });
  }

  const insertedSnapshots = plannedStories.map(({ inserted }) => inserted.snapshot);
  const resources = importStyleClosureWithNumbering({
    destination: baseReviewer,
    source: revisedReviewer,
    snapshots: insertedSnapshots,
    importedHeaderFooterSnapshots: insertedSnapshots,
    numberingReferences: insertedNumberingReferences(revisedReviewer, insertedSnapshots),
  });
  const { styleImport } = resources;
  const defined = definedNumIds(getFolioDocxComparisonAccess(baseReviewer).numberingDefinitions());
  for (const [index, { story, snapshot, revisedSnapshot, inserted }] of plannedStories.entries()) {
    const rebound =
      resources.snapshots.at(index) ?? panic("A redline resource import lost its snapshot");
    const insertedById = rebindInsertedResourceBlocks(inserted, rebound);
    const operations = buildRedlineOperations({
      baseSnapshot: snapshot,
      revisedBlocks: revisedSnapshot.blocks.map((block) => insertedById.get(block.id) ?? block),
      nextOperationId,
    })
      .filter((operation) => {
        if (operation.type !== "insertBeforeBlock" && operation.type !== "insertAfterBlock")
          return true;
        if (styleImport.status !== "unalignable") return true;
        skipped.push({ id: operation.id, reason: "missingStyle", message: styleImport.detail });
        return false;
      })
      .map((operation) => {
        if (operation.type !== "insertBeforeBlock" && operation.type !== "insertAfterBlock")
          return operation;
        if (operation.numbering?.kind !== "reference" || defined.has(operation.numbering.numId))
          return operation;
        return Object.assign(operation, { numbering: { kind: "none" as const } });
      });
    if (operations.length === 0) continue;
    const result = baseReviewer.applyDocumentOperationsToStory({
      story,
      snapshot,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations,
      },
      wordDiff,
      // Added paragraphs use the style closure imported through the collision owner.
      // Unimportable closures are reported as typed skips before applying the paragraph.
      undefinedStyles: "refuse",
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
