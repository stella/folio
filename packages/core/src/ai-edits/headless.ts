import JSZip from "jszip";
import { rebindDrawingImageRelationship } from "../docx/drawingRelationships";
import {
  captureSectionReferenceInventory,
  resolvedSectionReferenceLosses,
  withSectionReferenceResolution,
} from "../internal/sectionReferenceResolution";
import type { RemovedSectionReference } from "../internal/sectionEndpointResolution";
import {
  importReferencedStyleDefinitions,
  type ImportReferencedStyleDefinitionsResult,
} from "../compare/style-resources";
import { expectCharacterStyleMarkAttrs } from "../prosemirror/attrs";
/**
 * Headless `.docx` review path: buffer -> apply AI edits -> buffer, with
 * no `EditorView` and no DOM. A queue worker or agent can read a document,
 * apply `FolioAIEditOperation`s (tracked-changes or direct), and write a
 * reviewed `.docx` back out — the server-side counterpart to the React
 * editor's live apply flow.
 *
 * The operation applier is shared, not forked: {@link applyFolioAIEditOperations}
 * only needs a {@link FolioAIEditView} seam (`{ state, dispatch }`), which a
 * headless `EditorState` satisfies via `state.apply(tr)`. The React editor and
 * this reviewer therefore run byte-for-byte the same anchor resolution,
 * word-diff redlines, and tracked-change bookkeeping.
 *
 * Save mirrors the editor's own path: a selective patch of only the changed
 * paragraphs in `document.xml` (leaving every untouched part byte-exact), with
 * a full repack as the fallback for structural edits.
 *
 * Scope: main, header, footer, footnote, and endnote blocks.
 */

import { sectionReferenceHistory } from "../docx/sectionReferenceHistory";
import { sectionRejectProperties } from "../prosemirror/commands/propertyChangeScope";
import { panic, TaggedError } from "better-result";
import { canonicalJson } from "../utils/canonicalJson";
import {
  matchInlineProvenance,
  type InlineProvenanceTargetOptions,
} from "../compare/inline-provenance";
import { matchInlineAtoms, type MatchInlineAtomsOptions } from "../compare/inline-atoms";
import { stageSectionBoundaryProperties as stageMappedSectionBoundaryProperties } from "../compare/section-boundary-properties";
import { Fragment } from "prosemirror-model";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import type { Command } from "prosemirror-state";
import type { Plugin } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";

import { createReply } from "../docx/replyToComment";
import { attemptSelectiveSave } from "../docx/selectiveSave";
import { getHeaderFooterText } from "../docx/headerFooterParser";
import {
  getEndnoteText,
  getFootnoteText,
  isSeparatorEndnote,
  isSeparatorFootnote,
} from "../docx/footnoteParser";
import { parseDocx } from "../docx/parser";
import { repackDocx } from "../docx/rezip";
import { pluginsForHeadlessRevisionResolution } from "../internal/headlessRevisionResolutionGuard";
import {
  type TrackedSectionEndpointRemoval,
  withTrackedSectionEndpointRemoval,
} from "../internal/sectionEndpointResolution";
import {
  acceptAIEditRevision,
  rejectAIEditRevision,
  resolveAllChangesInHeadlessState,
} from "../prosemirror/commands/comments";
import { proseDocToBlocks, updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import {
  footnoteToProseDoc,
  headerFooterToProseDoc,
  toProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import { ensureBaseDirectionInState } from "../prosemirror/extensions/features/AutoBidiDetectionExtension";
import {
  getChangedParagraphIds,
  getTrackedSectionEndpointRemoval,
  hasStructuralChanges,
  hasUntrackedChanges,
} from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import {
  createDocumentStylesPlugin,
  withDocumentStyles,
  getDocumentStyleResolver,
} from "../prosemirror/plugins/documentStyles";
import {
  createDocumentNumberingPlugin,
  withDocumentNumbering,
} from "../prosemirror/plugins/documentNumbering";
import { schema, singletonManager } from "../prosemirror/schema";
import { REVIEW_CARRIERS } from "@stll/docx-core/model";
import { MAX_LIST_LEVEL } from "../prosemirror/listMarker";
import type { Comment } from "../types/content";
import type {
  Document,
  BlockContent,
  Endnote,
  Footnote,
  HeaderFooter,
  MediaFile,
  NumberingDefinitions,
  SectionProperties,
  StyleDefinitions,
} from "../types/document";
import { deterministicHexId } from "../utils/hexId";
import { getCachedNumberingMap } from "../docx/numberingParser";
import {
  recreateProseNodeWithParagraphPropertySource,
  transferProseParagraphPropertySource,
} from "../docx/paragraphPropertySource";
import {
  applyFolioDocumentOperations,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationUndoHandle,
  type FolioDocumentOperationUndoResult,
} from "../document-operations";
import type { FolioRevisionStamp, FolioWordDiffOptions } from "./apply";
import { buildAnnotatedBlockText } from "./clean-text";
import {
  getCommentAnchorsFromDoc,
  getTrackedChangeStatsFromDoc,
  getTrackedChangesFromDoc,
  getTrackedChangesFromSnapshot,
  type FolioReviewChange,
  type FolioReviewChangeKind,
} from "./read";
import {
  createFolioAIEditSnapshotWithStyleResolver,
  detachFolioAIEditSnapshotExternalHyperlinks,
  folioStoryTables,
  isFolioAIContentBlock,
  normalizeFolioAIBlockText,
  sourceDocumentOf,
  type FolioStoryTable,
} from "./snapshot";
import { matchTableGeometry, type TableGeometryPairing } from "./table-geometry";
import { tableTemplateCanCrossPackageLosslessly, type FolioTableTemplates } from "./table-template";
import type {
  FolioAIBlock,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "./types";

/**
 * Standalone comment-id allocator for the reviewer. The React editor uses
 * `Date.now()`-seeded ids from `commentsHelpers`; that module is DOM-bound
 * (`EditorView`, `findBodyPmAnchors`) and unimportable here, so the reviewer
 * mints its own. Seeded once per realm and incremented so ids stay unique
 * across reviewers created within the same millisecond.
 */
let commentIdCursor = Date.now();
let undoHandleCursor = Date.now();

/**
 * Build the note-free comment thread the apply layer references by id.
 * Mirrors `commentsHelpers.createComment` (a pure object literal there) so a
 * headless `commentOnBlock` / `comment` op serialises the same `comments.xml`
 * shape the editor produces.
 */
const createReviewerComment = (id: number, text: string, author: string): Comment => ({
  id,
  author,
  date: new Date().toISOString(),
  content: [
    {
      type: "paragraph",
      formatting: {},
      content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
    },
  ],
});

type FolioDocumentOperationUndoEntry = {
  undoHandle: FolioDocumentOperationUndoHandle;
  story: FolioEditableDocumentStoryHandle;
  beforeState: EditorState;
  afterState: EditorState;
  createdCommentsLengthBefore: number;
  createdCommentsLengthAfter: number;
};

/**
 * Assign a stable `w14:paraId` to every body paragraph that lacks one (or
 * duplicates an earlier id), derived deterministically from the paragraph's
 * text plus its document ordinal. Re-parsing the SAME bytes mints the SAME ids,
 * so the block ids a snapshot exposes are reproducible across
 * {@link FolioDocxReviewer.fromBuffer} calls — the documented "snapshot on one
 * parse, apply on another" flow resolves instead of skipping every op as a
 * stale anchor. Word-authored paraIds are preserved; only gaps and collisions
 * get a fresh id.
 *
 * The shared `ParaIdAllocatorExtension` mints RANDOM ids (correct for freshly
 * typed paragraphs in the live editor); this load-time pass is deterministic so
 * a paraId-less corpus document anchors reproducibly.
 *
 * Rebuilding only the changed branches, like {@link ensureParaIdsInDoc}, keeps
 * the pass linear in paragraph count. Seeding through a transaction instead
 * costs one `setNodeMarkup` step per paragraph, and both halves of that are
 * quadratic: every step rebuilds the containing fragment, and the plugin
 * `appendTransaction` chain rescans the accumulated step maps. Because the pass
 * runs before the state exists, no history, mapping, or change-tracking
 * semantics depend on it.
 */
const ensureDeterministicParaIdsInDoc = (doc: PMNode): PMNode => {
  const seen = new Set<string>();
  let ordinal = 0;

  const rewrite = (parent: PMNode): Fragment => {
    let changed = false;
    const children: PMNode[] = [];

    parent.forEach((child) => {
      let next = child;
      if (child.type.name === "paragraph") {
        ordinal += 1;
        const existing = child.attrs["paraId"];
        if (typeof existing === "string" && existing.length > 0 && !seen.has(existing)) {
          seen.add(existing);
        } else {
          let paraId = deterministicHexId(`${child.textContent}:${ordinal}`);
          for (let salt = 1; seen.has(paraId); salt++) {
            paraId = deterministicHexId(`${child.textContent}:${ordinal}:${salt}`);
          }
          seen.add(paraId);
          next = recreateProseNodeWithParagraphPropertySource(child, {
            attrs: { ...child.attrs, paraId, idStability: "positional" },
          });
        }
        const paraId = next.attrs["paraId"];
        if (typeof paraId === "string") {
          transferProseParagraphPropertySource(next, child, paraId);
        }
      } else if (child.childCount > 0) {
        const content = rewrite(child);
        if (content !== child.content) {
          next = recreateProseNodeWithParagraphPropertySource(child, { content });
        }
      }
      if (next !== child) {
        changed = true;
      }
      children.push(next);
    });

    return changed ? Fragment.fromArray(children) : parent.content;
  };

  const content = rewrite(doc);
  return content === doc.content
    ? doc
    : recreateProseNodeWithParagraphPropertySource(doc, { content });
};

/** Options for {@link FolioDocxReviewer.fromBuffer}. */
export type FolioDocxReviewerOptions = {
  /** Default author for tracked changes and comments. (default: `"AI"`) */
  author?: string;
  /** Password for Agile-encrypted .docx files (Office 2010+). */
  password?: string | undefined;
};

/** Options for {@link FolioDocxReviewer.applyOperations}. */
export type FolioApplyOperationsOptions = {
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
  /**
   * The snapshot the `operations`' block ids were built against. Omit to
   * snapshot the reviewer's current state — correct for the common flow of
   * `snapshot()` -> build ops -> `applyOperations()` on the same reviewer.
   */
  snapshot?: FolioAIEditSnapshot;
  /** Omit to stamp revisions from the wall clock and the shared id cursor. */
  revisionStamp?: FolioRevisionStamp;
  /** Token size a replacement's redline is cut at. Word-level by default. */
  wordDiff?: FolioWordDiffOptions;
};

/** Options for {@link FolioDocxReviewer.applyDocumentOperations}. */
export type FolioApplyDocumentOperationsOptions = Omit<FolioApplyOperationsOptions, "mode">;

/** Options for {@link FolioDocxReviewer.getContentAsText}. */
export type FolioGetContentAsTextOptions = {
  /**
   * Render tracked changes and comment anchors inline as `<ins>` / `<del>` /
   * `<comment>` tags instead of the default flattened, post-tracked-changes
   * text. (default: `false`)
   */
  annotated?: boolean;
};

export const FOLIO_REVIEWED_VIEWS = Object.freeze(["original", "current-markup", "final"] as const);

export type FolioReviewedView = (typeof FOLIO_REVIEWED_VIEWS)[number];

export const FOLIO_RESOLVED_REVIEWED_VIEWS = Object.freeze(["original", "final"] as const);

export type FolioResolvedReviewedView = (typeof FOLIO_RESOLVED_REVIEWED_VIEWS)[number];

export type FolioDocumentStoryHandle =
  | { type: "main" }
  | { type: "header"; relationshipId: string }
  | { type: "footer"; relationshipId: string }
  | { type: "footnote"; noteId: number }
  | { type: "endnote"; noteId: number };

export type FolioEditableDocumentStoryHandle = FolioDocumentStoryHandle;

/**
 * One numbering level as a reader meets it: which list, which depth, and the
 * format and template that produce the label.
 */
export type FolioNumberingLevel = {
  numId: number;
  level: number;
  format: string;
  levelText: string;
  start?: number;
};

export type FolioDocumentStory = {
  handle: FolioDocumentStoryHandle;
  text: string;
};

export type FolioReadReviewedStoryOptions = {
  story?: FolioEditableDocumentStoryHandle;
  view?: FolioReviewedView;
};

/** What {@link FolioDocxReviewer.matchStoryTableGeometry} needs to move a table's properties. */
export type FolioMatchStoryTableGeometryOptions = {
  story?: FolioEditableDocumentStoryHandle;
  /** The other document's tables, by the index its snapshot numbers them with. */
  targetTables: ReadonlyMap<number, PMNode>;
  /** Base cells and the target cells they were aligned with. */
  pairings: readonly TableGeometryPairing[];
  revisionStamp: FolioRevisionStamp;
};

export type FolioResolveReviewedStoryOptions = {
  story?: FolioEditableDocumentStoryHandle;
  view: FolioResolvedReviewedView;
};

export type FolioReviewedStory = {
  story: FolioEditableDocumentStoryHandle;
  view: FolioReviewedView;
  snapshot: FolioAIEditSnapshot;
  text: string;
  changes: FolioReviewChange[];
};

export class UnsupportedFolioReviewedViewError extends TaggedError(
  "UnsupportedFolioReviewedViewError",
)<{
  message: string;
  receivedView: unknown;
}> {}

export class FolioDocumentStoryNotFoundError extends TaggedError(
  "FolioDocumentStoryNotFoundError",
)<{
  message: string;
  story: FolioEditableDocumentStoryHandle;
}> {}

const FOLIO_RESOLVED_STORY_SERIALIZATION_MISMATCHES = Object.freeze({
  storyMissing: "story-missing",
  revisionMarkupRemains: "revision-markup-remains",
  textProjection: "text-projection",
  blockProjection: "block-projection",
} as const);

type FolioResolvedStorySerializationMismatch =
  (typeof FOLIO_RESOLVED_STORY_SERIALIZATION_MISMATCHES)[keyof typeof FOLIO_RESOLVED_STORY_SERIALIZATION_MISMATCHES];

class FolioResolvedStorySerializationError extends TaggedError(
  "FolioResolvedStorySerializationError",
)<{
  message: string;
  story: FolioEditableDocumentStoryHandle;
  mismatches: readonly FolioResolvedStorySerializationMismatch[];
  expectedBlockCount: number;
  actualBlockCount: number | null;
  remainingChangeCount: number | null;
}> {}

export type FolioApplyDocumentOperationsToStoryOptions = FolioApplyDocumentOperationsOptions & {
  story: FolioEditableDocumentStoryHandle;
  batch: FolioDocumentOperationBatch;
  /**
   * Tables and rows an `insertTable` / `insertTableRow` in the batch places
   * verbatim, by operation id. A caller copying a table it already holds —
   * `compareDocx` placing the target document's table — hands it over whole
   * instead of letting the operation rebuild it from its cell texts and lose
   * every table, row and cell property on the way.
   */
  tableTemplates?: FolioTableTemplates;
};

export type { FolioRevisionStamp };

type FolioHeaderFooterStoryHandle = Extract<
  FolioEditableDocumentStoryHandle,
  { type: "header" | "footer" }
>;

type FolioNoteStoryHandle = Extract<
  FolioEditableDocumentStoryHandle,
  { type: "footnote" | "endnote" }
>;

type FolioSecondaryStoryHandle = Exclude<FolioEditableDocumentStoryHandle, { type: "main" }>;

type FolioSecondaryStoryState = {
  handle: FolioSecondaryStoryHandle;
  initialState: EditorState;
  state: EditorState;
};

type FolioResolvedStoryBlock = Omit<FolioAIBlock, "idStability">;

/**
 * `idStability` records how an id entered the current snapshot. A synthesized
 * paraId is positional until save writes it into the package, then becomes an
 * authored stable id when reopened. Compare the persisted block projection
 * without changing that live snapshot identity contract.
 */
const resolvedStoryBlockProjection = (block: FolioAIBlock): FolioResolvedStoryBlock => {
  const persisted = { ...block };
  delete persisted.idStability;
  return persisted;
};

type FolioResolvedStoryExpectation = {
  story: FolioEditableDocumentStoryHandle;
  text: string;
  blocks: readonly FolioResolvedStoryBlock[];
};

type FolioReviewerStateSnapshot = {
  mainState: EditorState;
  finalSectionPropertiesOverride: SectionProperties | undefined;
  secondaryStoryStates: readonly FolioSecondaryStoryState[];
  sectionReferenceRemovals: readonly RemovedSectionReference[];
  removedHeaderFooterStories: readonly FolioHeaderFooterStoryHandle[];
  importedStyles: StyleDefinitions | undefined;
  importedMedia: ReadonlyMap<string, MediaFile>;
  importedHeaders: ReadonlyMap<string, HeaderFooter>;
  importedFooters: ReadonlyMap<string, HeaderFooter>;
  createdComments: readonly Comment[];
  resolvedOverrides: ReadonlyMap<number, boolean>;
  resolvedStoryExpectations: readonly FolioResolvedStoryExpectation[];
};

type FolioSavePath =
  | { type: "full-repack" }
  | { type: "selective-first"; changedParaIds: Set<string> };

type FolioSaveSnapshot = {
  document: Document;
  path: FolioSavePath;
  changedNoteParaIds: ReadonlySet<string>;
  sectionReferenceRemovals: readonly RemovedSectionReference[];
  sectionEndpointRemoval: TrackedSectionEndpointRemoval | null;
  resolvedStoryExpectations: readonly FolioResolvedStoryExpectation[];
};

type ApplyDocumentOperationsInternalOptions = {
  story: FolioEditableDocumentStoryHandle;
  batch: FolioDocumentOperationBatch;
  snapshot?: FolioAIEditSnapshot;
  revisionStamp?: FolioRevisionStamp;
  wordDiff?: FolioWordDiffOptions;
  tableTemplates?: FolioTableTemplates;
  createUndoEntry: boolean;
};

type FolioDocxComparisonProjectionMode = "with-revision-census" | "without-revision-census";

type FolioDocxComparisonStoryProjection = {
  handle: FolioDocumentStoryHandle;
  snapshot: FolioAIEditSnapshot | null;
};

type FolioDocxComparisonProjection = {
  stories: readonly FolioDocxComparisonStoryProjection[];
  revisions: {
    highestId: number;
    present: boolean;
  };
};

type MatchStoryInlineProvenanceOptions = InlineProvenanceTargetOptions & {
  story: FolioEditableDocumentStoryHandle;
};

type StoryInlineProvenanceResult =
  | {
      status: "matched";
      nextRevisionId: number;
      changedTargetBlockIds: readonly string[];
      rangeCount: number;
      documentChanged: boolean;
    }
  | { status: "unalignable" }
  | { status: "budget-exceeded" };

type MatchStoryInlineAtomsOptions = Omit<MatchInlineAtomsOptions, "state" | "author"> & {
  story: FolioEditableDocumentStoryHandle;
};

type StoryInlineAtomsResult =
  | {
      status: "matched";
      nextRevisionId: number;
      changedTargetBlockIds: readonly string[];
      rangeCount: number;
      documentChanged: boolean;
    }
  | { status: "unalignable" }
  | { status: "budget-exceeded" };

type StageTargetNumberingResult = "unchanged" | "staged" | "conflict";

const numberingLevelsOf = (
  numbering: NumberingDefinitions | null | undefined,
): FolioNumberingLevel[] => {
  if (!numbering) return [];
  const levelsByAbstractId = new Map(
    numbering.abstractNums.map((abstractNum) => [abstractNum.abstractNumId, abstractNum.levels]),
  );
  const levels: FolioNumberingLevel[] = [];
  for (const instance of numbering.nums) {
    const overrides = new Map(
      (instance.levelOverrides ?? []).map((override) => [override.ilvl, override]),
    );
    for (const level of levelsByAbstractId.get(instance.abstractNumId) ?? []) {
      const override = overrides.get(level.ilvl);
      const resolved = override?.lvl ?? level;
      levels.push({
        numId: instance.numId,
        level: resolved.ilvl,
        format: resolved.numFmt,
        levelText: resolved.lvlText,
        ...((override?.startOverride ?? resolved.start) !== undefined
          ? { start: override?.startOverride ?? resolved.start }
          : {}),
      });
    }
  }
  return levels;
};

type ReferencedNumberingLevelsOptions = {
  references: readonly { numId: number; level: number }[];
};

const referencedNumberingLevelsByNumId = ({
  references,
}: ReferencedNumberingLevelsOptions): ReadonlyMap<number, readonly number[]> => {
  const levelsByNumId = new Map<number, Set<number>>();
  for (const { numId, level } of references) {
    const levels = levelsByNumId.get(numId) ?? new Set<number>();
    levelsByNumId.set(numId, levels);
    levels.add(level);
    for (let ancestor = 0; ancestor <= Math.min(level, MAX_LIST_LEVEL); ancestor += 1) {
      levels.add(ancestor);
    }
  }
  return new Map(
    [...levelsByNumId].map(([numId, levels]) => [
      numId,
      [...levels].toSorted((left, right) => left - right),
    ]),
  );
};

type SameReferencedNumberingLevelsOptions = {
  current: NumberingDefinitions | null | undefined;
  target: NumberingDefinitions;
  numId: number;
  levelsByNumId: ReadonlyMap<number, readonly number[]>;
};

const sameReferencedNumberingLevels = ({
  current,
  target,
  numId,
  levelsByNumId,
}: SameReferencedNumberingLevelsOptions): boolean => {
  if (!current) return false;
  const currentNumbering = getCachedNumberingMap(current);
  const targetNumbering = getCachedNumberingMap(target);
  for (const level of levelsByNumId.get(numId) ?? []) {
    const currentLevel = currentNumbering.getLevel(numId, level);
    const targetLevel = targetNumbering.getLevel(numId, level);
    if (
      !currentLevel ||
      !targetLevel ||
      canonicalJson(currentLevel) !== canonicalJson(targetLevel)
    ) {
      return false;
    }
  }
  return true;
};

type FolioDocxComparisonAccess = {
  stageTerminalTableReviewCarrier: (target: PMNode) => boolean;
  stageTargetStyles: (
    source: FolioDocxReviewer,
    snapshots: readonly FolioAIEditSnapshot[],
    importedHeaderFooterSnapshots: readonly FolioAIEditSnapshot[],
  ) => ImportReferencedStyleDefinitionsResult;
  createComparisonHeaderFooter: (
    source: FolioDocxReviewer,
    story: FolioHeaderFooterStoryHandle,
  ) => Promise<FolioHeaderFooterStoryHandle | null>;
  matchInlineAtoms: (options: MatchStoryInlineAtomsOptions) => StoryInlineAtomsResult;
  matchInlineProvenance: (
    options: MatchStoryInlineProvenanceOptions,
  ) => StoryInlineProvenanceResult;
  projectStories: (mode: FolioDocxComparisonProjectionMode) => FolioDocxComparisonProjection;
  snapshotReviewedStory: (options?: FolioReadReviewedStoryOptions) => FolioAIEditSnapshot | null;
  numberingDefinitions: () => NumberingDefinitions | null | undefined;
  planTargetNumberingReferences: (
    target: NumberingDefinitions | null | undefined,
    references: readonly { numId: number; level: number }[],
  ) => ReadonlyMap<number, number> | null;
  stageTargetNumbering: (
    target: NumberingDefinitions | null | undefined,
    references: readonly { numId: number; level: number }[],
    remappedNumIds: ReadonlyMap<number, number>,
  ) => StageTargetNumberingResult;
  finalSectionProperties: () => SectionProperties | undefined;
  stageFinalSectionProperties: (options: {
    target: SectionProperties;
    previous: SectionProperties;
    revision: FolioRevisionStamp;
  }) => boolean;
  stageMappedSectionBoundaries: (options: {
    target: PMNode;
    originalRevisionIdSeed: number;
    maxRanges: number;
    revision: FolioRevisionStamp;
    mapTargetProperties: (args: {
      kind: "inserted" | "retained";
      current: SectionProperties | undefined;
      target: SectionProperties;
    }) =>
      | { kind: "inserted"; target: SectionProperties }
      | { kind: "retained"; previous: SectionProperties; target: SectionProperties }
      | null;
  }) =>
    | { status: "matched"; rangeCount: number; nextRevisionId: number; documentChanged: boolean }
    | { status: "unalignable"; detail: string }
    | { status: "budget-exceeded" };
};

const comparisonAccessByReviewer = new WeakMap<FolioDocxReviewer, FolioDocxComparisonAccess>();

const MAIN_STORY = Object.freeze({ type: "main" } as const);

const headerFooterStoryKey = ({ type, relationshipId }: FolioHeaderFooterStoryHandle): string =>
  `${type}:${relationshipId}`;

const noteStoryKey = ({ type, noteId }: FolioNoteStoryHandle): string => `${type}:${noteId}`;

const secondaryStoryKey = (story: FolioSecondaryStoryHandle): string =>
  story.type === "header" || story.type === "footer"
    ? headerFooterStoryKey(story)
    : noteStoryKey(story);

const editableStoryKey = (story: FolioEditableDocumentStoryHandle): string =>
  story.type === "main" ? "main" : secondaryStoryKey(story);

export const isFolioReviewedView = (value: unknown): value is FolioReviewedView =>
  FOLIO_REVIEWED_VIEWS.some((view) => view === value);

export const isFolioResolvedReviewedView = (value: unknown): value is FolioResolvedReviewedView =>
  FOLIO_RESOLVED_REVIEWED_VIEWS.some((view) => view === value);

const resolveReviewedState = (state: EditorState, view: FolioReviewedView): EditorState => {
  if (view === "current-markup") {
    return state;
  }
  return resolveAllChangesInHeadlessState(state, view === "original" ? "reject" : "accept");
};

const createHeadlessPlugins = (
  styles: Document["package"]["styles"],
  numbering: NumberingDefinitions | null | undefined,
): Plugin[] => [
  ...pluginsForHeadlessRevisionResolution(singletonManager.getPlugins()),
  createDocumentStylesPlugin(styles),
  createDocumentNumberingPlugin(numbering),
];

const createStateSnapshot = (state: EditorState): FolioAIEditSnapshot =>
  createFolioAIEditSnapshotWithStyleResolver(state.doc, getDocumentStyleResolver(state));

const formatStorySnapshotForLLM = (snapshot: FolioAIEditSnapshot, annotated: boolean): string => {
  // A reading surface shows content. The snapshot also carries the document's
  // blank paragraphs, which are structure rather than something to read.
  const blocks = snapshot.blocks.filter(isFolioAIContentBlock);
  if (!annotated) {
    return blocks.map(formatBlockForLLM).join("\n");
  }
  // One walk, not one `doc.nodeAt` per block: `nodeAt` re-descends from the
  // root and scans each level's fragment from index 0, so looking every block
  // up costs O(blocks^2) on a flat document.
  const nodeByStart = new Map<number, PMNode>();
  sourceDocumentOf(snapshot).descendants((node, pos) => {
    if (!node.isTextblock) {
      return true;
    }
    nodeByStart.set(pos, node);
    return false;
  });
  const startById = new Map<string, number>();
  for (const anchor of Object.values(snapshot.anchors)) {
    startById.set(anchor.id, anchor.from);
  }
  return blocks
    .map((block) => {
      const from = startById.get(block.id);
      const node = from === undefined ? undefined : nodeByStart.get(from);
      const text = node ? buildAnnotatedBlockText(node) : block.text;
      return formatBlockLine(block, text);
    })
    .join("\n");
};

// The change shape and its pure reader now live in `./read` so a live editor
// can produce the same `FolioReviewChange[]` from its own doc; the reviewer
// delegates to `getTrackedChangesFromDoc` and re-exports the types here so its
// public surface is unchanged.
export type { FolioReviewChange, FolioReviewChangeKind } from "./read";

/** Filter for {@link FolioDocxReviewer.getChanges}. */
export type FolioReviewChangeFilter = {
  author?: string;
  type?: FolioReviewChangeKind;
};

export type FolioReviewCommentReply = {
  id: number;
  author: string;
  date: string | null;
  text: string;
};

/** Input for {@link FolioDocxReviewer.replyTo}. */
export type FolioReviewReplyInput = {
  /** Reply body text. */
  text: string;
  /** Reply author; defaults to the reviewer's author. */
  author?: string;
  initials?: string;
};

/** A comment thread discovered in the document. */
export type FolioReviewComment = {
  id: number;
  author: string;
  date: string | null;
  /** The comment body text, its paragraphs joined by newlines. */
  text: string;
  /** The document text the comment is anchored to, or `""` when unanchored. */
  anchoredText: string;
  /** Stable id of the anchored body block, or `null` when the anchor is absent. */
  blockId: string | null;
  replies: FolioReviewCommentReply[];
  /** Whether the comment is marked resolved / done. */
  done: boolean;
};

/** Filter for {@link FolioDocxReviewer.getComments}. */
export type FolioReviewCommentFilter = {
  author?: string;
  done?: boolean;
};

/**
 * One LLM-ready line for a block: `[<blockId>] text`, with an `(h<level>)` tag
 * for headings and the list marker for list items, so a model can copy the
 * block id straight back into an operation.
 */
const formatBlockLine = (block: FolioAIBlock, text: string): string => {
  const label = `[${block.id}]`;
  if (block.kind === "heading") {
    return `${label} (h${headingLevel(block)}) ${text}`;
  }
  if (block.kind === "listItem") {
    return `${label} ${block.displayLabel ?? "•"} ${text}`;
  }
  return `${label} ${text}`;
};

const formatBlockForLLM = (block: FolioAIBlock): string => formatBlockLine(block, block.text);

const headingLevel = (block: FolioAIBlock): number => {
  const digits = /(\d+)/u.exec(block.styleId ?? block.displayLabel ?? "")?.[1];
  const level = digits ? Number.parseInt(digits, 10) : 1;
  return level >= 1 && level <= 9 ? level : 1;
};

const revisionIdOf = (target: FolioReviewChange | number): number =>
  typeof target === "number" ? target : target.id;

const commentPlainText = (comment: Comment): string =>
  // A comment parsed from a malformed package can omit `content`; the model
  // types it as required, so guard at this untrusted boundary.
  (comment.content ?? []).map(paragraphPlainText).join("\n");

const paragraphPlainText = (paragraph: Comment["content"][number]): string => {
  const parts: string[] = [];
  for (const item of paragraph.content ?? []) {
    if (item.type !== "run") {
      continue;
    }
    for (const runItem of item.content ?? []) {
      if (runItem.type === "text") {
        parts.push(runItem.text);
      }
    }
  }
  return parts.join("");
};

/**
 * Headless `.docx` reviewer. Parse a buffer, read blocks, apply
 * `FolioAIEditOperation`s against the document model, and write the reviewed
 * `.docx` back out — no editor instance, no DOM.
 *
 * @example
 * ```ts
 * const reviewer = await FolioDocxReviewer.fromBuffer(buffer, { author: "AI" });
 * const { blocks } = reviewer.snapshot();
 * reviewer.applyOperations([
 *   { id: "1", type: "replaceInBlock", blockId: blocks[0].id, find: "$50k", replace: "$500k" },
 * ]);
 * const reviewed = await reviewer.toBuffer();
 * ```
 */
export class FolioDocxReviewer {
  /** Default author for tracked changes and comments. */
  readonly author: string;
  private readonly baseDocument: Document;
  private finalSectionPropertiesOverride: SectionProperties | undefined;
  private readonly originalBuffer: ArrayBuffer;
  private state: EditorState;
  private readonly secondaryStoryStates = new Map<string, FolioSecondaryStoryState>();
  private readonly removedHeaderFooterStories = new Map<string, FolioHeaderFooterStoryHandle>();
  private readonly sectionReferenceRemovals: RemovedSectionReference[] = [];
  private importedStyles: StyleDefinitions | undefined;
  private readonly importedMedia = new Map<string, MediaFile>();
  private readonly importedHeaders = new Map<string, HeaderFooter>();
  private readonly importedFooters = new Map<string, HeaderFooter>();
  private readonly resolvedStoryExpectations = new Map<string, FolioResolvedStoryExpectation>();
  private readonly createdComments: Comment[] = [];
  private readonly usedCommentIds: Set<number>;
  private readonly documentOperationUndoEntries: FolioDocumentOperationUndoEntry[] = [];
  /**
   * Resolved-state overrides recorded by {@link resolveComment}, keyed by
   * comment id. Applied on read ({@link getComments}) and on write
   * ({@link toDocument}) rather than mutating the parsed `Comment` objects in
   * place, matching how the parser/serializer treat `Comment` as immutable
   * (`commentParser.ts` replaces array slots via spread; the serializer
   * types its inputs `readonly Comment[]`).
   */
  private readonly resolvedOverrides = new Map<number, boolean>();

  private constructor(args: {
    baseDocument: Document;
    originalBuffer: ArrayBuffer;
    state: EditorState;
    author: string;
  }) {
    this.baseDocument = args.baseDocument;
    this.originalBuffer = args.originalBuffer;
    this.state = args.state;
    this.author = args.author;
    this.usedCommentIds = new Set(
      (args.baseDocument.package.document.comments ?? []).map(({ id }) => id),
    );
    comparisonAccessByReviewer.set(
      this,
      Object.freeze({
        stageTerminalTableReviewCarrier: (target) => this.stageTerminalTableReviewCarrier(target),
        stageTargetStyles: (source, snapshots, importedHeaderFooterSnapshots) =>
          this.stageTargetStyles(source, snapshots, importedHeaderFooterSnapshots),
        createComparisonHeaderFooter: (source, story) =>
          this.createComparisonHeaderFooter(source, story),
        finalSectionProperties: () => this.currentFinalSectionProperties(),
        stageFinalSectionProperties: ({ target, previous, revision }) =>
          this.stageFinalSectionProperties({ target, previous, revision }),
        stageMappedSectionBoundaries: (options) => this.stageMappedSectionBoundaries(options),
        matchInlineAtoms: ({ story, ...options }) => {
          const state = this.getEditableStoryState(story);
          if (!state) return panic("A compared story lost its editable state", { story });
          const result = matchInlineAtoms({ ...options, state, author: this.author });
          switch (result.status) {
            case "matched":
              if (result.transaction.docChanged) {
                this.setEditableStoryState(story, state.apply(result.transaction));
              }
              return {
                status: "matched",
                nextRevisionId: result.nextRevisionId,
                changedTargetBlockIds: result.changedTargetBlockIds,
                rangeCount: result.rangeCount,
                documentChanged: result.transaction.docChanged,
              };
            case "unalignable":
            case "budget-exceeded":
              return result;
            default: {
              const unreachable: never = result;
              return panic("Unhandled inline atom result", { result: unreachable });
            }
          }
        },
        matchInlineProvenance: ({ story, ...options }) => {
          const state = this.getEditableStoryState(story);
          if (!state) return panic("A compared story lost its editable state", { story });
          const result = matchInlineProvenance({ ...options, state, author: this.author });
          switch (result.status) {
            case "matched":
              if (result.transaction.docChanged) {
                this.setEditableStoryState(story, state.apply(result.transaction));
              }
              return {
                status: "matched",
                nextRevisionId: result.nextRevisionId,
                changedTargetBlockIds: result.changedTargetBlockIds,
                rangeCount: result.rangeCount,
                documentChanged: result.transaction.docChanged,
              };
            case "unalignable":
            case "budget-exceeded":
              return result;
            default: {
              const unreachable: never = result;
              return panic("Unhandled inline provenance result", { result: unreachable });
            }
          }
        },
        projectStories: (mode) => this.projectComparisonStoriesInternal(mode),
        snapshotReviewedStory: (options) => this.snapshotReviewedStoryInternal(options),
        numberingDefinitions: () => this.baseDocument.package.numbering,
        planTargetNumberingReferences: (target, references) =>
          this.planTargetNumberingReferences(target, references),
        stageTargetNumbering: (target, references, remappedNumIds) =>
          this.stageTargetNumbering(target, references, remappedNumIds),
      }),
    );
  }

  private stageTargetStyles(
    source: FolioDocxReviewer,
    snapshots: readonly FolioAIEditSnapshot[],
    importedHeaderFooterSnapshots: readonly FolioAIEditSnapshot[],
  ): ImportReferencedStyleDefinitionsResult {
    const destination = this.baseDocument.package;
    const sourcePackage = source.baseDocument.package;
    const destinationStyles = this.importedStyles ?? destination.styles;
    const existing = new Set(destinationStyles?.styles.map(({ styleId }) => styleId));
    const sourceStyleIds = new Set(sourcePackage.styles?.styles.map(({ styleId }) => styleId));
    const collect = (document: PMNode): Set<string> => {
      const references = new Set<string>();
      document.descendants((node) => {
        const styleId = node.attrs["styleId"];
        if (typeof styleId === "string" && styleId.length > 0) references.add(styleId);
        for (const mark of node.marks) {
          if (mark.type.name === "characterStyle")
            references.add(expectCharacterStyleMarkAttrs(mark).styleId);
        }
      });
      return references;
    };
    const hasUnstyledParagraph = (document: PMNode): boolean => {
      let found = false;
      document.descendants((node) => {
        if (node.isTextblock && typeof node.attrs["styleId"] !== "string") {
          found = true;
          return false;
        }
        return true;
      });
      return found;
    };
    const contextsMatch =
      canonicalJson(sourcePackage.styles?.docDefaults) ===
        canonicalJson(destinationStyles?.docDefaults) &&
      canonicalJson(sourcePackage.theme) === canonicalJson(destination.theme);
    const resourceSnapshots = contextsMatch ? snapshots : importedHeaderFooterSnapshots;
    const referencedStyleIds = new Set<string>();
    let materializeDefaultParagraphStyle = false;
    for (const snapshot of resourceSnapshots) {
      if (hasUnstyledParagraph(sourceDocumentOf(snapshot))) {
        materializeDefaultParagraphStyle = true;
      }
      for (const styleId of collect(sourceDocumentOf(snapshot))) {
        if (sourceStyleIds.has(styleId) && !existing.has(styleId)) referencedStyleIds.add(styleId);
      }
    }
    const reservedStyleIds = new Set<string>();
    if (referencedStyleIds.size > 0) {
      for (const handle of this.listStoryHandlesInternal()) {
        const state = this.getEditableStoryState(handle);
        if (!state) continue;
        for (const styleId of collect(state.doc)) {
          if (referencedStyleIds.has(styleId)) reservedStyleIds.add(styleId);
        }
      }
    }
    const result = importReferencedStyleDefinitions({
      sourceStyles: sourcePackage.styles,
      destinationStyles,
      sourceTheme: sourcePackage.theme,
      destinationTheme: destination.theme,
      referencedStyleIds: [...referencedStyleIds],
      reservedStyleIds: [...reservedStyleIds],
      materializeDefaultParagraphStyle,
    });
    if (result.status === "imported") {
      this.importedStyles = result.styles;
      this.state = withDocumentStyles(this.state, result.styles);
      for (const entry of this.secondaryStoryStates.values()) {
        const wasUntouched = entry.state === entry.initialState;
        const refreshed = withDocumentStyles(entry.state, result.styles);
        entry.state = refreshed;
        if (wasUntouched) entry.initialState = refreshed;
      }
    }
    return result;
  }

  /**
   * A raw final table has no paragraph mark after it for Word to merge into
   * when the target ends in a paragraph. Folio-exact comparison adds this
   * untracked receiver before planning; its private marker makes resolving in
   * Folio restore the raw source view.
   */
  private stageTerminalTableReviewCarrier(target: PMNode): boolean {
    const baseTerminalTable = this.state.doc.lastChild;
    if (baseTerminalTable?.type.name !== "table" || target.lastChild?.type.name !== "paragraph") {
      return false;
    }
    const paragraphType = this.state.schema.nodes["paragraph"];
    if (!paragraphType) {
      return panic("The schema has no paragraph node for a terminal-table review carrier");
    }
    const carrier = paragraphType.create({ reviewCarrier: REVIEW_CARRIERS.TERMINAL_TABLE });
    this.state = this.state.apply(this.state.tr.insert(this.state.doc.content.size, carrier));
    return true;
  }

  private stageTargetNumbering(
    target: NumberingDefinitions | null | undefined,
    references: readonly { numId: number; level: number }[],
    remappedNumIds: ReadonlyMap<number, number>,
  ): StageTargetNumberingResult {
    if (references.length === 0) return "unchanged";
    if (!target) return "conflict";
    const current = this.baseDocument.package.numbering ?? { abstractNums: [], nums: [] };
    const nums = new Map(current.nums.map((entry) => [entry.numId, entry]));
    const abstracts = new Map(current.abstractNums.map((entry) => [entry.abstractNumId, entry]));
    const targetNums = new Map(target.nums.map((entry) => [entry.numId, entry]));
    const targetAbstracts = new Map(
      target.abstractNums.map((entry) => [entry.abstractNumId, entry]),
    );
    const levelsByNumId = referencedNumberingLevelsByNumId({ references });
    const importedAbstractIds = new Map<number, number>();
    let nextAbstractId = 0;
    for (const id of abstracts.keys()) nextAbstractId = Math.max(nextAbstractId, id + 1);
    for (const numId of levelsByNumId.keys()) {
      const targetNum = targetNums.get(numId);
      if (!targetNum) return "conflict";
      const targetAbstract = targetAbstracts.get(targetNum.abstractNumId);
      if (!targetAbstract) return "conflict";
      const stagedNumId = remappedNumIds.get(numId) ?? numId;
      const existingNum = nums.get(stagedNumId);
      if (existingNum) {
        if (
          stagedNumId === numId &&
          sameReferencedNumberingLevels({
            current,
            target,
            numId,
            levelsByNumId,
          })
        ) {
          continue;
        }
        return "conflict";
      }
      let abstractNumId = importedAbstractIds.get(targetAbstract.abstractNumId);
      if (abstractNumId === undefined) {
        abstractNumId = targetAbstract.abstractNumId;
        const collision = abstracts.get(abstractNumId);
        if (collision && canonicalJson(collision) !== canonicalJson(targetAbstract)) {
          abstractNumId = nextAbstractId++;
        }
        nextAbstractId = Math.max(nextAbstractId, abstractNumId + 1);
        importedAbstractIds.set(targetAbstract.abstractNumId, abstractNumId);
        abstracts.set(abstractNumId, { ...targetAbstract, abstractNumId });
      }
      nums.set(stagedNumId, { ...targetNum, numId: stagedNumId, abstractNumId });
    }
    if (nums.size === current.nums.length) return "unchanged";
    const numbering = {
      ...current,
      abstractNums: [...abstracts.values()],
      nums: [...nums.values()],
    };
    this.baseDocument.package.numbering = numbering;
    this.state = withDocumentNumbering(this.state, numbering);
    for (const entry of this.secondaryStoryStates.values()) {
      entry.state = withDocumentNumbering(entry.state, numbering);
    }
    return "staged";
  }

  private planTargetNumberingReferences(
    target: NumberingDefinitions | null | undefined,
    references: readonly { numId: number; level: number }[],
  ): ReadonlyMap<number, number> | null {
    if (references.length === 0) return new Map();
    if (!target) return null;
    const current = this.baseDocument.package.numbering ?? { abstractNums: [], nums: [] };
    const nums = new Map(current.nums.map((entry) => [entry.numId, entry]));
    const targetNums = new Map(target.nums.map((entry) => [entry.numId, entry]));
    const targetAbstracts = new Map(
      target.abstractNums.map((entry) => [entry.abstractNumId, entry]),
    );
    const levelsByNumId = referencedNumberingLevelsByNumId({ references });
    let nextNumId = 0;
    for (const id of nums.keys()) nextNumId = Math.max(nextNumId, id + 1);
    for (const id of targetNums.keys()) nextNumId = Math.max(nextNumId, id + 1);
    const remapped = new Map<number, number>();
    for (const numId of levelsByNumId.keys()) {
      const targetNum = targetNums.get(numId);
      const targetAbstract = targetNum ? targetAbstracts.get(targetNum.abstractNumId) : undefined;
      if (!targetNum || !targetAbstract) return null;
      const existingNum = nums.get(numId);
      if (!existingNum) continue;
      if (!sameReferencedNumberingLevels({ current, target, numId, levelsByNumId })) {
        remapped.set(numId, nextNumId++);
      }
    }
    return remapped;
  }

  /** Parse a `.docx` buffer into a reviewer. */
  static async fromBuffer(
    buffer: ArrayBuffer,
    options: FolioDocxReviewerOptions = {},
  ): Promise<FolioDocxReviewer> {
    const baseDocument = await parseDocx(buffer, {
      detectVariables: false,
      preloadFonts: false,
      password: options.password,
    });
    // The change tracker feeds selective save, and the paraId allocator hands
    // inserted paragraphs stable ids. Editor history and collaboration are
    // intentionally absent: reviewer undo stores complete state snapshots,
    // and headless resolution consumes its replacement transaction locally.
    const plugins = createHeadlessPlugins(
      baseDocument.package.styles,
      baseDocument.package.numbering,
    );
    // Allocate paraIds up front (the editor does this on load) so every block
    // anchors on a stable id and the selective-save path can key changed
    // paragraphs by paraId. Deterministic (not random) allocation so a
    // paraId-less document yields the SAME block ids on every parse: ops built
    // from one parse's snapshot then resolve against another parse of the same
    // bytes instead of skipping as stale anchors.
    //
    // `ensureBaseDirectionInState` seeds `dir` on RTL paragraphs the same way
    // the editor load path does. It used to run only as a side effect of the
    // paraId transaction, so a document that already carried a complete set of
    // Word-authored paraIds silently skipped bidi seeding.
    const state = ensureBaseDirectionInState(
      EditorState.create({
        schema,
        doc: ensureDeterministicParaIdsInDoc(toProseDoc(baseDocument)),
        plugins,
      }),
    );
    return new FolioDocxReviewer({
      baseDocument,
      originalBuffer: baseDocument.originalBuffer ?? buffer,
      state,
      author: options.author ?? "AI",
    });
  }

  /**
   * Snapshot the current document into AI-facing blocks (id, kind, text) plus
   * the anchor map the apply layer resolves operations against.
   */
  snapshot(): FolioAIEditSnapshot {
    return createStateSnapshot(this.state);
  }

  /**
   * The package's numbering, flattened to what a reader sees: one entry per
   * numbering instance and level, with the format and level text that produce
   * its labels.
   *
   * Labels are rendered from these definitions rather than stored on the
   * paragraphs, so nothing in a block projection changes when a list is
   * renumbered — which is right for an insertion that renumbers the items
   * below it, and wrong for a list whose FORMAT changed. Sorted so two
   * packages can be compared entry by entry.
   */
  readNumberingDefinitions(): FolioNumberingLevel[] {
    return numberingLevelsOf(this.baseDocument.package.numbering).toSorted(
      (left, right) => left.numId - right.numId || left.level - right.level,
    );
  }

  /** Return parsed package metadata without exposing the mutable document model. */
  getDocumentProperties(): Readonly<NonNullable<Document["package"]["properties"]>> | null {
    const properties = this.baseDocument.package.properties;
    if (!properties) {
      return null;
    }
    return Object.freeze({
      ...properties,
      ...(properties.created !== undefined
        ? { created: new Date(properties.created.getTime()) }
        : {}),
      ...(properties.modified !== undefined
        ? { modified: new Date(properties.modified.getTime()) }
        : {}),
    });
  }

  /** Snapshot one editable story into stable, operation-ready blocks. */
  snapshotStory(story: FolioEditableDocumentStoryHandle): FolioAIEditSnapshot | null {
    const state = this.getEditableStoryState(story);
    return state ? createStateSnapshot(state) : null;
  }

  private snapshotReviewedStoryInternal(
    options: FolioReadReviewedStoryOptions = {},
  ): FolioAIEditSnapshot | null {
    const story = options.story ?? MAIN_STORY;
    const view = options.view ?? "final";
    if (!isFolioReviewedView(view)) {
      throw new UnsupportedFolioReviewedViewError({
        message: "Unsupported reviewed document view.",
        receivedView: view,
      });
    }
    const sourceState = this.getEditableStoryState(story);
    return sourceState ? createStateSnapshot(resolveReviewedState(sourceState, view)) : null;
  }

  private projectComparisonStoriesInternal(
    mode: FolioDocxComparisonProjectionMode,
  ): FolioDocxComparisonProjection {
    const handles = this.listStoryHandlesInternal();
    let highestId = 0;
    let present = false;
    if (mode === "with-revision-census") {
      for (const change of this.currentFinalSectionProperties()?.propertyChanges ?? []) {
        highestId = Math.max(highestId, change.info.id);
        present = true;
      }
      // Census every arriving story before resolution mutates any reviewer
      // state. The shared interpreter omits block ids, so this costs one
      // carrier walk per story without constructing a throwaway snapshot.
      for (const handle of handles) {
        const state = this.getEditableStoryState(handle);
        if (!state) {
          continue;
        }
        const stats = getTrackedChangeStatsFromDoc(state.doc);
        highestId = Math.max(highestId, stats.highestId);
        present ||= stats.present;
      }
    }

    const stories: FolioDocxComparisonStoryProjection[] = [];
    for (const handle of handles) {
      stories.push({
        handle,
        snapshot: this.resolveReviewedStorySnapshotInternal({ story: handle, view: "final" }),
      });
    }
    return { stories, revisions: { highestId, present } };
  }

  /**
   * One story's tables, in the document order {@link snapshotStory} numbers
   * them with, read through a reviewed view.
   *
   * A block snapshot says which table a paragraph sits in and nothing about
   * the table itself. A caller that has to reproduce one — a comparison
   * copying the target's table into the base, or checking that accepting its
   * own redline reproduced it — needs the node: its `w:tblPr`, its `w:tblGrid`
   * widths, and every row's and cell's properties.
   */
  storyTables({
    story = MAIN_STORY,
    view = "final",
  }: FolioReadReviewedStoryOptions = {}): readonly FolioStoryTable[] {
    if (!isFolioReviewedView(view)) {
      throw new UnsupportedFolioReviewedViewError({
        message: "Unsupported reviewed document view.",
        receivedView: view,
      });
    }
    const sourceState = this.getEditableStoryState(story);
    return sourceState ? folioStoryTables(resolveReviewedState(sourceState, view).doc) : [];
  }

  /**
   * Move the paired tables' own properties onto this story's tables, as
   * tracked property changes.
   *
   * `w:tblPr`, `w:trPr` and `w:tcPr` belong to no block, so no block operation
   * can carry them: a table that stayed in place while its widths, shading,
   * borders or header row changed reads as unedited. Each difference is
   * written as the target's property set plus a `w:tblPrChange` /
   * `w:trPrChange` / `w:tcPrChange` holding the previous one, which is what a
   * reject restores. A set that already agrees produces no revision.
   */
  matchStoryTableGeometry({
    story = MAIN_STORY,
    targetTables,
    pairings,
    revisionStamp,
  }: FolioMatchStoryTableGeometryOptions): number {
    const state = this.getEditableStoryState(story);
    if (!state || pairings.length === 0) {
      return revisionStamp.idSeed;
    }
    const transaction = state.tr;
    const { nextRevisionId } = matchTableGeometry({
      tr: transaction,
      baseTables: folioStoryTables(state.doc),
      targetTables,
      pairings,
      revision: {
        author: this.author,
        date: revisionStamp.date,
        idSeed: revisionStamp.idSeed,
      },
    });
    if (transaction.docChanged) {
      this.setEditableStoryState(story, state.apply(transaction));
    }
    return nextRevisionId;
  }

  /** Read one story through an immutable reviewed-view projection. */
  readReviewedStory(options: FolioReadReviewedStoryOptions = {}): FolioReviewedStory | null {
    const story = options.story ?? MAIN_STORY;
    const view = options.view ?? "final";
    const snapshot = this.snapshotReviewedStoryInternal({ story, view });
    if (!snapshot) {
      return null;
    }
    return {
      story,
      view,
      snapshot,
      text: formatStorySnapshotForLLM(snapshot, view === "current-markup"),
      changes: getTrackedChangesFromSnapshot(snapshot),
    };
  }

  /** Resolve one editable story to its original or final state. */
  resolveReviewedStory({ story = MAIN_STORY, view }: FolioResolveReviewedStoryOptions): boolean {
    return this.resolveReviewedStorySnapshotInternal({ story, view }) !== null;
  }

  private resolveReviewedStorySnapshotInternal({
    story = MAIN_STORY,
    view,
  }: FolioResolveReviewedStoryOptions): FolioAIEditSnapshot | null {
    if (!isFolioResolvedReviewedView(view)) {
      throw new UnsupportedFolioReviewedViewError({
        message: "Only original and final views can replace editable story state.",
        receivedView: view,
      });
    }
    const sourceState = this.getEditableStoryState(story);
    if (!sourceState) {
      return null;
    }
    const resolvedState = resolveReviewedState(sourceState, view);
    this.setEditableStoryState(story, resolvedState);
    const snapshot = createStateSnapshot(resolvedState);
    this.resolvedStoryExpectations.set(editableStoryKey(story), {
      story,
      text: formatStorySnapshotForLLM(snapshot, false),
      blocks: snapshot.blocks.map(resolvedStoryBlockProjection),
    });
    return snapshot;
  }

  /**
   * Apply operations against the current state. Reuses the live-editor applier
   * verbatim via a headless `{ state, dispatch }` seam; the resulting state
   * (including any comments) is retained for {@link toBuffer}.
   */
  applyOperations(
    operations: FolioAIEditOperation[],
    options: FolioApplyOperationsOptions = {},
  ): FolioAIEditApplyResult {
    const { applied, skipped } = this.applyDocumentOperationsInternal({
      story: MAIN_STORY,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations,
        mode: options.mode ?? "tracked-changes",
      },
      ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
      ...(options.revisionStamp !== undefined && { revisionStamp: options.revisionStamp }),
      ...(options.wordDiff !== undefined && { wordDiff: options.wordDiff }),
      createUndoEntry: false,
    });
    return { applied, skipped };
  }

  /**
   * Apply a versioned document-operation batch against the current state.
   * This is the contract entry point for serialized callers; the legacy
   * {@link applyOperations} method delegates here so both APIs keep identical
   * edit semantics.
   */
  applyDocumentOperations(
    batch: FolioDocumentOperationBatch,
    options: FolioApplyDocumentOperationsOptions = {},
  ): FolioDocumentOperationResult {
    return this.applyDocumentOperationsInternal({
      story: MAIN_STORY,
      batch,
      ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
      ...(options.revisionStamp !== undefined && { revisionStamp: options.revisionStamp }),
      ...(options.wordDiff !== undefined && { wordDiff: options.wordDiff }),
      createUndoEntry: true,
    });
  }

  /** Apply a versioned operation batch to one editable document story. */
  applyDocumentOperationsToStory({
    story,
    batch,
    snapshot,
    revisionStamp,
    wordDiff,
    tableTemplates,
  }: FolioApplyDocumentOperationsToStoryOptions): FolioDocumentOperationResult {
    return this.applyDocumentOperationsInternal({
      story,
      batch,
      ...(snapshot !== undefined && { snapshot }),
      ...(revisionStamp !== undefined && { revisionStamp }),
      ...(wordDiff !== undefined && { wordDiff }),
      ...(tableTemplates !== undefined && { tableTemplates }),
      createUndoEntry: true,
    });
  }

  private applyDocumentOperationsInternal({
    story,
    batch,
    snapshot,
    revisionStamp,
    wordDiff,
    tableTemplates,
    createUndoEntry,
  }: ApplyDocumentOperationsInternalOptions): FolioDocumentOperationResult {
    const beforeState = this.requireEditableStoryState(story);
    const createdCommentsLengthBefore = this.createdComments.length;
    const view = {
      state: beforeState,
      dispatch: (transaction: Transaction) => {
        view.state = view.state.apply(transaction);
      },
    };

    const result = applyFolioDocumentOperations({
      view,
      snapshot: snapshot ?? createStateSnapshot(beforeState),
      batch,
      story: story.type === "main" ? "main" : story,
      author: this.author,
      ...(revisionStamp !== undefined && { revisionStamp }),
      ...(wordDiff !== undefined && { wordDiff }),
      ...(tableTemplates !== undefined && { tableTemplates }),
      createCommentId: (text) => {
        const comment = createReviewerComment(this.nextCommentId(), text, this.author);
        this.createdComments.push(comment);
        return comment.id;
      },
      ...(createUndoEntry && {
        createUndoHandle: () => ({
          type: "documentOperationUndo",
          id: `headless-${String(undoHandleCursor++)}`,
        }),
      }),
    });

    this.setEditableStoryState(story, view.state);
    if (result.undoHandle !== null) {
      this.documentOperationUndoEntries.push({
        undoHandle: result.undoHandle,
        story,
        beforeState,
        afterState: view.state,
        createdCommentsLengthBefore,
        createdCommentsLengthAfter: this.createdComments.length,
      });
    }
    return result;
  }

  /** Undo the latest unchanged document-operation batch and its created comments. */
  undoDocumentOperations(
    undoHandle: FolioDocumentOperationUndoHandle,
  ): FolioDocumentOperationUndoResult {
    const entryIndex = this.documentOperationUndoEntries.findIndex(
      (entry) => entry.undoHandle.type === undoHandle.type && entry.undoHandle.id === undoHandle.id,
    );
    if (entryIndex === -1) {
      return { status: "rejected", undoHandle, reason: "unknownHandle" };
    }
    if (entryIndex !== this.documentOperationUndoEntries.length - 1) {
      return { status: "rejected", undoHandle, reason: "notLatest" };
    }

    const entry = this.documentOperationUndoEntries.at(-1);
    if (!entry) {
      return { status: "rejected", undoHandle, reason: "unknownHandle" };
    }
    if (
      this.requireEditableStoryState(entry.story) !== entry.afterState ||
      this.createdComments.length !== entry.createdCommentsLengthAfter
    ) {
      return { status: "rejected", undoHandle, reason: "documentChanged" };
    }

    this.setEditableStoryState(entry.story, entry.beforeState);
    this.createdComments.length = entry.createdCommentsLengthBefore;
    this.documentOperationUndoEntries.pop();
    return { status: "undone", undoHandle };
  }

  /**
   * The body blocks with their stable ids, in document order — the same
   * `FolioAIBlock` shape {@link snapshot} builds, reused verbatim so the ids
   * feed straight back into {@link applyOperations} and accept / reject.
   */
  getContent(): FolioAIBlock[] {
    return [...this.snapshot().blocks];
  }

  /**
   * The body as LLM-ready plain text: one line per block, each prefixed with
   * its stable block id (`[<blockId>] text`). Copyable verbatim into a prompt
   * without JSON quote-escaping.
   *
   * With `{ annotated: true }` each block's text is rendered redline-aware:
   * tracked insertions/deletions and comment anchors appear inline as
   * `<ins>` / `<del>` / `<comment>` tags (see {@link buildAnnotatedBlockText})
   * for prompt-embedding parity with a live editor's redline view. The default
   * (clean) output flattens tracked changes and is unchanged.
   */
  getContentAsText(options: FolioGetContentAsTextOptions = {}): string {
    return formatStorySnapshotForLLM(this.snapshot(), options.annotated === true);
  }

  /**
   * Header / footer and footnote / endnote text as labeled, LLM-ready lines,
   * one per non-empty part: `[header default] …`, `[footer default] …`,
   * `[footnote #N] …`, `[endnote #N] …`. Lines reflect in-memory edits. Empty
   * parts and separator notes are omitted.
   */
  getNotesAsText(): string {
    const pkg = this.baseDocument.package;
    const lines: string[] = [];

    const pushHeaderFooter = (
      map: Map<string, HeaderFooter> | undefined,
      label: "header" | "footer",
    ) => {
      if (!map) {
        return;
      }
      for (const [relationshipId, hf] of map) {
        const handle = { type: label, relationshipId } as const;
        const text = this.getHeaderFooterStoryText(handle, hf);
        if (text.length > 0) {
          lines.push(`[${label} ${hf.hdrFtrType}] ${text}`);
        }
      }
    };

    pushHeaderFooter(pkg.headers, "header");
    pushHeaderFooter(pkg.footers, "footer");

    for (const footnote of pkg.footnotes ?? []) {
      if (isSeparatorFootnote(footnote)) {
        continue;
      }
      const handle = { type: "footnote", noteId: footnote.id } as const;
      const text = this.getNoteStoryText(handle, footnote);
      if (text.length > 0) {
        lines.push(`[footnote #${footnote.id}] ${text}`);
      }
    }
    for (const endnote of pkg.endnotes ?? []) {
      if (isSeparatorEndnote(endnote)) {
        continue;
      }
      const handle = { type: "endnote", noteId: endnote.id } as const;
      const text = this.getNoteStoryText(handle, endnote);
      if (text.length > 0) {
        lines.push(`[endnote #${endnote.id}] ${text}`);
      }
    }

    return lines.join("\n");
  }

  private listStoryHandlesInternal(): FolioDocumentStoryHandle[] {
    const pkg = this.baseDocument.package;
    const handles: FolioDocumentStoryHandle[] = [{ type: "main" }];
    for (const relationshipId of pkg.headers?.keys() ?? []) {
      handles.push({ type: "header", relationshipId });
    }
    for (const relationshipId of this.importedHeaders.keys()) {
      handles.push({ type: "header", relationshipId });
    }
    for (const relationshipId of pkg.footers?.keys() ?? []) {
      handles.push({ type: "footer", relationshipId });
    }
    for (const relationshipId of this.importedFooters.keys()) {
      handles.push({ type: "footer", relationshipId });
    }
    for (const footnote of pkg.footnotes ?? []) {
      if (!isSeparatorFootnote(footnote)) {
        handles.push({ type: "footnote", noteId: footnote.id });
      }
    }
    for (const endnote of pkg.endnotes ?? []) {
      if (!isSeparatorEndnote(endnote)) {
        handles.push({ type: "endnote", noteId: endnote.id });
      }
    }
    return handles.filter(
      (handle) =>
        (handle.type !== "header" && handle.type !== "footer") ||
        !this.removedHeaderFooterStories.has(headerFooterStoryKey(handle)),
    );
  }

  /** Discover every readable document story through a typed, serializable handle. */
  listStories(): FolioDocumentStory[] {
    return this.listStoryHandlesInternal().map((handle) => ({
      handle,
      text: this.getStoryText(handle),
    }));
  }

  /** Read one discovered story; returns null when its handle is no longer present. */
  readStory(handle: FolioDocumentStoryHandle): FolioDocumentStory | null {
    const stories = this.listStories();
    if (handle.type === "main") {
      return stories.at(0) ?? null;
    }
    if (handle.type === "header" || handle.type === "footer") {
      return (
        stories.find(
          ({ handle: candidate }) =>
            candidate.type === handle.type && candidate.relationshipId === handle.relationshipId,
        ) ?? null
      );
    }
    return (
      stories.find(
        ({ handle: candidate }) =>
          (candidate.type === "footnote" || candidate.type === "endnote") &&
          candidate.type === handle.type &&
          candidate.noteId === handle.noteId,
      ) ?? null
    );
  }

  /**
   * The tracked changes present in the body, including inline and structural
   * revisions. Each carries the revision id {@link acceptChange} /
   * {@link rejectChange} resolve against. Related revision sites fold into one
   * entry where the document model identifies them as one authored change.
   */
  getChanges(filter?: FolioReviewChangeFilter): FolioReviewChange[] {
    const changes = [
      ...getTrackedChangesFromDoc(this.state.doc),
      ...(this.currentFinalSectionProperties()?.propertyChanges ?? []).map(({ info }) => ({
        id: info.id,
        type: "sectionPropertiesChanged" as const,
        author: info.author,
        date: info.date ?? null,
        text: "",
        blockId: null,
      })),
    ];
    if (!filter) return changes;
    return changes.filter(
      (change) =>
        (filter.author === undefined || change.author === filter.author) &&
        (filter.type === undefined || change.type === filter.type),
    );
  }

  /**
   * The comment threads in the document — comments parsed from the package plus
   * any the reviewer authored — with their anchored text and containing block
   * id. Only run text is read from a comment body; richer comment content
   * (nested tracked changes, hyperlinks) is out of scope.
   */
  getComments(filter?: FolioReviewCommentFilter): FolioReviewComment[] {
    const definitions = this.withResolvedOverrides([
      ...(this.baseDocument.package.document.comments ?? []),
      ...this.createdComments,
    ]);
    if (definitions.length === 0) {
      return [];
    }

    const anchors = this.commentAnchors();
    const repliesByParent = new Map<number, Comment[]>();
    const topLevel: Comment[] = [];
    for (const comment of definitions) {
      if (comment.parentId === undefined) {
        topLevel.push(comment);
        continue;
      }
      const siblings = repliesByParent.get(comment.parentId) ?? [];
      siblings.push(comment);
      repliesByParent.set(comment.parentId, siblings);
    }

    const threads = topLevel.map<FolioReviewComment>((comment) => {
      const anchor = anchors.get(comment.id);
      return {
        id: comment.id,
        author: comment.author,
        date: comment.date ?? null,
        text: commentPlainText(comment),
        anchoredText: anchor?.text ?? "",
        blockId: anchor?.blockId ?? null,
        replies: (repliesByParent.get(comment.id) ?? []).map((reply) => ({
          id: reply.id,
          author: reply.author,
          date: reply.date ?? null,
          text: commentPlainText(reply),
        })),
        done: comment.done ?? false,
      };
    });

    if (!filter) {
      return threads;
    }
    return threads.filter(
      (thread) =>
        (filter.author === undefined || thread.author === filter.author) &&
        (filter.done === undefined || thread.done === filter.done),
    );
  }

  /**
   * Add a reply to a comment thread. Pass a {@link FolioReviewComment} from
   * {@link getComments} or its id; the reply threads under that comment's root
   * (Word threads are flat). On {@link toBuffer} the reply is written as a real
   * Word reply — linked via `commentsExtended.xml` and given its own
   * `commentRange` markers + reference anchored on the parent's range. Returns
   * the created reply, or `null` when the target comment is absent.
   */
  replyTo(
    target: FolioReviewComment | number,
    input: FolioReviewReplyInput,
  ): FolioReviewCommentReply | null {
    const parentId = typeof target === "number" ? target : target.id;
    const existing = [
      ...(this.baseDocument.package.document.comments ?? []),
      ...this.createdComments,
    ];
    const reply = createReply(existing, parentId, {
      author: input.author ?? this.author,
      text: input.text,
      ...(input.initials !== undefined ? { initials: input.initials } : {}),
    });
    if (!reply) {
      return null;
    }
    this.createdComments.push(reply);
    this.usedCommentIds.add(reply.id);
    return { id: reply.id, author: reply.author, date: reply.date ?? null, text: input.text };
  }

  /**
   * Mark a comment thread resolved, or reopen a previously resolved one. Pass
   * the id from {@link getComments} / {@link FolioReviewComment.id}. Applies
   * only to the target comment id, not cascaded to its replies: Word keys the
   * resolved marker off the comment's own `w15:commentEx` entry
   * (`commentSerializer.ts`'s `buildCommentExtendedEntries` reads only
   * `comment.done` for the id it is building an entry for), so resolving the
   * thread root is sufficient and reply entries need no `done` flag of their
   * own. On {@link toBuffer} the state is written to `commentsExtended.xml`
   * (`w15:done`) through the same channel `replyTo`-created comments use.
   * Returns `false` when no comment with that id exists.
   */
  resolveComment(commentId: string, options: { resolved?: boolean } = {}): boolean {
    const resolved = options.resolved ?? true;
    const existing = [
      ...(this.baseDocument.package.document.comments ?? []),
      ...this.createdComments,
    ];
    const target = existing.find((comment) => String(comment.id) === commentId);
    if (!target) {
      return false;
    }
    this.resolvedOverrides.set(target.id, resolved);
    return true;
  }

  private currentFinalSectionProperties(): SectionProperties | undefined {
    return (
      this.finalSectionPropertiesOverride ??
      this.baseDocument.package.document.finalSectionProperties
    );
  }

  private stageFinalSectionProperties({
    target,
    previous,
    revision,
  }: {
    target: SectionProperties;
    previous: SectionProperties;
    revision: FolioRevisionStamp;
  }): boolean {
    if (canonicalJson(target) === canonicalJson(previous)) return false;
    const previousReferences = sectionReferenceHistory({ previous, target });
    this.finalSectionPropertiesOverride = {
      ...target,
      propertyChanges: [
        {
          type: "sectionPropertyChange",
          info: { id: revision.idSeed, author: this.author, date: revision.date },
          previousProperties: previous,
          ...(previousReferences !== undefined && { previousReferences }),
          currentProperties: target,
        },
      ],
    };
    return true;
  }

  private stageMappedSectionBoundaries({
    target,
    originalRevisionIdSeed,
    maxRanges,
    revision,
    mapTargetProperties,
  }: {
    target: PMNode;
    originalRevisionIdSeed: number;
    maxRanges: number;
    revision: FolioRevisionStamp;
    mapTargetProperties: (args: {
      kind: "inserted" | "retained";
      current: SectionProperties | undefined;
      target: SectionProperties;
    }) =>
      | { kind: "inserted"; target: SectionProperties }
      | { kind: "retained"; previous: SectionProperties; target: SectionProperties }
      | null;
  }):
    | { status: "matched"; rangeCount: number; nextRevisionId: number; documentChanged: boolean }
    | { status: "unalignable"; detail: string }
    | { status: "budget-exceeded" } {
    const result = stageMappedSectionBoundaryProperties({
      state: this.state,
      target,
      originalRevisionIdSeed,
      revisionStamp: revision,
      author: this.author,
      maxRanges,
      mapTargetProperties,
    });
    if (result.status !== "matched") return result;
    const documentChanged = result.transaction.docChanged;
    if (documentChanged) this.state = this.state.apply(result.transaction);
    return {
      status: "matched",
      rangeCount: result.rangeCount,
      nextRevisionId: result.nextRevisionId,
      documentChanged,
    };
  }

  private resolveFinalSectionProperties(mode: "accept" | "reject", id: number): number {
    const current = this.currentFinalSectionProperties();
    const changes = current?.propertyChanges;
    const targetIndex = changes?.findIndex(({ info }) => info.id === id);
    if (
      !current ||
      !changes ||
      changes.length === 0 ||
      targetIndex === undefined ||
      targetIndex < 0
    ) {
      return 0;
    }
    const { propertyChanges: _propertyChanges, ...liveProperties } = current;
    const previousStates = changes.map((change) =>
      sectionRejectProperties({
        live: current,
        previousProperties: change.previousProperties,
        previousReferences: change.previousReferences,
      }),
    );
    let resolvedProperties: SectionProperties = previousStates[0] ?? liveProperties;
    const remaining = [];
    for (const [index, change] of changes.entries()) {
      const nextProperties = previousStates[index + 1] ?? liveProperties;
      if (index === targetIndex) {
        if (mode === "accept") resolvedProperties = nextProperties;
        continue;
      }
      remaining.push({
        ...change,
        previousProperties: resolvedProperties,
        ...(change.previousReferences !== undefined && {
          previousReferences: {
            ...(resolvedProperties.headerReferences !== undefined && {
              headerReferences: resolvedProperties.headerReferences,
            }),
            ...(resolvedProperties.footerReferences !== undefined && {
              footerReferences: resolvedProperties.footerReferences,
            }),
          },
        }),
      });
      resolvedProperties = nextProperties;
    }
    this.finalSectionPropertiesOverride = {
      ...resolvedProperties,
      ...(remaining.length > 0 && { propertyChanges: remaining }),
    };
    return 1;
  }

  /**
   * Accept an existing tracked change, keeping its text and dropping the
   * redline. Pass a {@link FolioReviewChange} from {@link getChanges} or its
   * revision id. Reuses the editor's own accept command headlessly, so the
   * resolved state persists on {@link toBuffer}. Returns `false` when the
   * revision is no longer present (already resolved, or never existed).
   */
  acceptChange(target: FolioReviewChange | number): boolean {
    return this.resolveWithSectionReferenceHistory(() => this.acceptChangeInternal(target));
  }

  private acceptChangeInternal(target: FolioReviewChange | number): boolean {
    const id = revisionIdOf(target);
    const bodyChanged = this.runCommand(acceptAIEditRevision(id));
    const sectionChanged = this.resolveFinalSectionProperties("accept", id) > 0;
    return bodyChanged || sectionChanged;
  }

  /**
   * Reject an existing tracked change: an insertion's text is removed, a
   * deletion's text is restored. See {@link acceptChange} for targeting.
   */
  rejectChange(target: FolioReviewChange | number): boolean {
    return this.resolveWithSectionReferenceHistory(() => this.rejectChangeInternal(target));
  }

  private rejectChangeInternal(target: FolioReviewChange | number): boolean {
    const id = revisionIdOf(target);
    const bodyChanged = this.runCommand(rejectAIEditRevision(id));
    const sectionChanged = this.resolveFinalSectionProperties("reject", id) > 0;
    return bodyChanged || sectionChanged;
  }

  /**
   * Accept every tracked change in the package. Returns the number of changes
   * present before the sweep.
   *
   * Every story, not just the body: a revision in a header or a footnote is
   * one a reviewer meant to resolve, and leaving it behind means an accepted
   * document still carries a redline nobody can see from the body.
   */
  acceptAll(): number {
    return this.resolveWithSectionReferenceHistory(
      () => this.resolveEveryStory("accept") + this.resolveEveryFinalSectionChange("accept"),
    );
  }

  /** Reject every tracked change in the package. See {@link acceptAll}. */
  rejectAll(): number {
    return this.resolveWithSectionReferenceHistory(
      () => this.resolveEveryStory("reject") + this.resolveEveryFinalSectionChange("reject"),
    );
  }

  private resolveWithSectionReferenceHistory<T>(resolve: () => T): T {
    const before = captureSectionReferenceInventory(
      this.state.doc,
      this.currentFinalSectionProperties(),
    );
    const result = resolve();
    const removedEndpointReferences =
      getTrackedSectionEndpointRemoval(this.state)?.removedReferences ?? [];
    if (before.revisionRelationships.size > 0 || removedEndpointReferences.length > 0) {
      const after = captureSectionReferenceInventory(
        this.state.doc,
        this.currentFinalSectionProperties(),
      );
      const removed = resolvedSectionReferenceLosses({ before, after });
      this.sectionReferenceRemovals.push(...removed);
      for (const reference of [...removed, ...removedEndpointReferences]) {
        if (
          after.references.some(
            ({ part, relationshipId }) =>
              part === reference.part && relationshipId === reference.relationshipId,
          )
        )
          continue;
        const handle = { type: reference.part, relationshipId: reference.relationshipId };
        this.removedHeaderFooterStories.set(headerFooterStoryKey(handle), handle);
      }
    }
    return result;
  }

  private resolveEveryFinalSectionChange(mode: "accept" | "reject"): number {
    const current = this.currentFinalSectionProperties();
    const firstChange = current?.propertyChanges?.at(0);
    if (!current || !firstChange) return 0;
    const { propertyChanges, ...liveProperties } = current;
    this.finalSectionPropertiesOverride =
      mode === "accept"
        ? liveProperties
        : sectionRejectProperties({
            live: current,
            previousProperties: firstChange.previousProperties,
            previousReferences: firstChange.previousReferences,
          });
    return propertyChanges?.length ?? 0;
  }

  private resolveEveryStory(mode: "accept" | "reject"): number {
    let count = 0;
    for (const { handle } of this.listStories()) {
      const state = this.getEditableStoryState(handle);
      if (!state) {
        continue;
      }
      count += getTrackedChangesFromDoc(state.doc).length;
      this.setEditableStoryState(handle, resolveAllChangesInHeadlessState(state, mode));
    }
    return count;
  }

  /** The current document model with edits merged back in. */
  toDocument(): Document {
    return this.documentFromStateSnapshot(this.captureReviewerState());
  }

  private captureReviewerState(): FolioReviewerStateSnapshot {
    const secondaryStoryStates: FolioSecondaryStoryState[] = [];
    for (const { handle, initialState, state } of this.secondaryStoryStates.values()) {
      secondaryStoryStates.push({ handle, initialState, state });
    }
    return {
      mainState: this.state,
      finalSectionPropertiesOverride: this.finalSectionPropertiesOverride,
      secondaryStoryStates,
      sectionReferenceRemovals: [...this.sectionReferenceRemovals],
      removedHeaderFooterStories: [...this.removedHeaderFooterStories.values()],
      importedStyles: this.importedStyles,
      importedMedia: new Map(this.importedMedia),
      importedHeaders: new Map(this.importedHeaders),
      importedFooters: new Map(this.importedFooters),
      createdComments: [...this.createdComments],
      resolvedOverrides: new Map(this.resolvedOverrides),
      resolvedStoryExpectations: [...this.resolvedStoryExpectations.values()],
    };
  }

  private documentFromStateSnapshot(snapshot: FolioReviewerStateSnapshot): Document {
    const sourceDocument =
      snapshot.importedStyles === undefined
        ? this.baseDocument
        : {
            ...this.baseDocument,
            package: { ...this.baseDocument.package, styles: snapshot.importedStyles },
          };
    const document = updateDocumentContent(sourceDocument, snapshot.mainState.doc);
    if (snapshot.finalSectionPropertiesOverride !== undefined) {
      document.package.document.finalSectionProperties = snapshot.finalSectionPropertiesOverride;
    }
    if (snapshot.importedStyles !== undefined) document.package.styles = snapshot.importedStyles;
    if (snapshot.importedMedia.size > 0)
      document.package.media = new Map([
        ...(document.package.media ?? []),
        ...snapshot.importedMedia,
      ]);
    if (snapshot.importedHeaders.size > 0) {
      document.package.headers = new Map([
        ...(document.package.headers ?? []),
        ...snapshot.importedHeaders,
      ]);
    }
    if (snapshot.importedFooters.size > 0) {
      document.package.footers = new Map([
        ...(document.package.footers ?? []),
        ...snapshot.importedFooters,
      ]);
    }
    for (const handle of snapshot.removedHeaderFooterStories) {
      if (handle.type === "header") {
        document.package.headers = new Map(document.package.headers);
        document.package.headers.delete(handle.relationshipId);
      } else {
        document.package.footers = new Map(document.package.footers);
        document.package.footers.delete(handle.relationshipId);
      }
    }
    this.mergeEditedSecondaryStories(document, snapshot.secondaryStoryStates);
    if (snapshot.createdComments.length > 0 || snapshot.resolvedOverrides.size > 0) {
      document.package.document.comments = this.withResolvedOverrides(
        [...(document.package.document.comments ?? []), ...snapshot.createdComments],
        snapshot.resolvedOverrides,
      );
    }
    return document;
  }

  private captureSaveSnapshot(): FolioSaveSnapshot {
    const snapshot = this.captureReviewerState();
    const changedParaIds = new Set(getChangedParagraphIds(snapshot.mainState));
    let structuralChange =
      hasStructuralChanges(snapshot.mainState) ||
      snapshot.finalSectionPropertiesOverride !== undefined ||
      snapshot.importedStyles !== undefined ||
      snapshot.importedHeaders.size > 0 ||
      snapshot.importedFooters.size > 0;
    let untrackedChanges = hasUntrackedChanges(snapshot.mainState);
    const changedNoteParaIds = new Set<string>();
    for (const entry of snapshot.secondaryStoryStates) {
      if (entry.handle.type !== "footnote" && entry.handle.type !== "endnote") {
        continue;
      }
      for (const paraId of getChangedParagraphIds(entry.state)) {
        changedParaIds.add(paraId);
        changedNoteParaIds.add(paraId);
      }
      structuralChange ||= hasStructuralChanges(entry.state);
      untrackedChanges ||= hasUntrackedChanges(entry.state);
    }
    return {
      document: this.documentFromStateSnapshot(snapshot),
      path:
        structuralChange || untrackedChanges
          ? { type: "full-repack" }
          : { type: "selective-first", changedParaIds },
      changedNoteParaIds,
      sectionReferenceRemovals: snapshot.sectionReferenceRemovals,
      sectionEndpointRemoval: getTrackedSectionEndpointRemoval(snapshot.mainState),
      resolvedStoryExpectations: snapshot.resolvedStoryExpectations,
    };
  }

  /**
   * Serialise the reviewed document to a new `.docx` buffer. Tries a selective
   * patch first (only changed paragraphs are rewritten; every other byte of the
   * original package is preserved), falling back to a full repack for
   * structural edits — the same two-tier path the editor's save uses.
   */
  async toBuffer(): Promise<ArrayBuffer> {
    const save = this.captureSaveSnapshot();
    const selective =
      save.path.type === "selective-first"
        ? await this.trySelectiveSave(save.document, save.path.changedParaIds)
        : null;
    if (selective) {
      await this.assertResolvedStoriesSerialized(selective, save.resolvedStoryExpectations);
      return selective;
    }
    const repackDocument = { ...save.document, originalBuffer: this.originalBuffer };
    const repack = () =>
      repackDocx(repackDocument, { changedNoteParaIds: save.changedNoteParaIds });
    const repackReferences = () =>
      save.sectionReferenceRemovals.length > 0
        ? withSectionReferenceResolution({
            document: repackDocument,
            removedReferences: save.sectionReferenceRemovals,
            repack,
          })
        : repack();
    const buffer = save.sectionEndpointRemoval
      ? await withTrackedSectionEndpointRemoval({
          document: repackDocument,
          resolution: save.sectionEndpointRemoval,
          repack: repackReferences,
        })
      : await repackReferences();
    await this.assertResolvedStoriesSerialized(buffer, save.resolvedStoryExpectations);
    return buffer;
  }

  private async assertResolvedStoriesSerialized(
    buffer: ArrayBuffer,
    expectations: readonly FolioResolvedStoryExpectation[],
  ): Promise<void> {
    if (expectations.length === 0) {
      return;
    }
    const reopened = await FolioDocxReviewer.fromBuffer(buffer);
    for (const { story, text, blocks } of expectations) {
      const serialized = reopened.readReviewedStory({ story, view: "current-markup" });
      const serializedText = serialized
        ? formatStorySnapshotForLLM(serialized.snapshot, false)
        : null;
      const serializedBlocks = serialized
        ? serialized.snapshot.blocks.map(resolvedStoryBlockProjection)
        : null;
      const mismatches: FolioResolvedStorySerializationMismatch[] = [];
      if (!serialized) {
        mismatches.push(FOLIO_RESOLVED_STORY_SERIALIZATION_MISMATCHES.storyMissing);
      } else {
        if (serialized.changes.length > 0) {
          mismatches.push(FOLIO_RESOLVED_STORY_SERIALIZATION_MISMATCHES.revisionMarkupRemains);
        }
        if (serializedText !== text) {
          mismatches.push(FOLIO_RESOLVED_STORY_SERIALIZATION_MISMATCHES.textProjection);
        }
        if (JSON.stringify(serializedBlocks) !== JSON.stringify(blocks)) {
          mismatches.push(FOLIO_RESOLVED_STORY_SERIALIZATION_MISMATCHES.blockProjection);
        }
      }
      if (mismatches.length === 0) {
        continue;
      }
      throw new FolioResolvedStorySerializationError({
        message: "Resolved document story did not persist to the serialized DOCX.",
        story,
        mismatches,
        expectedBlockCount: blocks.length,
        actualBlockCount: serializedBlocks?.length ?? null,
        remainingChangeCount: serialized?.changes.length ?? null,
      });
    }
  }

  /**
   * Attempt the selective patch, treating a throw the same as a decline. Odd
   * source XML can make the paragraph diff throw rather than return `null`; a
   * full repack is the correct lossless fallback in both cases, so the fallback
   * is the graceful handling — no separate error surface is needed here.
   */
  private async trySelectiveSave(
    document: Document,
    changedParaIds: Set<string>,
  ): Promise<ArrayBuffer | null> {
    try {
      return await attemptSelectiveSave(document, this.originalBuffer, {
        changedParaIds,
        structuralChange: false,
        hasUntrackedChanges: false,
      });
    } catch {
      return null;
    }
  }

  private getEditableStoryState(story: FolioEditableDocumentStoryHandle): EditorState | null {
    if (story.type === "main") {
      return this.state;
    }
    const key = secondaryStoryKey(story);
    const existing = this.secondaryStoryStates.get(key);
    if (existing) {
      return existing.state;
    }
    const source =
      story.type === "header" || story.type === "footer"
        ? this.getHeaderFooterStory(story)
        : this.getNoteStory(story);
    if (!source) {
      return null;
    }
    const conversionOptions = {
      ...(this.baseDocument.package.styles !== undefined && {
        styles: this.baseDocument.package.styles,
      }),
      ...(this.baseDocument.package.theme !== undefined && {
        theme: this.baseDocument.package.theme,
      }),
    };
    const storyDoc =
      story.type === "header" || story.type === "footer"
        ? headerFooterToProseDoc(source.content, conversionOptions)
        : footnoteToProseDoc(source.content, conversionOptions);
    const state = ensureBaseDirectionInState(
      EditorState.create({
        schema,
        doc: ensureDeterministicParaIdsInDoc(storyDoc),
        plugins: createHeadlessPlugins(
          this.baseDocument.package.styles,
          this.baseDocument.package.numbering,
        ),
      }),
    );
    this.secondaryStoryStates.set(key, { handle: story, initialState: state, state });
    return state;
  }

  private requireEditableStoryState(story: FolioEditableDocumentStoryHandle): EditorState {
    const state = this.getEditableStoryState(story);
    if (state) {
      return state;
    }
    throw new FolioDocumentStoryNotFoundError({
      message: `Document story ${JSON.stringify(story)} was not found.`,
      story,
    });
  }

  private setEditableStoryState(story: FolioEditableDocumentStoryHandle, state: EditorState): void {
    this.resolvedStoryExpectations.delete(editableStoryKey(story));
    if (story.type === "main") {
      this.state = state;
      return;
    }
    const entry = this.secondaryStoryStates.get(secondaryStoryKey(story));
    if (!entry) {
      throw new FolioDocumentStoryNotFoundError({
        message: `Document story ${JSON.stringify(story)} was not found.`,
        story,
      });
    }
    entry.state = state;
  }

  private getHeaderFooterStory(story: FolioHeaderFooterStoryHandle): HeaderFooter | undefined {
    if (this.removedHeaderFooterStories.has(headerFooterStoryKey(story))) return undefined;
    const imported = story.type === "header" ? this.importedHeaders : this.importedFooters;
    const importedStory = imported.get(story.relationshipId);
    if (importedStory) return importedStory;
    const stories =
      story.type === "header"
        ? this.baseDocument.package.headers
        : this.baseDocument.package.footers;
    return stories?.get(story.relationshipId);
  }

  private async createComparisonHeaderFooter(
    source: FolioDocxReviewer,
    story: FolioHeaderFooterStoryHandle,
  ): Promise<FolioHeaderFooterStoryHandle | null> {
    const target = source.getHeaderFooterStory(story);
    if (!target || !this.canImportHeaderFooterContent(source, story)) return null;
    let importedWatermark: Pick<
      HeaderFooter,
      "watermark" | "rawWatermarkXml" | "watermarkBlockIndex"
    > = {};
    let importedMedia: MediaFile | undefined;
    if (target.watermark !== undefined || target.rawWatermarkXml !== undefined) {
      const watermark = target.watermark;
      if (
        story.type !== "header" ||
        watermark?.kind !== "picture" ||
        watermark.imageTargetExternal ||
        !watermark.imageTarget ||
        !target.rawWatermarkXml
      )
        return null;
      const media = source.baseDocument.package.media?.get(watermark.imageTarget);
      if (!media || !/^image\/(?:png|jpeg|gif|tiff|bmp)$/u.test(media.mimeType)) return null;
      const xml = rebindDrawingImageRelationship({
        xml: target.rawWatermarkXml,
        previousId: watermark.imageRId,
        nextId: watermark.imageRId,
      });
      if (xml === null) return null;
      const extension = media.path.split(".").at(-1)?.toLowerCase();
      if (!extension || !/^[a-z0-9]+$/u.test(extension)) return null;
      const originalZip = await JSZip.loadAsync(this.originalBuffer);
      const existingPaths = new Set(
        Object.keys(originalZip.files).map((path) => path.toLowerCase()),
      );
      let suffix = 1;
      let path = `word/media/folio-import-${suffix}.${extension}`;
      while (
        existingPaths.has(path.toLowerCase()) ||
        this.baseDocument.package.media?.has(path) ||
        this.importedMedia.has(path)
      ) {
        path = `word/media/folio-import-${++suffix}.${extension}`;
      }
      importedMedia = { ...media, path, data: media.data.slice(0) };
      importedWatermark = {
        watermark: { ...watermark, imageTarget: path },
        rawWatermarkXml: xml,
        ...(target.watermarkBlockIndex !== undefined && {
          watermarkBlockIndex: target.watermarkBlockIndex,
        }),
      };
    }
    const taken = new Set([
      ...(this.baseDocument.package.relationships?.keys() ?? []),
      ...(this.baseDocument.package.headers?.keys() ?? []),
      ...(this.baseDocument.package.footers?.keys() ?? []),
      ...this.importedHeaders.keys(),
      ...this.importedFooters.keys(),
    ]);
    let suffix = 1;
    let relationshipId = `rId${suffix}`;
    while (taken.has(relationshipId)) relationshipId = `rId${++suffix}`;
    const handle = { type: story.type, relationshipId };
    const empty: HeaderFooter = {
      type: story.type,
      hdrFtrType: target.hdrFtrType,
      ...importedWatermark,
      content: [{ type: "paragraph", content: [] }],
    };
    if (importedMedia) this.importedMedia.set(importedMedia.path, importedMedia);
    (story.type === "header" ? this.importedHeaders : this.importedFooters).set(
      relationshipId,
      empty,
    );
    const state = ensureBaseDirectionInState(
      EditorState.create({
        schema,
        doc: ensureDeterministicParaIdsInDoc(headerFooterToProseDoc(empty.content)),
        plugins: createHeadlessPlugins(
          this.baseDocument.package.styles,
          this.baseDocument.package.numbering,
        ),
      }),
    );
    this.secondaryStoryStates.set(headerFooterStoryKey(handle), {
      handle,
      initialState: state,
      state,
    });
    return handle;
  }

  private canImportHeaderFooterContent(
    source: FolioDocxReviewer,
    story: FolioHeaderFooterStoryHandle,
  ): boolean {
    const state = source.getEditableStoryState(story);
    if (!state) return false;
    if (!tableTemplateCanCrossPackageLosslessly(state.doc, "rebind-external")) return false;
    const snapshot = source.snapshotStory(story);
    return snapshot !== null && detachFolioAIEditSnapshotExternalHyperlinks(snapshot) !== null;
  }

  private getNoteStory(story: FolioNoteStoryHandle): Footnote | Endnote | undefined {
    if (story.type === "footnote") {
      return this.baseDocument.package.footnotes?.find(
        (note) => note.id === story.noteId && !isSeparatorFootnote(note),
      );
    }
    return this.baseDocument.package.endnotes?.find(
      (note) => note.id === story.noteId && !isSeparatorEndnote(note),
    );
  }

  /**
   * A loaded story's blocks, so its text comes from the same walk an unloaded
   * one uses.
   *
   * The two used to be separate walks and disagreed: the model walk separates
   * paragraphs, joins table cells with a tab and reads the accepted
   * tracked-change view, while the editor walk concatenated text nodes and saw
   * none of that. The same note therefore read one way before it was loaded and
   * another after. This is the conversion the save path already performs on the
   * same states, so the text now describes exactly what a save would write.
   */
  private storyBlocks(state: EditorState | undefined, source: { content: BlockContent[] }) {
    return state
      ? proseDocToBlocks(state.doc, source.content, this.baseDocument.package.styles, {
          // A read reports the result the document holds; a save keeps the
          // visible fallback a result-less PAGE/NUMPAGES field serializes with.
          emptyFieldResult: "authored",
        })
      : source.content;
  }

  private getHeaderFooterStoryText(
    story: FolioHeaderFooterStoryHandle,
    source: HeaderFooter,
  ): string {
    const state = this.secondaryStoryStates.get(headerFooterStoryKey(story))?.state;
    return normalizeFolioAIBlockText(
      getHeaderFooterText({ ...source, content: this.storyBlocks(state, source) }),
    );
  }

  private getNoteStoryText(story: FolioNoteStoryHandle, source: Footnote | Endnote): string {
    const state = this.secondaryStoryStates.get(noteStoryKey(story))?.state;
    const content = this.storyBlocks(state, source);
    return normalizeFolioAIBlockText(
      source.type === "footnote"
        ? getFootnoteText({ ...source, content })
        : getEndnoteText({ ...source, content }),
    );
  }

  private getStoryText(story: FolioDocumentStoryHandle): string {
    if (story.type === "main") {
      return this.getContentAsText();
    }
    if (story.type === "header" || story.type === "footer") {
      const source = this.getHeaderFooterStory(story);
      if (!source) {
        return panic("A listed document story no longer exists", { story });
      }
      return this.getHeaderFooterStoryText(story, source);
    }
    const source = this.getNoteStory(story);
    if (!source) {
      return panic("A listed document story no longer exists", { story });
    }
    return this.getNoteStoryText(story, source);
  }

  private mergeEditedSecondaryStories(
    document: Document,
    secondaryStoryStates: Iterable<FolioSecondaryStoryState> = this.secondaryStoryStates.values(),
  ): void {
    let headers: Map<string, HeaderFooter> | undefined;
    let footers: Map<string, HeaderFooter> | undefined;
    let footnotes: Footnote[] | undefined;
    let endnotes: Endnote[] | undefined;
    for (const entry of secondaryStoryStates) {
      if (entry.state === entry.initialState) {
        continue;
      }
      if (entry.handle.type === "header" || entry.handle.type === "footer") {
        const source =
          entry.handle.type === "header"
            ? document.package.headers?.get(entry.handle.relationshipId)
            : document.package.footers?.get(entry.handle.relationshipId);
        if (!source) {
          continue;
        }
        const edited = {
          ...source,
          content: proseDocToBlocks(entry.state.doc, source.content, document.package.styles),
        };
        if (entry.handle.type === "header") {
          headers ??= new Map(document.package.headers);
          headers.set(entry.handle.relationshipId, edited);
          continue;
        }
        footers ??= new Map(document.package.footers);
        footers.set(entry.handle.relationshipId, edited);
        continue;
      }
      if (entry.handle.type === "footnote") {
        const { noteId } = entry.handle;
        const source = this.baseDocument.package.footnotes?.find(
          (note) => note.id === noteId && !isSeparatorFootnote(note),
        );
        if (!source) {
          continue;
        }
        const edited = {
          ...source,
          content: proseDocToBlocks(entry.state.doc, source.content, document.package.styles),
        };
        footnotes ??= [...(document.package.footnotes ?? [])];
        const index = footnotes.findIndex((note) => note.id === noteId);
        if (index !== -1) {
          footnotes[index] = edited;
        }
        continue;
      }
      const { noteId } = entry.handle;
      const source = this.baseDocument.package.endnotes?.find(
        (note) => note.id === noteId && !isSeparatorEndnote(note),
      );
      if (!source) {
        continue;
      }
      const edited = {
        ...source,
        content: proseDocToBlocks(entry.state.doc, source.content, document.package.styles),
      };
      endnotes ??= [...(document.package.endnotes ?? [])];
      const index = endnotes.findIndex((note) => note.id === noteId);
      if (index !== -1) {
        endnotes[index] = edited;
      }
    }
    if (headers) {
      document.package.headers = headers;
    }
    if (footers) {
      document.package.footers = footers;
    }
    if (footnotes) {
      document.package.footnotes = footnotes;
    }
    if (endnotes) {
      document.package.endnotes = endnotes;
    }
  }

  /** Apply any {@link resolveComment} overrides recorded for these comments. */
  private withResolvedOverrides(
    comments: readonly Comment[],
    resolvedOverrides: ReadonlyMap<number, boolean> = this.resolvedOverrides,
  ): Comment[] {
    if (resolvedOverrides.size === 0) {
      return [...comments];
    }
    return comments.map((comment) => {
      const override = resolvedOverrides.get(comment.id);
      return override === undefined ? comment : { ...comment, done: override };
    });
  }

  /** Map each anchored comment id to its anchored text and containing block id. */
  private commentAnchors(): Map<number, { text: string; blockId: string | null }> {
    const anchors = new Map<number, { text: string; blockId: string | null }>();
    for (const anchor of getCommentAnchorsFromDoc(this.state.doc)) {
      anchors.set(anchor.commentId, { text: anchor.quote, blockId: anchor.blockId });
    }
    return anchors;
  }

  /** Allocate outside every parsed or newly-created comment and reply id. */
  private nextCommentId(): number {
    while (this.usedCommentIds.has(commentIdCursor)) {
      commentIdCursor += 1;
    }
    const id = commentIdCursor++;
    this.usedCommentIds.add(id);
    return id;
  }

  /**
   * Drive a ProseMirror command against the reviewer's headless state via the
   * same `{ state, dispatch }` seam {@link applyOperations} uses, retaining the
   * resulting state for {@link toBuffer}.
   */
  private runCommand(command: Command): boolean {
    return this.runStoryCommand(command, MAIN_STORY);
  }

  private runStoryCommand(command: Command, story: FolioEditableDocumentStoryHandle): boolean {
    const state = this.getEditableStoryState(story);
    if (!state) {
      return false;
    }
    const view = {
      state,
      dispatch: (transaction: Transaction) => {
        view.state = view.state.apply(transaction);
      },
    };
    const handled = command(view.state, view.dispatch);
    this.setEditableStoryState(story, view.state);
    return handled;
  }
}

/** @internal Comparison-only access to atomic and fresh reviewer projections. */
export const getFolioDocxComparisonAccess = (
  reviewer: FolioDocxReviewer,
): FolioDocxComparisonAccess =>
  comparisonAccessByReviewer.get(reviewer) ??
  panic("Comparison access was requested for an uninitialized DOCX reviewer");

/** Options for {@link applyFolioAIEditsToBuffer}. */
export type ApplyFolioAIEditsToBufferOptions = {
  /** Default author for tracked changes and comments. (default: `"AI"`) */
  author?: string;
  /** `"tracked-changes"` (default) produces ins/del redlines; `"direct"` edits in place. */
  mode?: FolioAIEditApplyMode;
  /**
   * The snapshot the `operations`' block ids were built against. Omit to
   * snapshot the freshly parsed buffer. Pass it when the ids were derived in a
   * separate process (e.g. server-side `@stll/folio-core/server` block ids) so
   * anchor resolution matches exactly.
   */
  snapshot?: FolioAIEditSnapshot;
};

/** Result of {@link applyFolioAIEditsToBuffer}. */
export type ApplyFolioAIEditsToBufferResult = FolioAIEditApplyResult & {
  /** The reviewed `.docx` as a new buffer. */
  buffer: ArrayBuffer;
};

/**
 * One-shot headless review: parse `buffer`, apply `operations`, and return the
 * reviewed `.docx` buffer alongside the applied / skipped breakdown. Convenience
 * wrapper over {@link FolioDocxReviewer} for callers that already hold the
 * operations to run.
 */
export const applyFolioAIEditsToBuffer = async (
  buffer: ArrayBuffer,
  operations: FolioAIEditOperation[],
  options: ApplyFolioAIEditsToBufferOptions = {},
): Promise<ApplyFolioAIEditsToBufferResult> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(buffer, {
    ...(options.author !== undefined && { author: options.author }),
  });
  const { applied, skipped } = reviewer.applyOperations(operations, {
    ...(options.mode !== undefined && { mode: options.mode }),
    ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
  });
  const reviewed = await reviewer.toBuffer();
  return { buffer: reviewed, applied, skipped };
};
