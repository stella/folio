import {
  panic,
  type Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, Selection, TextSelection } from "prosemirror-state";

import {
  CONTENT_CONTROL_LOCK,
  lockedControlMessage,
  type LockedControlViolation,
  numberingLevelEntries,
} from "@stll/docx-core/model";
import {
  captureStoryParts,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  freshCommentId,
  OP_STORIES,
  type OpStory,
  packageResourcesOpOf,
  pendingSuggestionIdsIn,
  reconcileDocumentPartRelationships,
  resolveSectionChange,
  resolveSectionChanges,
  sectionResolutionPatch,
  type SequencedBatch,
  type StoryParts,
  withoutPendingSuggestions,
} from "@stll/docx-core/ops";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  canonicalCommentBody,
  type CanonicalCommentCommand,
  compileCanonicalComments,
} from "../canonical/canonicalComments";
import { createCanonicalSectionPropertiesOperation } from "../canonical/canonicalOperations";
import { CanonicalPublicOperations } from "../canonical/canonicalPublicOperations";
import { canonicalRevisionLocks } from "../canonical/canonicalRevisionLocks";
import {
  type CanonicalRoomBase,
  openCanonicalDocxSession,
  openCanonicalRoomBase,
} from "../canonical/canonicalRoomBase";
import {
  type CanonicalCommit,
  type CanonicalSession,
  CanonicalSessionError,
  createCanonicalSession,
  publishCanonicalProjection,
} from "../canonical/canonicalSession";
import { type MatchInlineAtomsOptions } from "../compare/inline-atoms";
import { matchInlineAtomsOnModel } from "../compare/inline-atoms-model";
import {
  type InlineProvenanceTargetOptions,
  planInlineProvenance,
} from "../compare/inline-provenance";
import { inlineProvenanceOperations } from "../compare/inline-provenance-model";
import { stageSectionBoundaryOperations } from "../compare/section-boundary-properties";
import {
  importReferencedStyleDefinitions,
  type ImportReferencedStyleDefinitionsResult,
  styleClosureNumberingReferences,
} from "../compare/style-resources";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationBatchInput,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationUndoHandle,
  type FolioDocumentOperationUndoResult,
  parseFolioDocumentOperationBatch,
} from "../document-operations";
import {
  CanonicalSaveRepackRefusedError,
  serializeCanonicalSave,
} from "../docx/canonicalSave";
/**
 * Headless `.docx` review path: buffer -> apply AI edits -> buffer, with
 * no `EditorView` and no DOM. A queue worker or agent can read a document,
 * apply `FolioAIEditOperation`s (tracked-changes or direct), and write a
 * reviewed `.docx` back out — the server-side counterpart to the React
 * editor's live apply flow.
 *
 * The editor and reviewer share the canonical operation compiler and document
 * session for anchor resolution, tracked changes, and save snapshots.
 *
 * Save mirrors the editor's own path: a selective patch of only the changed
 * paragraphs in `document.xml` (leaving every untouched part byte-exact), with
 * a full repack as the fallback for structural edits.
 *
 * Scope: main, header, footer, footnote, and endnote blocks.
 */
import { rebindDrawingImageRelationship } from "../docx/drawingRelationships";
import {
  isSeparatorEndnote,
  isSeparatorFootnote,
} from "../docx/footnoteParser";
import {
  createNumberingIdAllocator,
  numberingNumIdsTaken,
} from "../docx/numberingIds";
import { getCachedNumberingMap } from "../docx/numberingParser";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import {
  getEndnoteText,
  getFootnoteText,
  getHeaderFooterText,
} from "../docx/storyPlainText";
import { expectCharacterStyleMarkAttrs } from "../prosemirror/attrs";
import {
  getSuggestions,
  suggestionIdOfRevision,
  suggestionOwnersByRevision,
} from "../prosemirror/commands/comments";
import {
  type CommentIdAllocator,
  createCommentIdAllocator,
  seedCommentAllocator,
} from "../prosemirror/commentIdAllocator";
import { commentThreadIds } from "../prosemirror/commentRecords";
import { paragraphListReferences } from "../prosemirror/listInstanceReferences";
import { MAX_LIST_LEVEL } from "../prosemirror/listMarker";
import {
  type NoteKey,
  noteReferenceStates,
} from "../prosemirror/noteReferenceReview";
import { ensureParaIdsInDoc } from "../prosemirror/paraIds";
import {
  createDocumentNumberingPlugin,
  getDocumentNumbering,
} from "../prosemirror/plugins/documentNumbering";
import {
  createDocumentStyleContextPlugin,
  getDocumentStyleResolver,
} from "../prosemirror/plugins/documentStyles";
import { getFolioNodeRevisionCarriers } from "../prosemirror/revisionCarriers";
import { schema } from "../prosemirror/schema";
import { sectionPropertiesOf } from "../prosemirror/sectionCarrier";
import {
  CANONICAL_GAP,
  CANONICAL_RULE,
  type HeadlessReviewerSession,
} from "../types/canonicalCapabilities";
import type { Comment } from "../types/content";
import type {
  Document,
  Endnote,
  Footnote,
  HeaderFooter,
  HeaderFooterType,
  MediaFile,
  NumberingDefinitions,
  SectionProperties,
  StyleDefinitions,
} from "../types/document";
import { readModel, type ReadonlyDocument } from "../types/documentModel";
import { canonicalJson } from "../utils/canonicalJson";
import type {
  FolioReplacementBackground,
  FolioRevisionStamp,
  FolioWordDiffOptions,
} from "./applyTypes";
import { buildAnnotatedBlockTextWithNoteReferences } from "./clean-text";
import {
  anchoredCommentIdsInBlocks,
  anchoredCommentIdsInProseDoc,
  coveredCommentIdsInProseDoc,
  withoutLostCommentThreads,
} from "./comment-lifecycle";
import { stagedFinalSectionProperties } from "./finalSectionReview";
import {
  createPendingSuggestionSourceSnapshot,
  type FolioPendingSuggestionLoadResult,
  type FolioPendingSuggestionRecord,
  FolioPendingSuggestionRegistry,
} from "./pending-suggestions";
import {
  type FolioAbsorbedBlock,
  type FolioReviewChange,
  type FolioReviewChangeKind,
  type FolioReviewResolution,
  type FolioReviewResolutionSkip,
  getCommentAnchorsFromDoc,
  getTrackedChangesFromDoc,
  getTrackedChangesFromSnapshot,
  getTrackedChangeStatsFromDoc,
  listedRevisionGroups,
} from "./read";
import { classifyResourceReference } from "./referenceClassification";
import {
  collectNoteReferenceLabels,
  createReviewerSnapshot,
  detachFolioAIEditSnapshotExternalHyperlinks,
  type FolioStoryTable,
  folioStoryTables,
  isFolioAIContentBlock,
  normalizeFolioAIBlockText,
  sourceDocumentOf,
} from "./snapshot";
import {
  matchTableGeometryOperations,
  type TableGeometryPairing,
} from "./table-geometry";
import {
  type FolioTableTemplates,
  tableTemplateCanCrossPackageLosslessly,
} from "./table-template";
import type {
  FolioAIBlock,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "./types";
import type { FolioUndefinedReferencePolicy } from "./undefinedReferences";

/**
 * How many document-operation batches a reviewer can step back through.
 * Each entry holds the story's state before and after its batch; older
 * entries are dropped, as an editor's undo history drops its oldest steps.
 */
export const FOLIO_DOCUMENT_OPERATION_UNDO_LIMIT = 100;

/** Options for {@link FolioDocxReviewer.fromBuffer}. */
export type FolioDocxReviewerOptions = {
  /** Default author for tracked changes and comments. (default: `"AI"`) */
  author?: string;
  /** Password for Agile-encrypted .docx files (Office 2010+). */
  password?: string | undefined;
  /**
   * The engine the document is held in: the canonical session the editors
   * run. Every read, operation, undo, resolution, resolved view and save runs
   * on it; what the session cannot express throws a `CanonicalSessionError`.
   */
  session?: HeadlessReviewerSession;
};

/**
 * Where a reviewer's document lives. The reads are one implementation over
 * either: a canonical reviewer's main state holds the session's projection,
 * exactly the document a canonical editor shows, and its model is the
 * session's document.
 */
type FolioReviewerEngineSource = {
  type: "canonical";
  session: CanonicalSession;
};

/**
 * A canonical reviewer's engine: the session, and the editor's public
 * operation compiler over it, which publishes into the reviewer's state.
 */
type FolioReviewerEngine = {
  type: "canonical";
  session: CanonicalSession;
  operations: CanonicalPublicOperations;
};

/**
 * The main story's state of a canonical reviewer: the session's projection
 * with the package's numbering and styles, which the readers resolve
 * against. It is rebuilt from the session after every commit, so it never
 * holds anything the session does not.
 */
const canonicalReviewerState = (session: CanonicalSession): EditorState =>
  canonicalStoryState(session, session.projection.doc, session.projection);

/** The projection a story state reads its addresses from. */
type AddressedProjection = Pick<CanonicalSession["projection"], "addressAt">;

/**
 * The caret a story state starts at: the start of the story, or where the
 * session first holds an address when it holds none there, as in the filler
 * paragraph of a content control the source leaves empty.
 */
const startingSelection = (
  doc: PMNode,
  projection: AddressedProjection | undefined,
): Selection | undefined => {
  const start = Selection.atStart(doc);
  if (projection === undefined || projection.addressAt(start.from).isOk())
    return undefined;
  let found: number | undefined;
  doc.descendants((node, position) => {
    if (found !== undefined) return false;
    if (!node.isTextblock) return true;
    if (projection.addressAt(position + 1).isOk()) found = position + 1;
    return false;
  });
  return found === undefined ? undefined : TextSelection.create(doc, found);
};

/** A story's state over `doc`, with the session document's numbering and styles. */
const canonicalStoryState = (
  session: CanonicalSession,
  doc: PMNode,
  projection?: AddressedProjection,
): EditorState => {
  const document = readModel(session.document);
  const selection = startingSelection(doc, projection);
  return EditorState.create({
    schema,
    doc,
    ...(selection === undefined ? {} : { selection }),
    plugins: [
      createDocumentNumberingPlugin(document.package.numbering),
      // The session's projection owns formatting: the style context gives
      // the readers the resolver and appends no repair transaction.
      createDocumentStyleContextPlugin(document.package.styles),
    ],
  });
};

/** The session's name for a reviewer's story. */
const canonicalOpStory = (story: FolioEditableDocumentStoryHandle): OpStory => {
  switch (story.type) {
    case "main":
      return OP_STORIES.MAIN;
    case "header":
    case "footer":
      return { kind: story.type, rId: story.relationshipId };
    case "footnote":
    case "endnote":
      return { kind: story.type, id: story.noteId };
    default: {
      const unreachable: never = story;
      return panic("Unknown reviewer story", { story: unreachable });
    }
  }
};

/** The refusal of a canonical reviewer for what it cannot do on the session yet. */
const canonicalReviewerRefusal = (action: string): CanonicalSessionError =>
  new CanonicalSessionError({
    ground: { gap: CANONICAL_GAP.publicHeadlessSession },
    reason: "refused",
    message: `A canonical headless reviewer cannot ${action} yet.`,
  });

/** Options for {@link FolioDocxReviewer.applyOperations}. */
export type FolioApplyOperationsOptions = {
  /** Required policy for undefined style and numbering references. */
  undefinedReferences: FolioUndefinedReferencePolicy;
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
export type FolioApplyDocumentOperationsOptions = Omit<
  FolioApplyOperationsOptions,
  "mode"
>;

/**
 * Why a save rewrote the whole package instead of patching the edited
 * paragraphs into the original: `structuralChange` (paragraphs added or
 * removed, or styles, headers, footers or final section properties staged),
 * `untrackedChange` (an edit the change tracker cannot key to a paragraph), or
 * `selectiveDeclined` (the patch could not express the edit safely).
 */
export type FolioReviewerRepackReason =
  "structuralChange" | "untrackedChange" | "selectiveDeclined";

/** Options for {@link FolioDocxReviewer.save}. */
export type FolioReviewerSaveOptions = {
  /**
   * `"allow"` (default) falls back to a full repack when the selective patch
   * does not apply; `"refuse"` reports the reason and writes nothing.
   */
  repack?: "allow" | "refuse";
};

/** Result of {@link FolioDocxReviewer.save}. */
export type FolioReviewerSaveResult =
  | { type: "selective"; buffer: ArrayBuffer }
  | {
      type: "full-repack";
      reason: FolioReviewerRepackReason;
      buffer: ArrayBuffer;
    }
  | { type: "repackRefused"; reason: FolioReviewerRepackReason };

/** Options for {@link FolioDocxReviewer.getContentAsText}. */
export type FolioGetContentAsTextOptions = {
  /**
   * Render tracked changes and comment anchors inline as `<ins>` / `<del>` /
   * `<comment>` tags instead of the default flattened, post-tracked-changes
   * text. (default: `false`)
   */
  annotated?: boolean;
};

export const FOLIO_REVIEWED_VIEWS: readonly [
  "original",
  "current-markup",
  "final",
] = Object.freeze(["original", "current-markup", "final"] as const);

export type FolioReviewedView = (typeof FOLIO_REVIEWED_VIEWS)[number];

export const FOLIO_RESOLVED_REVIEWED_VIEWS: readonly ["original", "final"] =
  Object.freeze(["original", "final"] as const);

export type FolioResolvedReviewedView =
  (typeof FOLIO_RESOLVED_REVIEWED_VIEWS)[number];

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

const UnsupportedFolioReviewedViewErrorBase: TaggedErrorClass<"UnsupportedFolioReviewedViewError"> =
  TaggedError("UnsupportedFolioReviewedViewError");

export class UnsupportedFolioReviewedViewError extends UnsupportedFolioReviewedViewErrorBase<{
  message: string;
  receivedView: unknown;
}> {}

const FolioRoomJournalErrorBase: TaggedErrorClass<"FolioRoomJournalError"> =
  TaggedError("FolioRoomJournalError");

/**
 * Why a reviewer could not replay a room's journal. `revisionGap`: a batch is
 * not the one after the last the reviewer applied.
 */
export class FolioRoomJournalError extends FolioRoomJournalErrorBase<{
  message: string;
  reason: "revisionGap";
  expectedRevision: number;
  receivedRevision: number | undefined;
}> {}

const FolioDocumentStoryNotFoundErrorBase: TaggedErrorClass<"FolioDocumentStoryNotFoundError"> =
  TaggedError("FolioDocumentStoryNotFoundError");

export class FolioDocumentStoryNotFoundError extends FolioDocumentStoryNotFoundErrorBase<{
  message: string;
  story: FolioEditableDocumentStoryHandle;
}> {}

export type FolioApplyDocumentOperationsToStoryOptions =
  FolioApplyDocumentOperationsOptions & {
    story: FolioEditableDocumentStoryHandle;
    batch: FolioDocumentOperationBatchInput;
    /**
     * Tables and rows an `insertTable` / `insertTableRow` in the batch places
     * verbatim, by operation id. A caller copying a table it already holds —
     * `compareDocx` placing the target document's table — hands it over whole
     * instead of letting the operation rebuild it from its cell texts and lose
     * every table, row and cell property on the way.
     */
    tableTemplates?: FolioTableTemplates;
    /**
     * What a replacement in the batch does with a highlight or `w:shd` the text
     * it replaces carries. Clearing it is the authoring default; a caller that
     * states the replacement's own run properties — `compareDocx` reproducing
     * the target document — keeps them instead.
     */
    replacementBackground?: FolioReplacementBackground;
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

type FolioSecondaryStoryHandle = Exclude<
  FolioEditableDocumentStoryHandle,
  { type: "main" }
>;

type ApplyDocumentOperationsInternalOptions = {
  story: FolioEditableDocumentStoryHandle;
  batch: FolioDocumentOperationBatch;
  author?: string;
  snapshot?: FolioAIEditSnapshot;
  revisionStamp?: FolioRevisionStamp;
  wordDiff?: FolioWordDiffOptions;
  tableTemplates?: FolioTableTemplates;
  replacementBackground?: FolioReplacementBackground;
  undefinedReferences: FolioUndefinedReferencePolicy;
  createUndoEntry: boolean;
  /** The batch restages suggestions that already carry their ids. */
  restagesSuggestions?: boolean;
};

type FolioDocxComparisonProjectionMode =
  "with-revision-census" | "without-revision-census";

type FolioDocxComparisonStoryProjection = {
  handle: FolioDocumentStoryHandle;
  snapshot: FolioAIEditSnapshot | null;
};

type FolioDocxComparisonProjection = {
  stories: readonly FolioDocxComparisonStoryProjection[];
  revisions: {
    ids: readonly number[];
    present: boolean;
  };
};

type MatchStoryInlineProvenanceOptions = InlineProvenanceTargetOptions & {
  story: FolioEditableDocumentStoryHandle;
  targetReviewer: FolioDocxReviewer;
  targetStory: FolioEditableDocumentStoryHandle;
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

type MatchStoryInlineAtomsOptions = Omit<
  MatchInlineAtomsOptions,
  "state" | "author"
> & {
  story: FolioEditableDocumentStoryHandle;
  targetReviewer: FolioDocxReviewer;
  targetStory: FolioEditableDocumentStoryHandle;
};

type StoryInlineAtomsResult =
  | {
      status: "matched";
      nextRevisionId: number;
      changedTargetBlockIds: readonly string[];
      rangeCount: number;
      documentChanged: boolean;
    }
  | { status: "unalignable"; detail?: string }
  | { status: "budget-exceeded" };

type StageTargetNumberingResult = "unchanged" | "staged" | "conflict";

const numberingLevelsOf = (
  numbering: NumberingDefinitions | null | undefined,
): FolioNumberingLevel[] => {
  if (!numbering) return [];
  return numberingLevelEntries(numbering).map(({ numId, ilvl, level }) => {
    const entry: FolioNumberingLevel = {
      numId,
      level: ilvl,
      format: level.numFmt,
      levelText: level.lvlText,
    };
    if (level.start !== undefined) entry.start = level.start;
    return entry;
  });
};

type ReferencedNumberingLevelsOptions = {
  references: readonly { numId: number; level: number }[];
};

const referencedNumberingLevelsByNumId = ({
  references,
}: ReferencedNumberingLevelsOptions): ReadonlyMap<
  number,
  readonly number[]
> => {
  const levelsByNumId = new Map<number, Set<number>>();
  for (const { numId, level } of references) {
    const levels = levelsByNumId.get(numId) ?? new Set<number>();
    levelsByNumId.set(numId, levels);
    levels.add(level);
    for (
      let ancestor = 0;
      ancestor <= Math.min(level, MAX_LIST_LEVEL);
      ancestor += 1
    ) {
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
    // A level neither side defines — a `w:numPr` naming an `w:abstractNum`
    // with no such `w:ilvl` — is the same absence on both, not a difference.
    // Reading it as one restamped the paragraph onto a freshly minted `numId`,
    // so a document compared with itself reported a numbering change and the
    // direct `w:ind` the new numbering displaced could no longer be moved.
    if (currentLevel === null && targetLevel === null) {
      continue;
    }
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

type StageTargetStylesOptions = {
  source: FolioDocxReviewer;
  snapshots: readonly FolioAIEditSnapshot[];
  importedHeaderFooterSnapshots: readonly FolioAIEditSnapshot[];
  numberingReferenceMap?: ReadonlyMap<number, number>;
};

type FolioDocxComparisonAccess = {
  stageTerminalTableReviewCarrier: (target: PMNode) => boolean;
  stageTargetStyles: (
    options: StageTargetStylesOptions,
  ) => ImportReferencedStyleDefinitionsResult;
  createComparisonHeaderFooter: (
    source: FolioDocxReviewer,
    story: FolioHeaderFooterStoryHandle,
  ) => Promise<FolioHeaderFooterStoryHandle | null>;
  matchInlineAtoms: (
    options: MatchStoryInlineAtomsOptions,
  ) => StoryInlineAtomsResult;
  matchInlineProvenance: (
    options: MatchStoryInlineProvenanceOptions,
  ) => StoryInlineProvenanceResult;
  projectStories: (
    mode: FolioDocxComparisonProjectionMode,
  ) => FolioDocxComparisonProjection;
  snapshotReviewedStory: (
    options?: FolioReadReviewedStoryOptions,
  ) => FolioAIEditSnapshot | null;
  numberingDefinitions: () => NumberingDefinitions | null | undefined;
  styleDefinitions: () => Document["package"]["styles"];
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
      | {
          kind: "retained";
          previous: SectionProperties;
          target: SectionProperties;
        }
      | null;
  }) =>
    | {
        status: "matched";
        rangeCount: number;
        nextRevisionId: number;
        documentChanged: boolean;
      }
    | { status: "unalignable"; detail: string }
    | { status: "budget-exceeded" };
};

const comparisonAccessByReviewer = new WeakMap<
  FolioDocxReviewer,
  FolioDocxComparisonAccess
>();

const MAIN_STORY = Object.freeze({ type: "main" } as const);

const headerFooterStoryKey = ({
  type,
  relationshipId,
}: FolioHeaderFooterStoryHandle): string => `${type}:${relationshipId}`;

const noteStoryKey = ({ type, noteId }: FolioNoteStoryHandle): NoteKey =>
  `${type}:${noteId}`;

/** The note story a {@link noteStoryKey} key names. */
const noteStoryHandleOf = (key: string): FolioNoteStoryHandle | null => {
  const [type, noteId] = key.split(":");
  return (type === "footnote" || type === "endnote") && noteId !== undefined
    ? ({ type, noteId: Number(noteId) } as FolioNoteStoryHandle)
    : null;
};

/** One place a header or footer part is shown: its role in one section. */
type HeaderFooterPlacement = { type: HeaderFooterType; section: number };

/**
 * A part's roles as a reader label: `first`, or, in a document with several
 * sections, `first (sections 1, 3), default (section 2)`.
 */
const describePlacements = (
  placements: readonly HeaderFooterPlacement[],
  sectionCount: number,
): string => {
  const sectionsByRole = new Map<HeaderFooterType, number[]>();
  for (const { type, section } of placements) {
    sectionsByRole.set(type, [...(sectionsByRole.get(type) ?? []), section]);
  }
  return [...sectionsByRole]
    .map(([type, sections]) =>
      sectionCount > 1
        ? `${type} (${sections.length > 1 ? "sections" : "section"} ${sections.join(", ")})`
        : type,
    )
    .join(", ");
};

const secondaryStoryKey = (story: FolioSecondaryStoryHandle): string =>
  story.type === "header" || story.type === "footer"
    ? headerFooterStoryKey(story)
    : noteStoryKey(story);

const editableStoryKey = (story: FolioEditableDocumentStoryHandle): string =>
  story.type === "main" ? "main" : secondaryStoryKey(story);

export const isFolioReviewedView = (
  value: unknown,
): value is FolioReviewedView =>
  FOLIO_REVIEWED_VIEWS.some((view) => view === value);

export const isFolioResolvedReviewedView = (
  value: unknown,
): value is FolioResolvedReviewedView =>
  FOLIO_RESOLVED_REVIEWED_VIEWS.some((view) => view === value);

/** A state's snapshot, of `doc` when it names the state's document otherwise spelled. */
const createStateSnapshot = (
  state: EditorState,
  positionalParaIds: ReadonlySet<string>,
  doc: PMNode = state.doc,
): FolioAIEditSnapshot =>
  createReviewerSnapshot({
    doc,
    styleResolver: getDocumentStyleResolver(state),
    numbering: getDocumentNumbering(state),
    positionalParaIds,
  });

const formatStorySnapshotForLLM = (
  snapshot: FolioAIEditSnapshot,
  annotated: boolean,
): string => {
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
  const sourceDocument = sourceDocumentOf(snapshot);
  const noteReferences = collectNoteReferenceLabels(sourceDocument);
  sourceDocument.descendants((node, pos) => {
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
      const text = node
        ? buildAnnotatedBlockTextWithNoteReferences(node, noteReferences)
        : block.text;
      return formatBlockLine(block, text);
    })
    .join("\n");
};

// The change shape and its pure reader now live in `./read` so a live editor
// can produce the same `FolioReviewChange[]` from its own doc; the reviewer
// delegates to `getTrackedChangesFromDoc` and re-exports the types here so its
// public surface is unchanged.
export type {
  FolioAbsorbedBlock,
  FolioReviewChange,
  FolioReviewChangeKind,
  FolioReviewResolution,
  FolioReviewResolutionSkip,
} from "./read";

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
  /** ISO-8601 reply date; defaults to the wall clock. */
  date?: string;
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
  /**
   * Stable id of the anchored block, or `null` when the anchor is absent. The
   * id is the one {@link FolioDocxReviewer.snapshotStory} gives the block in
   * {@link FolioReviewComment.story}.
   */
  blockId: string | null;
  /**
   * The story holding the anchor — the body, a header or footer, a footnote or
   * an endnote — or `null` when the anchor is absent.
   */
  story: FolioDocumentStoryHandle | null;
  replies: FolioReviewCommentReply[];
  /** Whether the comment is marked resolved / done. */
  done: boolean;
};

type FolioCommentAnchor = {
  text: string;
  blockId: string | null;
  story: FolioDocumentStoryHandle;
};

/** Filter for {@link FolioDocxReviewer.getComments}. */
export type FolioReviewCommentFilter = {
  author?: string;
  done?: boolean;
};

/**
 * One LLM-ready line for a block: `[<blockId>] text`, with an `(h<level>)` tag
 * for headings, then the number a numbered heading shows, and the list marker
 * for list items, so a model can copy the block id straight back into an
 * operation and match the numbers the document shows.
 */
const formatBlockLine = (block: FolioAIBlock, text: string): string => {
  const label = `[${block.id}]`;
  if (block.kind === "heading") {
    // An unnumbered heading's `displayLabel` is its style id, not a number.
    const number =
      block.displayLabel !== undefined && block.displayLabel !== block.styleId
        ? `${block.displayLabel} `
        : "";
    return `${label} (h${headingLevel(block)}) ${number}${text}`;
  }
  if (block.kind === "listItem") {
    return `${label} ${block.displayLabel ?? "•"} ${text}`;
  }
  return `${label} ${text}`;
};

const formatBlockForLLM = (block: FolioAIBlock): string =>
  formatBlockLine(block, block.text);

/**
 * The level the snapshot already classified. Scraping a digit out of the style
 * id instead read `ClauseParagraph2` as an h2 and every localized heading id
 * without a digit as an h1.
 */
const headingLevel = (block: FolioAIBlock): number => block.headingLevel ?? 1;

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

type ReviewLockFact = {
  id: number;
  suggestionId: string | null;
  violation: LockedControlViolation | undefined;
  /** Other revision ids of its story that resolving this change removes. */
  carries: ReadonlySet<number>;
  /**
   * Other revision ids of its story that resolving this change runs into the
   * next paragraph, their text kept.
   */
  moves: ReadonlySet<number>;
};

type ClassifyReviewLocksOptions<Source, Change extends ReviewLockFact> = {
  stories: readonly { source: Source; changes: readonly Change[] }[];
  grouping: "whole-proposal" | "individual-change";
};

type HeldReviewLock = { violation: LockedControlViolation; message: string };

/**
 * Source-specific probes supply facts; a held proposal stays whole across all
 * stories, and a change whose resolution would carry a held change away (its
 * removed content or block holds it) is held with it. A change that only runs
 * a held change's paragraph into the next is held with it when the held change
 * belongs to a proposal, which is restaged on the paragraphs it was made on;
 * an ordinary held change keeps its text and record wherever its paragraph
 * goes.
 */
const classifyReviewLocks = <Source, Change extends ReviewLockFact>({
  stories,
  grouping,
}: ClassifyReviewLocksOptions<Source, Change>) => {
  switch (grouping) {
    case "whole-proposal":
    case "individual-change":
      break;
    default: {
      const unhandled: never = grouping;
      return panic("Unknown review lock grouping", { grouping: unhandled });
    }
  }
  const heldSuggestions = new Map<string, LockedControlViolation>();
  const held = stories.map(() => new Map<Change, HeldReviewLock>());
  const heldIds = stories.map(() => new Map<number, LockedControlViolation>());
  const heldProposalIds = stories.map(() => new Set<number>());
  const reasonFor = (
    change: Change,
    storyIndex: number,
  ): HeldReviewLock | undefined => {
    if (change.violation !== undefined)
      return {
        violation: change.violation,
        message: lockedControlMessage(change.violation),
      };
    const proposal =
      change.suggestionId === null
        ? undefined
        : heldSuggestions.get(change.suggestionId);
    if (proposal !== undefined)
      return {
        violation: proposal,
        message: `Another change in this proposal is held: ${lockedControlMessage(proposal)}`,
      };
    for (const id of change.carries) {
      const carried = heldIds[storyIndex]?.get(id);
      if (carried !== undefined)
        return {
          violation: carried,
          message: `Resolving this change would take a held change with it: ${lockedControlMessage(carried)}`,
        };
    }
    for (const id of change.moves) {
      if (!heldProposalIds[storyIndex]?.has(id)) continue;
      const moved = heldIds[storyIndex]?.get(id);
      if (moved !== undefined)
        return {
          violation: moved,
          message: `Resolving this change would run a held proposal's paragraph into the next: ${lockedControlMessage(moved)}`,
        };
    }
    return undefined;
  };
  // Holding one change can hold its proposal and the changes that carry it.
  for (let grew = true; grew;) {
    grew = false;
    for (const [storyIndex, { changes }] of stories.entries()) {
      const storyHeld =
        held[storyIndex] ?? panic("Every story has its held changes");
      for (const change of changes) {
        if (storyHeld.has(change)) continue;
        const reason = reasonFor(change, storyIndex);
        if (reason === undefined) continue;
        storyHeld.set(change, reason);
        heldIds[storyIndex]?.set(change.id, reason.violation);
        if (change.suggestionId !== null)
          heldProposalIds[storyIndex]?.add(change.id);
        if (
          grouping === "whole-proposal" &&
          change.suggestionId !== null &&
          !heldSuggestions.has(change.suggestionId)
        )
          heldSuggestions.set(change.suggestionId, reason.violation);
        grew = true;
      }
    }
  }
  return {
    heldSuggestionIds: new Set(heldSuggestions.keys()),
    stories: stories.map(({ source, changes }, storyIndex) => ({
      source,
      changes: changes.map((change) => {
        const reason = held[storyIndex]?.get(change);
        if (reason === undefined) return { type: "free", change } as const;
        return { type: "held", change, ...reason } as const;
      }),
    })),
  };
};

const PARAGRAPH_MARK_REMOVED_BY = {
  accept: "paragraphMarkDeleted",
  reject: "paragraphMarkInserted",
} as const;
const RECORD_REMOVED_BY = { accept: "deletion", reject: "insertion" } as const;
const ROW_REMOVED_BY = { accept: "trDel", reject: "trIns" } as const;
const CELL_REMOVED_BY = { accept: "del", reject: "ins" } as const;

const numericRevisionId = (value: unknown): number | undefined =>
  typeof value === "object" &&
  value !== null &&
  "revisionId" in value &&
  typeof value.revisionId === "number"
    ? value.revisionId
    : undefined;

/** What resolving each change takes with it, by revision id. */
type CarriedByResolution = {
  /** The records it removes. */
  removed: ReadonlyMap<number, ReadonlySet<number>>;
  /** The records it runs into the next paragraph, their text kept. */
  moved: ReadonlyMap<number, ReadonlySet<number>>;
};

/**
 * The revision ids each change's resolution takes with it: every record inside
 * content it removes (the text it inserted, when rejecting; deleted, when
 * accepting) or inside a row or cell whose own carrier it removes. Removing a
 * paragraph's mark removes the paragraph's own records and runs its content,
 * with the records in it, into the next paragraph.
 */
const revisionIdsCarriedByResolution = (
  doc: PMNode,
  mode: "accept" | "reject",
): CarriedByResolution => {
  const removed = new Map<number, Set<number>>();
  const moved = new Map<number, Set<number>>();
  const add = (
    into: Map<number, Set<number>>,
    remover: number,
    ids: ReadonlySet<number>,
  ) => {
    const set = into.get(remover) ?? new Set<number>();
    for (const id of ids) if (id !== remover) set.add(id);
    into.set(remover, set);
  };
  const visit = (node: PMNode, pos: number): Set<number> => {
    const ids = new Set<number>();
    const removers = new Set<number>();
    const markRemovers = new Set<number>();
    for (const mark of node.marks) {
      const id = numericRevisionId(mark.attrs);
      if (id !== undefined) {
        ids.add(id);
        if (mark.type.name === RECORD_REMOVED_BY[mode]) removers.add(id);
      }
      const ancestors: unknown = mark.attrs["_docxRevisionAncestors"];
      if (!Array.isArray(ancestors)) continue;
      for (const ancestor of ancestors as readonly unknown[]) {
        const ancestorId = numericRevisionId(ancestor);
        if (ancestorId === undefined) continue;
        ids.add(ancestorId);
        if (
          typeof ancestor === "object" &&
          ancestor !== null &&
          "type" in ancestor &&
          ancestor.type === RECORD_REMOVED_BY[mode]
        )
          removers.add(ancestorId);
      }
    }
    for (const carrier of getFolioNodeRevisionCarriers(node, pos)) {
      ids.add(carrier.id);
      if (carrier.type === PARAGRAPH_MARK_REMOVED_BY[mode])
        markRemovers.add(carrier.id);
    }
    if (node.type.name === "tableRow")
      for (const attr of ["trIns", "trDel"] as const) {
        const id = numericRevisionId(node.attrs[attr]);
        if (id === undefined) continue;
        ids.add(id);
        if (attr === ROW_REMOVED_BY[mode]) removers.add(id);
      }
    const cellMarker: unknown = node.attrs["cellMarker"];
    if (
      typeof cellMarker === "object" &&
      cellMarker !== null &&
      "info" in cellMarker &&
      "kind" in cellMarker
    ) {
      const id = numericRevisionId(cellMarker.info);
      if (id !== undefined) {
        ids.add(id);
        if (cellMarker.kind === CELL_REMOVED_BY[mode]) removers.add(id);
      }
    }
    const own = new Set(ids);
    const content = new Set<number>();
    node.forEach((child, offset) => {
      for (const id of visit(child, pos + 1 + offset)) {
        ids.add(id);
        content.add(id);
      }
    });
    for (const remover of removers) add(removed, remover, ids);
    for (const remover of markRemovers) {
      add(removed, remover, own);
      add(moved, remover, content);
    }
    return ids;
  };
  doc.forEach((child, offset) => {
    visit(child, offset);
  });
  return { removed, moved };
};

/**
 * What a change's group takes along: its other members and what each removes
 * (`carries`), and what each runs into the next paragraph (`moves`).
 */
const carriedByGroup = (
  id: number,
  groupIds: readonly number[],
  { removed, moved }: CarriedByResolution,
): Pick<ReviewLockFact, "carries" | "moves"> => {
  const carries = new Set<number>();
  const moves = new Set<number>();
  for (const member of [id, ...groupIds]) {
    carries.add(member);
    for (const inner of removed.get(member) ?? []) carries.add(inner);
    for (const inner of moved.get(member) ?? []) moves.add(inner);
  }
  carries.delete(id);
  for (const inner of carries) moves.delete(inner);
  moves.delete(id);
  return { carries, moves };
};

type StoryOwnNumberingOptions = {
  /** The session's numbering, with every list its operations started. */
  current: NumberingDefinitions | undefined;
  /** The numbering the package was opened with. */
  loaded: NumberingDefinitions | undefined;
  doc: PMNode;
};

/**
 * One story's reading of the numbering: the lists the package was opened
 * with, and those started since that `doc`'s paragraphs name. A list another
 * story started is defined in that story's reading, not in this one.
 */
const storyOwnNumbering = ({
  current,
  loaded,
  doc,
}: StoryOwnNumberingOptions): NumberingDefinitions | undefined => {
  if (current === undefined) return undefined;
  const kept = new Set((loaded?.nums ?? []).map(({ numId }) => numId));
  doc.descendants((node) => {
    for (const numId of paragraphListReferences(node)) kept.add(numId);
    return !node.isTextblock;
  });
  return {
    ...current,
    nums: current.nums.filter(({ numId }) => kept.has(numId)),
  };
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
 * ], { undefinedReferences: "refuse" });
 * const reviewed = await reviewer.toBuffer();
 * ```
 */

export class FolioDocxReviewer {
  /**
   * A state's snapshot. On the canonical session a paragraph with no id (the
   * filler of a content control its source leaves empty) is named by the
   * positional id the legacy engine mints for it, so both engines name it
   * alike; a legacy state already holds every id.
   */
  private snapshotOf(state: EditorState): FolioAIEditSnapshot {
    return createStateSnapshot(
      state,
      this.positionalParaIds,
      ensureParaIdsInDoc(state.doc),
    );
  }

  /** Default author for tracked changes and comments. */
  readonly author: string;
  /** The model the reviewer was opened on; the readers follow the session's. */
  private readonly loadedDocument: Document;
  private readonly originalBuffer: ArrayBuffer;
  private state: EditorState;
  /** Comments the body anchored when it was loaded, read as the editor projects it. */
  private readonly loadedMainCommentIds: ReadonlySet<number>;
  /**
   * Paragraph ids the canonical session's input minted for paragraphs the
   * package named none for: a snapshot reads them as positional, so nothing
   * pairs two documents' paragraphs by them.
   */
  private readonly positionalParaIds: ReadonlySet<string>;
  private pendingSuggestions = new FolioPendingSuggestionRegistry();
  /** The last room revision {@link applyRoomJournal} applied. */
  private roomRevision = 0;
  /** The resolved copies {@link resolvedSessionReader} last read, by view. */
  private readonly resolvedSessionReaders = new Map<
    FolioResolvedReviewedView,
    WeakRef<{ document: ReadonlyDocument; reviewer: FolioDocxReviewer }>
  >();
  private readonly engine: FolioReviewerEngine;
  /** This document's comment ids: loaded, created, and never reissued. */
  private readonly commentIds: CommentIdAllocator;

  private constructor(args: {
    baseDocument: Document;
    originalBuffer: ArrayBuffer;
    state: EditorState;
    author: string;
    engine: FolioReviewerEngineSource;
    loadedMainCommentIds?: ReadonlySet<number>;
    /** Paragraph ids minted when the document was opened. */
    positionalParaIds?: ReadonlySet<string>;
    /** The comment ids of the document this reviewer copies. */
    commentIds?: CommentIdAllocator;
  }) {
    this.loadedDocument = args.baseDocument;
    this.positionalParaIds = args.positionalParaIds ?? new Set();
    this.originalBuffer = args.originalBuffer;
    this.state = args.state;
    this.author = args.author;
    this.engine = this.canonicalEngine(args.engine.session);
    this.commentIds = args.commentIds?.clone() ?? createCommentIdAllocator();
    seedCommentAllocator(
      this.commentIds,
      args.baseDocument.package.document.comments,
      {
        state: args.state,
      },
    );
    this.loadedMainCommentIds = new Set(
      args.loadedMainCommentIds ?? anchoredCommentIdsInProseDoc(args.state.doc),
    );
    this.registerComparisonAccess();
  }

  private registerComparisonAccess(): void {
    comparisonAccessByReviewer.set(this, this.canonicalComparisonAccess());
  }

  private stageTargetStyles({
    source,
    snapshots,
    importedHeaderFooterSnapshots,
    numberingReferenceMap,
  }: StageTargetStylesOptions): ImportReferencedStyleDefinitionsResult {
    const destination = this.baseDocument.package;
    const sourcePackage = source.baseDocument.package;
    const destinationStyles = destination.styles;
    const sourceStylesById = new Map(
      sourcePackage.styles?.styles.map((style) => [style.styleId, style]),
    );
    const destinationStyleIds = new Set(
      destinationStyles?.styles.map(({ styleId }) => styleId),
    );
    const collect = (document: PMNode): Set<string> => {
      const references = new Set<string>();
      document.descendants((node) => {
        const styleId = node.attrs["styleId"];
        if (typeof styleId === "string" && styleId.length > 0)
          references.add(styleId);
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
    const resourceSnapshots = contextsMatch
      ? snapshots
      : importedHeaderFooterSnapshots;
    const referencedStyleIds = new Set<string>();
    let materializeDefaultParagraphStyle = false;
    for (const snapshot of resourceSnapshots) {
      if (hasUnstyledParagraph(sourceDocumentOf(snapshot))) {
        materializeDefaultParagraphStyle = true;
      }
      for (const styleId of collect(sourceDocumentOf(snapshot))) {
        if (
          classifyResourceReference({
            kind: "style",
            id: styleId,
            sourceDefines: (id) => sourceStylesById.has(id),
            destinationDefines: (id) => destinationStyleIds.has(id),
          }) === "sourceDefined"
        ) {
          referencedStyleIds.add(styleId);
        }
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
    // Without a numbering import, style definitions retain their source references.
    let styleNumberingReferenceMap =
      numberingReferenceMap ?? new Map<number, number>();
    if (numberingReferenceMap !== undefined) {
      const references = styleClosureNumberingReferences(sourcePackage.styles, [
        ...referencedStyleIds,
      ]);
      const remaining = references.filter(
        ({ numId }) => !numberingReferenceMap.has(numId),
      );
      const planned = this.planTargetNumberingReferences(
        sourcePackage.numbering,
        remaining,
      );
      if (planned === null)
        return {
          status: "unalignable",
          detail: "a referenced style has undefined numbering",
        };
      styleNumberingReferenceMap = new Map([
        ...planned,
        ...numberingReferenceMap,
      ]);
      if (
        this.stageTargetNumbering(
          sourcePackage.numbering,
          remaining,
          planned,
        ) === "conflict"
      ) {
        return {
          status: "unalignable",
          detail: "a referenced style's numbering could not be imported",
        };
      }
    }
    const result = importReferencedStyleDefinitions({
      numberingReferenceMap: styleNumberingReferenceMap,
      sourceStyles: sourcePackage.styles,
      destinationStyles,
      sourceTheme: sourcePackage.theme,
      destinationTheme: destination.theme,
      referencedStyleIds: [...referencedStyleIds],
      reservedStyleIds: [...reservedStyleIds],
      materializeDefaultParagraphStyle,
    });
    if (result.status === "imported")
      this.stagePackageResourcesOnSession({ styles: result.styles });
    return result;
  }

  private stageTargetNumbering(
    target: NumberingDefinitions | null | undefined,
    references: readonly { numId: number; level: number }[],
    remappedNumIds: ReadonlyMap<number, number>,
  ): StageTargetNumberingResult {
    if (references.length === 0) return "unchanged";
    if (!target) return "conflict";
    const current = this.baseDocument.package.numbering ?? {
      abstractNums: [],
      nums: [],
    };
    const nums = new Map(current.nums.map((entry) => [entry.numId, entry]));
    const abstracts = new Map(
      current.abstractNums.map((entry) => [entry.abstractNumId, entry]),
    );
    const targetNums = new Map(
      target.nums.map((entry) => [entry.numId, entry]),
    );
    const targetAbstracts = new Map(
      target.abstractNums.map((entry) => [entry.abstractNumId, entry]),
    );
    const levelsByNumId = referencedNumberingLevelsByNumId({ references });
    const importedAbstractIds = new Map<number, number>();
    const existingAbstractIds = [
      ...abstracts.keys(),
      ...targetAbstracts.keys(),
    ];
    const abstractIds = createNumberingIdAllocator(
      "abstract",
      existingAbstractIds,
    );
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
        if (
          collision &&
          canonicalJson(collision) !== canonicalJson(targetAbstract)
        ) {
          abstractNumId = abstractIds.next();
        }
        importedAbstractIds.set(targetAbstract.abstractNumId, abstractNumId);
        abstracts.set(abstractNumId, { ...targetAbstract, abstractNumId });
      }
      nums.set(stagedNumId, {
        ...targetNum,
        numId: stagedNumId,
        abstractNumId,
      });
    }
    if (nums.size === current.nums.length) return "unchanged";
    const numbering = {
      ...current,
      abstractNums: [...abstracts.values()],
      nums: [...nums.values()],
    };
    this.stagePackageResourcesOnSession({ numbering });
    return "staged";
  }

  private planTargetNumberingReferences(
    target: NumberingDefinitions | null | undefined,
    references: readonly { numId: number; level: number }[],
  ): ReadonlyMap<number, number> | null {
    if (references.length === 0) return new Map();
    if (!target) return null;
    const current = this.baseDocument.package.numbering ?? {
      abstractNums: [],
      nums: [],
    };
    const nums = new Map(current.nums.map((entry) => [entry.numId, entry]));
    const targetNums = new Map(
      target.nums.map((entry) => [entry.numId, entry]),
    );
    const targetAbstracts = new Map(
      target.abstractNums.map((entry) => [entry.abstractNumId, entry]),
    );
    const levelsByNumId = referencedNumberingLevelsByNumId({ references });
    const existingNumIds = [
      ...numberingNumIdsTaken(current),
      ...numberingNumIdsTaken(target),
    ];
    const numberingIds = createNumberingIdAllocator("num", existingNumIds);
    const remapped = new Map<number, number>();
    for (const numId of levelsByNumId.keys()) {
      const targetNum = targetNums.get(numId);
      const targetAbstract = targetNum
        ? targetAbstracts.get(targetNum.abstractNumId)
        : undefined;
      if (!targetNum || !targetAbstract) return null;
      const existingNum = nums.get(numId);
      if (!existingNum) continue;
      if (
        !sameReferencedNumberingLevels({
          current,
          target,
          numId,
          levelsByNumId,
        })
      ) {
        remapped.set(numId, numberingIds.next());
      }
    }
    return remapped;
  }

  /**
   * A canonical room's revision 0, which its authority is seeded with: the
   * document an editor's canonical session reads from `buffer`, and the
   * DOCX's snapshot identity.
   */
  static async openRoomBase(buffer: ArrayBuffer): Promise<CanonicalRoomBase> {
    const base = await openCanonicalRoomBase(buffer);
    if (base.isErr()) throw base.error;
    return base.value;
  }

  /** Parse a `.docx` buffer into a reviewer. */
  // canonical-gap: publicOps.headlessSession
  static fromBuffer(
    buffer: ArrayBuffer,
    options: FolioDocxReviewerOptions = {},
  ): Promise<FolioDocxReviewer> {
    return FolioDocxReviewer.fromCanonicalBuffer(buffer, options);
  }

  /**
   * Load as a canonical editor does: the same input normalization (paragraph
   * and content control ids minted in the package), then the session over the
   * parsed document. A document the session cannot hold throws its refusal.
   *
   * The body is the session's projection. Headers, footers and notes are
   * still read by the legacy story converters, from the session's document;
   * they move to the session's own story projections with the secondary story
   * edits.
   */
  private static async fromCanonicalBuffer(
    buffer: ArrayBuffer,
    options: FolioDocxReviewerOptions,
  ): Promise<FolioDocxReviewer> {
    const opened = await openCanonicalDocxSession(buffer, options.password);
    if (opened.isErr()) throw opened.error;
    const { document: parsed, mintedParaIds, session } = opened.value;
    return new FolioDocxReviewer({
      // The model at load; the readers follow the session from here on.
      baseDocument: readModel(session.document),
      originalBuffer: parsed.originalBuffer ?? buffer,
      // No paragraph id or bidi pass: the projection is what the editor shows.
      state: canonicalReviewerState(session),
      author: options.author ?? "AI",
      engine: { type: "canonical", session },
      positionalParaIds: new Set([
        ...mintedParaIds,
        ...session.sourceMintedParaIds,
      ]),
    });
  }

  /**
   * The session and the editor's public operation compiler over it: the body
   * compiles against the reviewer's main state, another story against its
   * projection, and every commit publishes through {@link publishOnSession}.
   */
  private canonicalEngine(session: CanonicalSession): FolioReviewerEngine {
    return {
      type: "canonical",
      session,
      operations: new CanonicalPublicOperations({
        session,
        undoLimit: FOLIO_DOCUMENT_OPERATION_UNDO_LIMIT,
        getState: (story) => {
          if (story === OP_STORIES.MAIN) return this.state;
          const projected = session.projectStory(story);
          return projected.isErr()
            ? null
            : canonicalStoryState(
                session,
                projected.value.doc,
                projected.value,
              );
        },
        publish: (commit) => this.publishOnSession(commit),
      }),
    };
  }

  /**
   * A live, independent copy of the reviewer: the same document, stories,
   * comments and pending suggestions, which a DOCX save would drop. Edits to
   * either leave the other as it is. The copy's undo history starts empty.
   */
  fork(): FolioDocxReviewer {
    // A second session over the current document, as `toResolvedDocument`
    // resolves on.
    const document = readModel(this.engine.session.document);
    const created = createCanonicalSession(document, document.package.styles);
    if (created.isErr()) throw created.error;
    const copy = new FolioDocxReviewer({
      baseDocument: document,
      originalBuffer: this.originalBuffer,
      state: canonicalReviewerState(created.value),
      author: this.author,
      engine: { type: "canonical", session: created.value },
      positionalParaIds: this.positionalParaIds,
    });
    copy.pendingSuggestions = this.pendingSuggestions.clone();
    return copy;
  }

  private canonicalComparisonAccess(): FolioDocxComparisonAccess {
    const refuse = (action: string): never => {
      throw canonicalReviewerRefusal(action);
    };
    return Object.freeze({
      // The session writes a story-final table's deletion in standard
      // markup that resolves either way, so it needs no receiver.
      stageTerminalTableReviewCarrier: () => false,
      stageTargetStyles: (options) => this.stageTargetStyles(options),
      createComparisonHeaderFooter: (source, story) =>
        this.createComparisonHeaderFooter(source, story),
      matchInlineAtoms: ({
        story,
        targetReviewer,
        targetStory,
        targetSnapshot,
        noteReferenceLabels,
        ...options
      }) => {
        const result = matchInlineAtomsOnModel({
          ...options,
          document: this.engine.session.document,
          story: canonicalOpStory(story),
          target: {
            document: targetReviewer.engine.session.document,
            story: canonicalOpStory(targetStory),
            snapshot: targetSnapshot,
          },
          noteReferenceLabels:
            noteReferenceLabels ??
            refuse("match inline atoms without the comparison's note labels"),
          author: this.author,
        });
        if (result.status !== "matched") return result;
        // The whole batch is one session commit; a refusal throws.
        if (result.operations.length > 0)
          this.commitOpsOnSession(result.operations);
        return {
          status: "matched",
          nextRevisionId: result.nextRevisionId,
          changedTargetBlockIds: result.changedTargetBlockIds,
          rangeCount: result.rangeCount,
          documentChanged: result.operations.length > 0,
        };
      },
      matchInlineProvenance: ({
        story,
        targetReviewer: _targetReviewer,
        targetStory: _targetStory,
        revisionStamp,
        ...options
      }) => {
        const { session } = this.engine;
        const projected = session.projectStory(canonicalOpStory(story));
        if (projected.isErr()) throw projected.error;
        const projection = projected.value;
        const plan = planInlineProvenance({
          ...options,
          state: canonicalStoryState(session, projection.doc),
        });
        if (plan.status !== "planned") return plan;
        const written = inlineProvenanceOperations({
          document: session.document,
          changes: plan.changes,
          addressAt: (position) => {
            const address = projection.addressAt(position);
            return address.isOk() ? address.value : null;
          },
          revisionStamp,
          author: this.author,
        });
        if (written.status !== "matched") return { status: "unalignable" };
        // The whole batch is one session commit; a refusal throws.
        if (written.operations.length > 0)
          this.commitOpsOnSession(written.operations);
        return {
          status: "matched",
          nextRevisionId: written.nextRevisionId,
          changedTargetBlockIds: written.changedTargetBlockIds,
          rangeCount: plan.changes.length,
          documentChanged: written.operations.length > 0,
        };
      },
      projectStories: (mode) => this.projectComparisonStoriesOnSession(mode),
      snapshotReviewedStory: (options) =>
        this.snapshotReviewedStoryInternal(options),
      numberingDefinitions: () => this.baseDocument.package.numbering,
      styleDefinitions: () => this.baseDocument.package.styles,
      planTargetNumberingReferences: (target, references) =>
        this.planTargetNumberingReferences(target, references),
      stageTargetNumbering: (target, references, remappedNumIds) =>
        this.stageTargetNumbering(target, references, remappedNumIds),
      finalSectionProperties: () => this.currentFinalSectionProperties(),
      stageFinalSectionProperties: (options) =>
        this.stageFinalSectionPropertiesOnSession(options),
      stageMappedSectionBoundaries: (options) =>
        this.stageMappedSectionBoundariesOnSession(options),
    });
  }

  private stageFinalSectionPropertiesOnSession(options: {
    target: SectionProperties;
    previous: SectionProperties;
    revision: FolioRevisionStamp;
  }): boolean {
    const staged = stagedFinalSectionProperties({
      ...options,
      author: this.author,
    });
    const op = staged && this.finalSectionOperation(() => staged);
    return op ? this.commitOpsOnSession([op]) : false;
  }

  /** Resolve on an independent fork, then publish only the complete semantic projection. */
  private projectComparisonStoriesOnSession(
    mode: FolioDocxComparisonProjectionMode,
  ): FolioDocxComparisonProjection {
    const copy = this.fork();
    const handles = copy.listStoryHandlesInternal();
    const ids = new Set<number>();
    let present = false;
    if (mode === "with-revision-census") {
      for (const change of copy.currentFinalSectionProperties()
        ?.propertyChanges ?? []) {
        ids.add(change.info.id);
        present = true;
      }
    }
    const operations: DocumentOp[] = [];
    const session = copy.engine.session;
    const unsubscribe = session.onCommitted(({ ops }) =>
      operations.push(...ops),
    );
    let stories: FolioDocxComparisonStoryProjection[];
    try {
      // Staged suggestions project to final as the rest does: each becomes
      // an ordinary change, in the same projection commit.
      copy.promoteSuggestionsOnSession(session);
      // A main-story resolution can retire a note; read its pending state
      // before that.
      for (const handle of handles) {
        const state = copy.getEditableStoryState(handle);
        if (!state) continue;
        if (getSuggestions(state).length > 0)
          throw canonicalReviewerRefusal(
            "project a story holding a suggestion the session will not accept",
          );
        if (mode === "with-revision-census") {
          const stats = getTrackedChangeStatsFromDoc(state.doc);
          for (const id of stats.ids) ids.add(id);
          present ||= stats.present;
        }
      }
      stories = handles.map((handle) => {
        const exists = copy.resolveStoryOnSession(session, handle, "final");
        return { handle, snapshot: exists ? copy.snapshotStory(handle) : null };
      });
    } finally {
      unsubscribe();
    }
    if (operations.length > 0) this.commitOpsOnSession(operations);
    return { stories, revisions: { ids: [...ids], present } };
  }

  private stagePackageResourcesOnSession(resources: {
    styles?: StyleDefinitions | undefined;
    numbering?: NumberingDefinitions | undefined;
  }): void {
    const before = this.baseDocument;
    this.commitOpsOnSession([
      packageResourcesOpOf({
        before,
        after: { ...before, package: { ...before.package, ...resources } },
      }),
    ]);
  }

  /**
   * Apply a canonical room's sequenced batches in revision order, as each
   * peer applies them. A reviewer opened on the room's base DOCX and given
   * its journal holds the room's document, which `toBuffer` saves.
   */
  applyRoomJournal(journal: readonly SequencedBatch[]): void {
    for (const batch of journal) {
      const expectedRevision = this.roomRevision + 1;
      if (batch.revision !== expectedRevision)
        throw new FolioRoomJournalError({
          message: `The room journal skips from revision ${this.roomRevision} to ${batch.revision}.`,
          reason: "revisionGap",
          expectedRevision,
          receivedRevision: batch.revision,
        });
      this.commitOpsOnSession(batch.ops);
      this.roomRevision = batch.revision;
    }
  }

  private commitOpsOnSession(ops: readonly DocumentOp[]): boolean {
    const prepared = this.engine.session.prepareOperations(this.state, ops);
    if (prepared.isErr()) {
      if (prepared.error.reason === "noChange") return false;
      throw prepared.error;
    }
    return this.publishOnSession(prepared.value);
  }

  private stageMappedSectionBoundariesOnSession(
    options: Parameters<
      FolioDocxComparisonAccess["stageMappedSectionBoundaries"]
    >[0],
  ): ReturnType<FolioDocxComparisonAccess["stageMappedSectionBoundaries"]> {
    const { session } = this.engine;
    const result = stageSectionBoundaryOperations({
      document: readModel(session.document),
      accepted: () => this.resolvedSessionCopy(session, "final").state.doc,
      target: options.target,
      originalRevisionIdSeed: options.originalRevisionIdSeed,
      revisionStamp: options.revision,
      author: this.author,
      maxRanges: options.maxRanges,
      mapTargetProperties: options.mapTargetProperties,
    });
    if (result.status !== "matched") return result;
    return {
      status: "matched",
      rangeCount: result.operations.length,
      nextRevisionId: result.nextRevisionId,
      documentChanged:
        result.operations.length > 0 &&
        this.commitOpsOnSession(result.operations),
    };
  }

  /** The model the readers read: the session's document after every commit. */
  private get baseDocument(): Document {
    return readModel(this.engine.session.document);
  }

  /**
   * Publish a commit the session prepared on the reviewer's main state. The
   * state is then rebuilt from the session. A projection the state refuses
   * throws the session's refusal; nothing falls back to the legacy engine.
   */
  private publishOnSession(commit: CanonicalCommit): boolean {
    const { session } = this.engine;
    // A header, footer or note commit applies to that story's projection;
    // the body follows from the session below.
    const published = publishCanonicalProjection({
      state:
        commit.selection.anchor.story === OP_STORIES.MAIN
          ? this.state
          : canonicalStoryState(session, commit.transaction.before),
      commit,
      session,
    });
    if (published.isErr()) throw published.error;
    this.state = canonicalReviewerState(session);
    return true;
  }

  /**
   * Run a review resolution the session prepared. `false` when it changes
   * nothing or a content control's lock holds every change it names, as on
   * the legacy engine; any other refusal throws.
   */
  private resolveOnSession(
    prepared: Result<CanonicalCommit, CanonicalSessionError>,
  ):
    | {
        resolved: true;
        lockHeld: readonly number[];
        absorbedBlocks: readonly FolioAbsorbedBlock[];
      }
    | { resolved: false; reason: "noChange" | "lockHeld" } {
    if (prepared.isErr()) {
      const { error } = prepared;
      if (error.reason === "noChange")
        return { resolved: false, reason: "noChange" };
      if (error.ground.rule === CANONICAL_RULE.contentControlLock)
        return { resolved: false, reason: "lockHeld" };
      throw error;
    }
    const absorptions = prepared.value.touched.absorbed ?? [];
    const sourceSnapshotForStory = (story: FolioEditableDocumentStoryHandle) =>
      this.pendingSuggestionSource(story);
    // A live proposal on an absorbed paragraph now lives in the survivor, and
    // its record follows it there. Compare the session's saved snapshots:
    // rejecting a suggestion in a different projection can keep another id.
    const absorption =
      absorptions.length === 0
        ? null
        : this.pendingSuggestions.prepareAbsorption({
            survivors: new Map(
              absorptions.map(({ absorbed, survivor }) => [absorbed, survivor]),
            ),
            activeSuggestionIds: this.activeSuggestionIds(),
            sourceSnapshotForStory,
          });
    this.publishOnSession(prepared.value);
    absorption?.({
      sourceSnapshotForStory,
      activeSuggestionIds: this.activeSuggestionIds(),
    });
    // One whose paragraph the resolution removed whole stands beside the
    // paragraphs its suggestion added.
    this.pendingSuggestions.placeRemovedAnchors({
      activeSuggestionIds: this.activeSuggestionIds(),
      stateForStory: (story) => this.getEditableStoryState(story),
      savedSource: sourceSnapshotForStory,
    });
    return {
      resolved: true,
      lockHeld: prepared.value.lockHeld ?? [],
      absorbedBlocks: absorptions.map(({ absorbed, survivor }) => ({
        blockId: absorbed,
        survivorBlockId: survivor,
      })),
    };
  }

  /** A semantic operation for the final section, shared by staging and review resolution. */
  private finalSectionOperation(
    resolve: (current: SectionProperties) => SectionProperties | null,
  ): DocumentOp | null {
    const current = this.currentFinalSectionProperties();
    const next = current && resolve(current);
    if (!current || !next) return null;
    return createCanonicalSectionPropertiesOperation(
      this.baseDocument,
      sectionResolutionPatch({ current, next }),
    );
  }

  /** Classify every story before a held fragment keeps its whole proposal pending. */
  private storyResolutions(resolution: "accept" | "reject"): {
    ops: DocumentOp[];
    count: number;
    skipped: FolioReviewResolutionSkip[];
    stagedIds: number[];
    heldSuggestionIds: Set<string>;
  } {
    const stories = this.listStoryHandlesInternal()
      // Notes resolve before body references can remove their stories.
      .toSorted(
        (left, right) =>
          Number(left.type === "main") - Number(right.type === "main"),
      )
      .flatMap((handle) => {
        const state = this.getEditableStoryState(handle);
        if (!state) return [];
        const story = canonicalOpStory(handle);
        const listed = listedRevisionGroups(state.doc);
        if (listed.length === 0) return [];
        const owners = suggestionOwnersByRevision(state);
        // Bulk acceptance leaves pending proposals staged.
        const changes = listed.flatMap(({ id, groupIds }) => {
          const suggestionId = owners.get(id) ?? null;
          return resolution === "accept" && suggestionId !== null
            ? []
            : [{ id, groupIds, suggestionId }];
        });
        const locks = canonicalRevisionLocks({
          document: this.baseDocument,
          story,
          groups: changes.map(({ groupIds }) => groupIds),
          decision: resolution,
        });
        const held = new Map(
          locks.held.flatMap(({ ids, violation }) =>
            ids.map((id) => [id, violation] as const),
          ),
        );
        const carried = revisionIdsCarriedByResolution(state.doc, resolution);
        return [
          {
            source: { story },
            changes: changes.map(({ id, suggestionId, groupIds }) => {
              const { carries, moves } = carriedByGroup(id, groupIds, carried);
              return {
                id,
                suggestionId,
                groupIds,
                violation: held.get(id),
                carries,
                moves,
              };
            }),
          },
        ];
      });
    const classified = classifyReviewLocks({
      stories,
      grouping: "whole-proposal",
    });
    const ops: DocumentOp[] = [];
    const skipped: FolioReviewResolutionSkip[] = [];
    const stagedIds: number[] = [];
    let count = 0;
    for (const {
      source: { story },
      changes,
    } of classified.stories) {
      const revisionIds = new Set<number>();
      for (const entry of changes) {
        switch (entry.type) {
          case "held":
            skipped.push({
              id: entry.change.id,
              reason: CONTENT_CONTROL_LOCK,
              message: entry.message,
            });
            break;
          case "free":
            count++;
            stagedIds.push(entry.change.id);
            for (const id of entry.change.groupIds) revisionIds.add(id);
            break;
          default: {
            const unhandled: never = entry;
            return panic("Unknown classified review lock", {
              entry: unhandled,
            });
          }
        }
      }
      if (revisionIds.size > 0)
        ops.push({
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story,
          revisionIds: [...revisionIds],
          decision: resolution,
        });
    }
    return {
      ops,
      count,
      skipped,
      stagedIds,
      heldSuggestionIds: classified.heldSuggestionIds,
    };
  }

  /** A story's snapshot as the compiler resolves operations against it. */
  private snapshotOnSession(
    story: FolioEditableDocumentStoryHandle,
  ): FolioAIEditSnapshot {
    if (story.type === "main") return this.snapshotOf(this.state);
    const snapshot = this.snapshotStory(story);
    if (!snapshot)
      throw new FolioDocumentStoryNotFoundError({
        message: `Document story ${JSON.stringify(story)} was not found.`,
        story,
      });
    return snapshot;
  }

  /**
   * Apply a batch through the editor's canonical public operation compiler.
   * A suggested batch is recorded for {@link exportPendingSuggestions} as on
   * the legacy engine; the session holds its revisions and comments.
   */
  private applyDocumentOperationsOnSession({
    operations,
    story,
    batch,
    snapshot,
    author,
    revisionStamp,
    wordDiff,
    tableTemplates,
    replacementBackground,
    undefinedReferences,
    restagesSuggestions,
  }: Omit<ApplyDocumentOperationsInternalOptions, "createUndoEntry"> & {
    operations: CanonicalPublicOperations;
  }): FolioDocumentOperationResult {
    const sourceSnapshot = snapshot ?? this.snapshotOnSession(story);
    const operationAuthor = author ?? this.author;
    const activeBefore = this.activeSuggestionIds();
    const result = operations.apply({
      snapshot: sourceSnapshot,
      batch,
      story: story.type === "main" ? "main" : story,
      author: operationAuthor,
      ...(revisionStamp !== undefined && { revisionStamp }),
      ...(wordDiff !== undefined && { wordDiff }),
      ...(tableTemplates !== undefined && { tableTemplates }),
      ...(replacementBackground !== undefined && { replacementBackground }),
      undefinedReferences,
      ...(restagesSuggestions === true
        ? { restagesSuggestions }
        : { pendingSuggestionIds: activeBefore }),
    });
    if (batch.mode === "suggested") {
      const activeSuggestionIds = this.activeSuggestionIds();
      for (const id of activeBefore) activeSuggestionIds.add(id);
      this.pendingSuggestions.recordApplied({
        snapshot: sourceSnapshot,
        story,
        operations: batch.operations,
        applied: result.applied,
        author: operationAuthor,
        activeSuggestionIds,
      });
    }
    return result;
  }

  /**
   * Snapshot the current document into AI-facing blocks (id, kind, text) plus
   * the anchor map the apply layer resolves operations against.
   */
  snapshot(): FolioAIEditSnapshot {
    return this.snapshotOf(this.state);
  }

  /** Every suggestion id the document's stories still carry. */
  private activeSuggestionIds(): Set<string> {
    const active = new Set(
      getSuggestions(this.state).map(({ suggestionId }) => suggestionId),
    );
    for (const handle of this.listStoryHandlesInternal()) {
      if (handle.type === "main") continue;
      const state = this.getEditableStoryState(handle);
      for (const { suggestionId } of state ? getSuggestions(state) : [])
        active.add(suggestionId);
    }
    return active;
  }

  /** Export staged AI operations for storage by the host, outside the DOCX. */
  exportPendingSuggestions(): FolioPendingSuggestionRecord[] {
    const active = this.activeSuggestionIds();
    return this.pendingSuggestions.exportPendingSuggestions({
      activeSuggestionIds: active,
      snapshotForStory: (story) => this.pendingSuggestionSource(story),
      stateForStory: (story) => this.getEditableStoryState(story),
    });
  }

  /**
   * A story as a save writes it, every pending suggestion rejected: what a
   * suggestion record's fingerprint describes. It is read from a session over
   * the document the save writes, so a reviewer reopened on the saved package
   * reads the same story.
   */
  private pendingSuggestionSource(
    story: FolioEditableDocumentStoryHandle,
  ): FolioAIEditSnapshot | null {
    if (!this.getEditableStoryState(story)) return null;
    const document = this.toDocument();
    const created = createCanonicalSession(document, document.package.styles);
    if (created.isErr()) throw created.error;
    const saved = new FolioDocxReviewer({
      baseDocument: document,
      originalBuffer: this.originalBuffer,
      state: canonicalReviewerState(created.value),
      author: this.author,
      engine: { type: "canonical", session: created.value },
      positionalParaIds: this.positionalParaIds,
    }).getEditableStoryState(story);
    return saved
      ? createPendingSuggestionSourceSnapshot(saved, new Set())
      : null;
  }

  /** Validate saved proposals against this document before restaging them. */
  loadPendingSuggestions(
    records: readonly unknown[],
  ): FolioPendingSuggestionLoadResult[] {
    return this.pendingSuggestions.loadPendingSuggestions({
      records,
      snapshotForStory: (story) => this.snapshotStory(story),
      sourceSnapshotForStory: (story) => this.pendingSuggestionSource(story),
      apply: (record, snapshot) => {
        const restage = {
          undefinedReferences: "refuse",
          restagesSuggestions: true,
          story: record.story,
          batch: {
            version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
            operations: [record.operation],
            mode: "suggested",
          },
          snapshot,
          author: record.author,
        } satisfies Omit<
          ApplyDocumentOperationsInternalOptions,
          "createUndoEntry"
        >;
        const result = this.applyDocumentOperationsOnSession({
          ...restage,
          operations: this.engine.operations,
        });
        return { applied: result.applied, skipped: result.skipped };
      },
      activeSuggestionIds: () => this.activeSuggestionIds(),
    });
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
   *
   * Read for one story (the main one by default): a header, footer or note
   * starts its lists against the package alone, so the lists it started are
   * defined in its own reading, as its labels are, and in no other story's.
   */
  readNumberingDefinitions(): FolioNumberingLevel[];
  /** `null` when the document has no such story. */
  readNumberingDefinitions(
    story: FolioEditableDocumentStoryHandle,
  ): FolioNumberingLevel[] | null;
  readNumberingDefinitions(
    story: FolioEditableDocumentStoryHandle = MAIN_STORY,
  ): FolioNumberingLevel[] | null {
    // A read retains no state for a story nothing edits.
    const doc =
      story.type === "main"
        ? this.state.doc
        : this.readSecondaryStoryDoc(story);
    if (!doc) return null;
    // The session defines every list it started in the package; a story
    // reads those its paragraphs name.
    const numbering = storyOwnNumbering({
      current: this.baseDocument.package.numbering,
      loaded: this.loadedDocument.package.numbering,
      doc,
    });
    return numberingLevelsOf(numbering).toSorted(
      (left, right) => left.numId - right.numId || left.level - right.level,
    );
  }

  /** Return parsed package metadata without exposing the mutable document model. */
  getDocumentProperties(): Readonly<
    NonNullable<Document["package"]["properties"]>
  > | null {
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
  snapshotStory(
    story: FolioEditableDocumentStoryHandle,
  ): FolioAIEditSnapshot | null {
    const state = this.getEditableStoryState(story);
    return state ? this.snapshotOf(state) : null;
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
    if (view !== "current-markup")
      return this.resolvedReader(view).snapshotReviewedStoryInternal({
        story,
        view: "current-markup",
      });
    const sourceState = this.getEditableStoryState(story);
    return sourceState ? this.snapshotOf(sourceState) : null;
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
    if (view !== "current-markup")
      return this.resolvedReader(view).storyTables({
        story,
        view: "current-markup",
      });
    const sourceState = this.getEditableStoryState(story);
    return sourceState ? folioStoryTables(sourceState.doc) : [];
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
    {
      const matched = matchTableGeometryOperations({
        story: canonicalOpStory(story),
        baseTables: folioStoryTables(state.doc),
        targetTables,
        pairings,
        revision: {
          author: this.author,
          date: revisionStamp.date,
          idSeed: revisionStamp.idSeed,
        },
      });
      if (matched.status === "unaddressable")
        throw canonicalReviewerRefusal(
          "change the properties of a table cell without a paragraph id",
        );
      if (matched.operations.length > 0)
        this.commitOpsOnSession(matched.operations);
      return matched.nextRevisionId;
    }
  }

  /**
   * Read one story through an immutable reviewed-view projection. A canonical
   * resolved view refuses when locks leave ordinary revisions pending.
   */
  readReviewedStory(
    options: FolioReadReviewedStoryOptions = {},
  ): FolioReviewedStory | null {
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

  /**
   * Resolve one editable story to its original or final state, its staged
   * suggestions with its other changes. A suggestion that reaches another
   * story resolves only whole, through the suggestion methods. A lock that
   * prevents any target change from resolving refuses before mutation.
   */
  resolveReviewedStory({
    story = MAIN_STORY,
    view,
  }: FolioResolveReviewedStoryOptions): boolean {
    return this.resolveStoryOnSession(this.engine.session, story, view);
  }

  /**
   * Apply operations against the current state. Reuses the live-editor applier
   * verbatim via a headless `{ state, dispatch }` seam; the resulting state
   * (including any comments) is retained for {@link toBuffer}.
   */
  applyOperations(
    operations: FolioAIEditOperation[],
    options: FolioApplyOperationsOptions,
  ): FolioAIEditApplyResult {
    const internal = {
      story: MAIN_STORY,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations,
        mode: options.mode ?? "tracked-changes",
      },
      ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
      ...(options.revisionStamp !== undefined && {
        revisionStamp: options.revisionStamp,
      }),
      ...(options.wordDiff !== undefined && { wordDiff: options.wordDiff }),
      undefinedReferences: options.undefinedReferences,
    } satisfies Omit<ApplyDocumentOperationsInternalOptions, "createUndoEntry">;
    const { applied, skipped } = this.applyDocumentOperationsOnSession({
      ...internal,
      operations: this.engine.operations,
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
    batch: FolioDocumentOperationBatchInput,
    options: FolioApplyDocumentOperationsOptions,
  ): FolioDocumentOperationResult {
    const internal = {
      story: MAIN_STORY,
      batch: parseFolioDocumentOperationBatch(batch),
      ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
      ...(options.revisionStamp !== undefined && {
        revisionStamp: options.revisionStamp,
      }),
      ...(options.wordDiff !== undefined && { wordDiff: options.wordDiff }),
      undefinedReferences: options.undefinedReferences,
    } satisfies Omit<ApplyDocumentOperationsInternalOptions, "createUndoEntry">;
    return this.applyDocumentOperationsOnSession({
      ...internal,
      operations: this.engine.operations,
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
    replacementBackground,
    undefinedReferences,
  }: FolioApplyDocumentOperationsToStoryOptions): FolioDocumentOperationResult {
    const internal = {
      story,
      batch: parseFolioDocumentOperationBatch(batch),
      ...(snapshot !== undefined && { snapshot }),
      ...(revisionStamp !== undefined && { revisionStamp }),
      ...(wordDiff !== undefined && { wordDiff }),
      ...(tableTemplates !== undefined && { tableTemplates }),
      ...(replacementBackground !== undefined && { replacementBackground }),
      undefinedReferences,
    } satisfies Omit<ApplyDocumentOperationsInternalOptions, "createUndoEntry">;
    return this.applyDocumentOperationsOnSession({
      ...internal,
      operations: this.engine.operations,
    });
  }

  /**
   * Undo the latest unchanged document-operation batch and its created
   * comments. Only the last {@link FOLIO_DOCUMENT_OPERATION_UNDO_LIMIT}
   * batches are kept; an older handle is `unknownHandle`.
   */
  undoDocumentOperations(
    undoHandle: FolioDocumentOperationUndoHandle,
  ): FolioDocumentOperationUndoResult {
    // The session's journal undoes the batch, comments included.
    return this.engine.operations.undo(undoHandle);
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
    return formatStorySnapshotForLLM(
      this.snapshot(),
      options.annotated === true,
    );
  }

  /**
   * Header / footer and footnote / endnote text as labeled, LLM-ready lines,
   * one per non-empty part: `[header default] …`, `[footer default] …`,
   * `[footnote #N] …`, `[endnote #N] …`. Lines reflect in-memory edits. Empty
   * parts and separator notes are omitted.
   *
   * A header or footer is labelled with the role its section references give
   * it (`default`, `first`, `even`); in a document with several sections, also
   * with the sections that show it: `[header first (sections 1, 3)] …`.
   */
  getNotesAsText(): string {
    const lines: string[] = [];
    const placements = this.headerFooterPlacements();
    for (const handle of this.listStoryHandlesInternal()) {
      if (handle.type === "main") continue;
      const text = this.getStoryText(handle);
      if (text.length === 0) continue;
      switch (handle.type) {
        case "header":
        case "footer": {
          const source = this.getHeaderFooterStory(handle);
          if (!source)
            panic("A listed document story no longer exists", {
              story: handle,
            });
          const roles = placements.byPart.get(
            `${handle.type}:${handle.relationshipId}`,
          );
          const label = roles
            ? describePlacements(roles, placements.sectionCount)
            : source.hdrFtrType;
          lines.push(`[${handle.type} ${label}] ${text}`);
          break;
        }
        case "footnote":
        case "endnote":
          lines.push(`[${handle.type} #${handle.noteId}] ${text}`);
          break;
        default: {
          const unreachable: never = handle;
          return unreachable;
        }
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
    for (const relationshipId of pkg.footers?.keys() ?? []) {
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
    // On the session, a part a resolution retired is gone, as a save drops it.
    const retired = this.engine.session.retiredHeaderFooters();
    return handles.filter((handle) => {
      if (handle.type === "header" || handle.type === "footer") {
        const key = headerFooterStoryKey(handle);
        return !retired.has(key);
      }
      return true;
    });
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
            candidate.type === handle.type &&
            candidate.relationshipId === handle.relationshipId,
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
      ...(this.currentFinalSectionProperties()?.propertyChanges ?? []).map(
        ({ info }) => ({
          id: info.id,
          type: "sectionPropertiesChanged" as const,
          author: info.author,
          date: info.date ?? null,
          text: "",
          blockId: null,
        }),
      ),
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
    const anchors = this.commentAnchors();
    const covered = new Set<number>();
    for (const story of this.listStoryHandlesInternal()) {
      const doc =
        story.type === "main"
          ? this.state.doc
          : this.readSecondaryStoryDoc(story);
      if (!doc) continue;
      for (const id of coveredCommentIdsInProseDoc(doc)) covered.add(id);
    }
    const definitions = this.keptCommentThreads(
      this.currentCommentDefinitions(),
      this.state,
    ).filter(
      ({ id, parentId, anchorKind }) =>
        anchorKind !== "range" || parentId !== undefined || covered.has(id),
    );
    if (definitions.length === 0) {
      return [];
    }

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
        story: anchor?.story ?? null,
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
   * Delete comments by numeric id or exact author, with every dependent reply.
   * Returns the deleted ids in ascending order.
   */
  deleteComments(
    selector:
      | { ids: readonly number[]; author?: never }
      | { author: string; ids?: never },
  ): { deleted: number[] } {
    const comments = this.currentCommentDefinitions();
    const requested =
      selector.ids !== undefined
        ? new Set(selector.ids)
        : new Set(
            comments
              .filter(({ author }) => author === selector.author)
              .map(({ id }) => id),
          );
    const removed = new Set<number>();
    const roots = new Map<number, "thread" | "reply">();
    const knownIds = new Set(comments.map(({ id }) => id));
    for (const { id, parentId } of comments) {
      if (!requested.has(id) || removed.has(id)) continue;
      for (const descendant of commentThreadIds(comments, id)) {
        removed.add(descendant);
        roots.delete(descendant);
      }
      roots.set(
        id,
        parentId !== undefined && knownIds.has(parentId) ? "reply" : "thread",
      );
    }
    if (removed.size === 0) return { deleted: [] };
    const prepared = this.engine.session.prepareOperations(
      this.state,
      [...roots].map(([id, scope]) => ({
        type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
        id,
        scope,
      })),
    );
    if (prepared.isErr()) throw prepared.error;
    this.publishOnSession(prepared.value);
    return { deleted: [...removed].sort((left, right) => left - right) };
  }

  private currentCommentDefinitions(): Comment[] {
    return [...(this.baseDocument.package.document.comments ?? [])];
  }

  /**
   * Add a reply to a comment thread. Pass a {@link FolioReviewComment} from
   * {@link getComments} or its id; the reply threads under that comment's root
   * (Word threads are flat). On {@link toBuffer} the reply is written as a real
   * Word reply — linked via `commentsExtended.xml` and given its own
   * `commentRange` markers + reference anchored on the parent's range. The
   * live document carries the same anchor. Returns
   * the created reply, or `null` when the target comment is absent.
   */
  replyTo(
    target: FolioReviewComment | number,
    input: FolioReviewReplyInput,
  ): FolioReviewCommentReply | null {
    const parentId = typeof target === "number" ? target : target.id;
    return this.replyOnSession(this.engine.session, parentId, input);
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
  resolveComment(
    commentId: string,
    options: { resolved?: boolean } = {},
  ): boolean {
    const resolved = options.resolved ?? true;
    return this.resolveCommentOnSession(
      this.engine.session,
      commentId,
      resolved,
    );
  }

  /** Run a semantic comment command as one session commit, as the editor does. */
  private commitCommentOnSession(
    session: CanonicalSession,
    command: CanonicalCommentCommand,
  ): void {
    const compiled = compileCanonicalComments({
      document: readModel(session.document),
      command,
      idLease: session.idLease,
    });
    if (compiled.isErr()) throw compiled.error;
    if (compiled.value.ops.length === 0) return;
    const prepared = session.prepareOperations(this.state, compiled.value.ops);
    if (prepared.isErr()) throw prepared.error;
    this.publishOnSession(prepared.value);
  }

  /**
   * A reply on the session: a comment anchored as a reply to `parentId`,
   * which the session threads under the parent's root and anchors on its
   * range, as the editor's reply does.
   */
  private replyOnSession(
    session: CanonicalSession,
    parentId: number,
    input: FolioReviewReplyInput,
  ): FolioReviewCommentReply | null {
    const document = readModel(session.document);
    if (
      !(document.package.document.comments ?? []).some(
        ({ id }) => id === parentId,
      )
    )
      return null;
    const id = freshCommentId(document, session.idLease);
    if (id.isErr()) throw canonicalReviewerRefusal("mint a comment id");
    const author = input.author ?? this.author;
    const date = input.date ?? new Date().toISOString();
    this.commitCommentOnSession(session, {
      type: "create",
      anchor: { kind: "reply", parentId },
      comment: {
        id: id.value,
        author,
        done: false,
        date,
        ...(input.initials !== undefined && { initials: input.initials }),
        content: canonicalCommentBody(document, input.text, session.idLease),
      },
    });
    return { id: id.value, author, date, text: input.text };
  }

  /** Resolve or reopen one comment through the session's comment record. */
  private resolveCommentOnSession(
    session: CanonicalSession,
    commentId: string,
    resolved: boolean,
  ): boolean {
    const target = (
      readModel(session.document).package.document.comments ?? []
    ).find(({ id }) => String(id) === commentId);
    if (!target) return false;
    this.commitCommentOnSession(session, {
      type: "resolve",
      id: target.id,
      status: resolved ? "resolved" : "open",
    });
    return true;
  }

  /**
   * Where each header and footer part is shown, as the live section
   * references state it: its role (`w:type`) and the 1-based section, in
   * section order. A part's own `hdrFtrType` holds one role, and one part can
   * serve several sections, in different roles.
   */
  private headerFooterPlacements(): {
    byPart: Map<string, HeaderFooterPlacement[]>;
    sectionCount: number;
  } {
    const records: SectionProperties[] = [];
    this.state.doc.descendants((node) => {
      if (node.type.name !== "paragraph") return true;
      const record = sectionPropertiesOf(node);
      if (record && !records.includes(record)) records.push(record);
      return false;
    });
    const final = this.currentFinalSectionProperties();
    if (final) records.push(final);
    const byPart = new Map<string, HeaderFooterPlacement[]>();
    for (const [index, record] of records.entries()) {
      const place = (
        part: "header" | "footer",
        type: HeaderFooterType,
        rId: string,
      ): void => {
        const key = `${part}:${rId}`;
        byPart.set(key, [
          ...(byPart.get(key) ?? []),
          { type, section: index + 1 },
        ]);
      };
      for (const { type, rId } of record.headerReferences ?? [])
        place("header", type, rId);
      for (const { type, rId } of record.footerReferences ?? [])
        place("footer", type, rId);
    }
    return { byPart, sectionCount: records.length };
  }

  private currentFinalSectionProperties(): SectionProperties | undefined {
    return this.baseDocument.package.document.finalSectionProperties;
  }

  /**
   * Accept an existing tracked change, keeping its text and dropping the
   * redline. Pass a {@link FolioReviewChange} from {@link getChanges} or its
   * revision id. Reuses the editor's own accept command headlessly, so the
   * resolved state persists on {@link toBuffer}. Returns `false` when the
   * revision is no longer present (already resolved, or never existed), or
   * when a content control's lock holds it: accepting would change a
   * content-locked control's content or remove a control locked against
   * removal. A held change stays pending as it is.
   */
  acceptChange(target: FolioReviewChange | number): boolean {
    return this.resolveChangeOnSession(target, "accept");
  }

  /**
   * Reject an existing tracked change: an insertion's text is removed, a
   * deletion's text is restored. See {@link acceptChange} for targeting.
   */
  rejectChange(target: FolioReviewChange | number): boolean {
    return this.resolveChangeOnSession(target, "reject");
  }

  /** Convert one pending suggestion into an ordinary tracked change. */
  acceptSuggestion(
    suggestionId: string,
    options?: { author?: string },
  ): boolean {
    return this.resolveSuggestionOnSession(
      this.engine.session,
      suggestionId,
      "accept",
      options?.author ?? this.author,
    );
  }

  /** Remove one pending suggestion without applying it. */
  rejectSuggestion(suggestionId: string): boolean {
    return this.resolveSuggestionOnSession(
      this.engine.session,
      suggestionId,
      "reject",
      this.author,
    );
  }

  /**
   * Accept every ordinary tracked change in the package. Pending suggestions
   * remain staged until {@link acceptSuggestion} is called. Returns the number
   * of ordinary changes present before the sweep.
   *
   * Every story, not just the body: a revision in a header or a footnote is
   * one a reviewer meant to resolve, and leaving it behind means an accepted
   * document still carries a redline nobody can see from the body.
   *
   * A change a content control's lock holds (see {@link acceptChange}) stays
   * pending as it is and is listed in `skipped`; `resolved` counts the
   * ordinary changes resolved.
   */
  acceptAll(): FolioReviewResolution {
    return this.resolveAllOnSession(this.engine.session, "accept");
  }

  /** Reject every tracked change in the package. See {@link acceptAll}. */
  rejectAll(): FolioReviewResolution {
    return this.resolveAllOnSession(this.engine.session, "reject");
  }

  /**
   * One change of the body, resolved by the session's review intent: every
   * revision of the change (one a comment or a hyperlink splits carries
   * several ids) resolves together, as the editor resolves it.
   */
  private resolveChangeOnSession(
    target: FolioReviewChange | number,
    resolution: "accept" | "reject",
  ): boolean {
    const { session } = this.engine;
    const id = revisionIdOf(target);
    // A suggested change is rejected whole; accepting it first adopts it as
    // an ordinary tracked change, then accepts that change.
    const suggestionId = suggestionIdOfRevision(this.state, id);
    if (suggestionId !== null) {
      const adopted = this.resolveSuggestionOnSession(
        session,
        suggestionId,
        resolution,
        this.author,
      );
      if (!adopted || resolution === "reject") return adopted;
    }
    const section = this.finalSectionOperation((current) =>
      resolveSectionChange({ current, decision: resolution, id }),
    );
    const prepared = session.prepareResolve(this.state, {
      revisionIds: [id],
      resolution,
      before: section ? [section] : [],
    });
    // A change the engine will not resolve on its own stays pending, as a
    // locked one does: the caller reads `false`.
    if (prepared.isErr() && prepared.error.reason === "refused") return false;
    return this.resolveOnSession(prepared).resolved;
  }

  /**
   * One pending suggestion through the session's suggestion intent: rejected
   * with the comments it proposes, or adopted as ordinary tracked changes by
   * `author`. All stories belonging to the proposal publish together.
   */
  private resolveSuggestionOnSession(
    session: CanonicalSession,
    suggestionId: string,
    resolution: "accept" | "reject",
    author: string,
  ): boolean {
    if (!this.activeSuggestionIds().has(suggestionId)) return false;
    const prepared = session.prepareSuggestion(this.state, {
      suggestionId,
      resolution,
      author,
      date: new Date().toISOString(),
    });
    // A proposal the engine will not resolve stays staged: `false`.
    if (prepared.isErr() && prepared.error.reason === "refused") return false;
    return this.resolveOnSession(prepared).resolved;
  }

  /**
   * Every reviewable change in the package in one session commit. A change a
   * content control's lock holds stays pending as it is and is listed in
   * `skipped`, as on the legacy engine.
   */
  private resolveAllOnSession(
    session: CanonicalSession,
    resolution: "accept" | "reject",
  ): FolioReviewResolution {
    // The session never restages proposals, so the registry is what exports
    // them: a suggestion in the document without a record would be lost.
    const recorded = this.pendingSuggestions.suggestionIds();
    for (const suggestionId of this.activeSuggestionIds()) {
      if (!recorded.has(suggestionId))
        return panic(
          "A pending suggestion in the document has no record in the reviewer's registry",
          { suggestionId, resolution },
        );
    }
    // Rejections include proposals in the same commit as ordinary changes;
    // acceptance leaves proposals staged.
    const stories = this.storyResolutions(resolution);
    const skipped = [...stories.skipped];
    const sectionIds = (
      this.currentFinalSectionProperties()?.propertyChanges ?? []
    ).map(({ info }) => info.id);
    const section = this.finalSectionOperation((current) =>
      resolveSectionChanges(current, resolution),
    );
    const proposedComments =
      resolution === "reject"
        ? (this.baseDocument.package.document.comments ?? []).filter(
            ({ suggestionId }) =>
              suggestionId !== undefined &&
              !stories.heldSuggestionIds.has(suggestionId),
          )
        : [];
    // Headers, footers and notes resolve in the same commit as the body.
    if (
      stories.count === 0 &&
      proposedComments.length === 0 &&
      section === null
    )
      return { resolved: 0, skipped };
    const outcome = this.resolveOnSession(
      session.prepareResolve(this.state, {
        revisionIds: [],
        resolution,
        before: [
          ...proposedComments.map(({ id }) => ({
            type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
            id,
            scope: "thread" as const,
          })),
          ...stories.ops,
          ...(section ? [section] : []),
        ],
      }),
    );
    if (!outcome.resolved) {
      if (outcome.reason === "noChange") return { resolved: 0, skipped };
      skipped.push(
        ...[...stories.stagedIds, ...sectionIds].map(
          (id) =>
            ({
              id,
              reason: CONTENT_CONTROL_LOCK,
              message: `Resolving change ${String(id)} changes a locked content control.`,
            }) as const,
        ),
      );
      return { resolved: 0, skipped };
    }
    return {
      resolved: stories.count + sectionIds.length,
      skipped,
      ...(outcome.absorbedBlocks.length === 0
        ? {}
        : { absorbedBlocks: outcome.absorbedBlocks }),
    };
  }

  /** Resolve ordinary revisions on a detached in-memory document. */
  toResolvedDocument(decision: "accept" | "reject"): Document {
    return this.resolvedCopy(
      decision === "accept" ? "final" : "original",
    ).toDocument();
  }

  /**
   * A detached reviewer holding the document with every change resolved to
   * `view`, which every resolved-view reader reads. Both engines resolve
   * through docx-core's revision resolution: a second session resolved that
   * way, or a legacy reviewer over the model that resolution returns.
   */
  private resolvedCopy(view: FolioResolvedReviewedView): FolioDocxReviewer {
    return this.resolvedSessionCopy(this.engine.session, view);
  }

  /** The {@link resolvedCopy} a reader takes one story of. */
  private resolvedReader(view: FolioResolvedReviewedView): FolioDocxReviewer {
    return this.resolvedSessionReader(this.engine.session, view);
  }

  /**
   * The {@link resolvedSessionCopy} a reader takes one story of: the stories
   * of one document share it until a commit replaces the document. Held
   * weakly, so a reviewer does not keep the copy's session alive.
   */
  private resolvedSessionReader(
    session: CanonicalSession,
    view: FolioResolvedReviewedView,
  ): FolioDocxReviewer {
    const { document } = session;
    const held = this.resolvedSessionReaders.get(view)?.deref();
    if (held?.document === document) return held.reviewer;
    const reviewer = this.resolvedSessionCopy(session, view);
    this.resolvedSessionReaders.set(view, new WeakRef({ document, reviewer }));
    return reviewer;
  }

  /**
   * A reviewer on a second session over the current document, as a reviewer
   * reopened on it would hold it, with every change resolved to `view`. This
   * reviewer's session and journal are left as they are; what the session
   * cannot resolve yet refuses as `acceptAll` and `rejectAll` do.
   */
  private resolvedSessionCopy(
    session: CanonicalSession,
    view: FolioResolvedReviewedView,
  ): FolioDocxReviewer {
    const document = readModel(session.document);
    const created = createCanonicalSession(document, document.package.styles);
    if (created.isErr()) throw created.error;
    const copy = new FolioDocxReviewer({
      baseDocument: document,
      originalBuffer: this.originalBuffer,
      state: canonicalReviewerState(created.value),
      author: this.author,
      engine: { type: "canonical", session: created.value },
      positionalParaIds: this.positionalParaIds,
    });
    copy.pendingSuggestions = this.pendingSuggestions.clone();
    // Every change resolves to the view, staged suggestions too: accepted,
    // each becomes an ordinary change first; a rejection takes them with the
    // rest.
    if (view === "final") copy.promoteSuggestionsOnSession(created.value);
    const resolution = view === "final" ? copy.acceptAll() : copy.rejectAll();
    if (resolution.skipped.length > 0)
      throw new CanonicalSessionError({
        reason: "refused",
        ground: { rule: CANONICAL_RULE.contentControlLock },
        message: `A locked content control prevents the ${view} view.`,
      });
    if (pendingSuggestionIdsIn(copy.baseDocument).size > 0)
      throw canonicalReviewerRefusal(
        `read the ${view} view while suggestions remain pending`,
      );
    return copy;
  }

  /**
   * Turn every staged suggestion into an ordinary tracked change, so that
   * accepting every change accepts the suggestions too.
   */
  private promoteSuggestionsOnSession(session: CanonicalSession): void {
    for (const suggestionId of this.activeSuggestionIds())
      this.resolveSuggestionOnSession(
        session,
        suggestionId,
        "accept",
        this.author,
      );
  }

  /**
   * A story holding staged suggestions resolved to `view`: its suggestions
   * resolve with its other changes, on a fork, and the fork's operations
   * land here as one commit, so a refusal leaves this reviewer as it was. A
   * suggestion that reaches another story resolves only whole, so resolving
   * one of its stories refuses.
   */
  private resolveStoryWithSuggestions({
    story,
    view,
    staged,
  }: {
    story: FolioEditableDocumentStoryHandle;
    view: FolioResolvedReviewedView;
    staged: ReadonlySet<string>;
  }): boolean {
    const key = editableStoryKey(story);
    for (const handle of this.listStoryHandlesInternal()) {
      if (editableStoryKey(handle) === key) continue;
      const other = this.getEditableStoryState(handle);
      if (
        other &&
        getSuggestions(other).some(({ suggestionId }) =>
          staged.has(suggestionId),
        )
      )
        throw canonicalReviewerRefusal(
          "resolve one story of a suggestion that reaches another",
        );
    }
    const copy = this.fork();
    const { session } = copy.engine;
    const operations: DocumentOp[] = [];
    const unsubscribe = session.onCommitted(({ ops }) =>
      operations.push(...ops),
    );
    let resolved: boolean;
    try {
      for (const suggestionId of staged)
        copy.resolveSuggestionOnSession(
          session,
          suggestionId,
          view === "final" ? "accept" : "reject",
          this.author,
        );
      const after = copy.getEditableStoryState(story);
      if (after && getSuggestions(after).length > 0)
        throw canonicalReviewerRefusal(
          "resolve a story holding a suggestion the session will not resolve",
        );
      resolved = copy.resolveStoryOnSession(session, story, view);
    } finally {
      unsubscribe();
    }
    if (operations.length > 0) this.commitOpsOnSession(operations);
    return resolved;
  }

  /**
   * Resolve one story's changes to `view` in one session commit: the body's
   * through the review intent, another story's as its resolution operation.
   */
  private resolveStoryOnSession(
    session: CanonicalSession,
    story: FolioEditableDocumentStoryHandle,
    view: FolioResolvedReviewedView,
  ): boolean {
    if (!isFolioResolvedReviewedView(view))
      throw new UnsupportedFolioReviewedViewError({
        message:
          "Only original and final views can replace editable story state.",
        receivedView: view,
      });
    // As on the legacy engine: `true` for every story that exists, resolved
    // to the view, `false` for one that does not.
    const state = this.getEditableStoryState(story);
    if (!state) return false;
    const staged = new Set(
      getSuggestions(state).map(({ suggestionId }) => suggestionId),
    );
    if (staged.size > 0)
      return this.resolveStoryWithSuggestions({ story, view, staged });
    const resolution = view === "final" ? "accept" : "reject";
    const section =
      story.type === "main"
        ? this.finalSectionOperation((current) =>
            resolveSectionChanges(current, resolution),
          )
        : null;
    const groups = listedRevisionGroups(state.doc).map(
      ({ groupIds }) => groupIds,
    );
    const ids = groups.flat();
    if (ids.length === 0 && section === null) return true;
    const locks = canonicalRevisionLocks({
      document: this.baseDocument,
      story: canonicalOpStory(story),
      groups,
      decision: resolution,
    });
    const firstHeld = locks.held.at(0);
    if (firstHeld !== undefined)
      throw new CanonicalSessionError({
        reason: "refused",
        ground: { rule: CANONICAL_RULE.contentControlLock },
        message: lockedControlMessage(firstHeld.violation),
      });
    const outcome = this.resolveOnSession(
      story.type === "main"
        ? session.prepareResolve(this.state, {
            revisionIds: ids,
            resolution,
            before: section ? [section] : [],
          })
        : session.prepareResolve(this.state, {
            revisionIds: [],
            resolution,
            before: [
              {
                type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
                story: canonicalOpStory(story),
                revisionIds: ids,
                decision: resolution,
              },
            ],
          }),
    );
    if (
      (outcome.resolved && outcome.lockHeld.length > 0) ||
      (!outcome.resolved && outcome.reason === "lockHeld")
    )
      throw new CanonicalSessionError({
        reason: "refused",
        ground: { rule: CANONICAL_RULE.contentControlLock },
        message: `A locked content control prevents the ${view} story view.`,
      });
    const remaining = this.getEditableStoryState(story);
    return (
      remaining === null || getTrackedChangesFromDoc(remaining.doc).length === 0
    );
  }

  /** The current document model with edits merged back in. */
  toDocument(): Document {
    // The session's document as it holds it, with pending suggestions
    // rejected as the legacy model has them. The save's own pruning (notes no
    // reference reaches, lists nothing uses) is the save's, not the model's.
    const model = withoutPendingSuggestions(
      this.engine.session.captureSaveSnapshot().document,
    );
    if (model.isErr()) throw model.error;
    // The package relationships the save writes for the parts the model
    // holds, a numbering part a new list started included, as the legacy
    // model states them. The caller owns the copy: the session's operations
    // remember the packages they made, so a change to one would reach them.
    return cloneDocumentWithParagraphPropertySources(
      reconcileDocumentPartRelationships(model.value),
    );
  }

  /**
   * Serialise the reviewed document to a new `.docx` buffer. Tries a selective
   * patch first (only changed paragraphs are rewritten; every other byte of the
   * original package is preserved), falling back to a full repack for
   * structural edits — the same two-tier path the editor's save uses.
   */
  async toBuffer(): Promise<ArrayBuffer> {
    const result = await this.save();
    if (result.type === "repackRefused") {
      return panic("A save that allows a full repack refused to repack", {
        result,
      });
    }
    return result.buffer;
  }

  /**
   * Serialise like {@link toBuffer} and report which path wrote the package.
   * With `repack: "refuse"`, a save the selective patch cannot express
   * returns the reason instead of rewriting every part.
   */
  save({
    repack: repackPolicy = "allow",
  }: FolioReviewerSaveOptions = {}): Promise<FolioReviewerSaveResult> {
    return this.saveOnSession(this.engine.session, repackPolicy);
  }

  /**
   * The editor's canonical save of the session's committed snapshot: the
   * selective patch of the loaded package when it applies, a repack
   * otherwise. Pending suggestions are written as rejected, as the editor
   * writes them.
   */
  private async saveOnSession(
    session: CanonicalSession,
    repackPolicy: "allow" | "refuse",
  ): Promise<FolioReviewerSaveResult> {
    const snapshot = session.captureSaveSnapshot();
    const reason =
      snapshot.structure === "changed"
        ? "structuralChange"
        : "selectiveDeclined";
    // A refused repack is decided before one is built.
    let saved: Awaited<ReturnType<typeof serializeCanonicalSave>>;
    try {
      saved = await serializeCanonicalSave({ snapshot, repack: repackPolicy });
    } catch (error) {
      if (CanonicalSaveRepackRefusedError.is(error))
        return { type: "repackRefused", reason };
      throw error;
    }
    if (saved.path === "selective")
      return { type: "selective", buffer: saved.buffer };
    return { type: "full-repack", reason, buffer: saved.buffer };
  }

  private getEditableStoryState(
    story: FolioEditableDocumentStoryHandle,
  ): EditorState | null {
    if (story.type === "main") {
      return this.state;
    }
    if (!this.secondaryStorySource(story)) return null;
    // Nothing is held for a canonical story: its state is the session's
    // projection at every read.
    const projected = this.engine.session.projectStory(canonicalOpStory(story));
    if (projected.isErr()) throw projected.error;
    return canonicalStoryState(
      this.engine.session,
      projected.value.doc,
      projected.value,
    );
  }

  /**
   * A secondary story as its editable state holds it: the loaded state's
   * document once an operation has loaded it, otherwise the conversion loading
   * performs — with the same deterministic block ids — without retaining a
   * state for a story nothing edits.
   */
  private readSecondaryStoryDoc(
    story: FolioSecondaryStoryHandle,
  ): PMNode | null {
    return this.secondaryStoryDoc(story);
  }

  private secondaryStorySource(story: FolioSecondaryStoryHandle) {
    return story.type === "header" || story.type === "footer"
      ? this.getHeaderFooterStory(story)
      : this.getNoteStory(story);
  }

  /**
   * On either engine a secondary story is converted from the reviewer's
   * model. A canonical reviewer's model is the session's document, so its
   * headers, footers and notes are read from the session but not through the
   * session's story projections yet.
   */
  private secondaryStoryDoc(story: FolioSecondaryStoryHandle): PMNode | null {
    // The story as the session projects it, the document its edits act on.
    if (!this.secondaryStorySource(story)) return null;
    const projected = this.engine.session.projectStory(canonicalOpStory(story));
    if (projected.isErr()) throw projected.error;
    return projected.value.doc;
  }

  private getHeaderFooterStory(
    story: FolioHeaderFooterStoryHandle,
  ): HeaderFooter | undefined {
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
    if (!target || !this.canImportHeaderFooterContent(source, story))
      return null;
    let importedWatermark: Pick<
      HeaderFooter,
      "watermark" | "rawWatermarkXml" | "watermarkBlockIndex"
    > = {};
    let importedMedia: MediaFile | undefined;
    if (
      target.watermark !== undefined ||
      target.rawWatermarkXml !== undefined
    ) {
      const watermark = target.watermark;
      if (
        story.type !== "header" ||
        watermark?.kind !== "picture" ||
        watermark.imageTargetExternal ||
        !watermark.imageTarget ||
        !target.rawWatermarkXml
      )
        return null;
      const media = source.baseDocument.package.media?.get(
        watermark.imageTarget,
      );
      if (!media || !/^image\/(?:png|jpeg|gif|tiff|bmp)$/u.test(media.mimeType))
        return null;
      const xml = rebindDrawingImageRelationship({
        xml: target.rawWatermarkXml,
        previousId: watermark.imageRId,
        nextId: watermark.imageRId,
      });
      if (xml === null) return null;
      const extension = media.path.split(".").at(-1)?.toLowerCase();
      if (!extension || !/^[a-z0-9]+$/u.test(extension)) return null;
      const originalZip = await ZipArchive.read(this.originalBuffer);
      const existingPaths = new Set(
        originalZip.paths().map((path) => path.toLowerCase()),
      );
      let suffix = 1;
      let path = `word/media/folio-import-${suffix}.${extension}`;
      while (
        existingPaths.has(path.toLowerCase()) ||
        this.baseDocument.package.media?.has(path)
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
    // The part and its media join the package unreferenced, as on the
    // legacy engine: the section's reference is staged on its own. Its
    // paragraph takes an identity the session mints.
    const document = this.baseDocument;
    const minted = this.engine.session.freshBlockIds([1]);
    if (minted.isErr()) throw minted.error;
    const paraId = minted.value.at(0)?.at(0);
    if (paraId === undefined)
      return panic("The session minted no paragraph id for a header part");
    const entries: (readonly [string, HeaderFooter])[] = [
      ...((story.type === "header"
        ? document.package.headers
        : document.package.footers) ?? new Map<string, HeaderFooter>()),
      [
        relationshipId,
        { ...empty, content: [{ type: "paragraph", paraId, content: [] }] },
      ],
    ];
    const parts: StoryParts =
      story.type === "header" ? { headers: entries } : { footers: entries };
    const ops: DocumentOp[] = [];
    if (importedMedia)
      ops.push(
        packageResourcesOpOf({
          before: document,
          after: {
            ...document,
            package: {
              ...document.package,
              media: new Map([
                ...(document.package.media ?? new Map()),
                [importedMedia.path, importedMedia],
              ]),
            },
          },
        }),
      );
    ops.push({
      type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
      expected: captureStoryParts(document, parts),
      parts,
    });
    this.commitOpsOnSession(ops);
    return handle;
  }

  private canImportHeaderFooterContent(
    source: FolioDocxReviewer,
    story: FolioHeaderFooterStoryHandle,
  ): boolean {
    const state = source.getEditableStoryState(story);
    if (!state) return false;
    if (!tableTemplateCanCrossPackageLosslessly(state.doc, "rebind-external"))
      return false;
    const snapshot = source.snapshotStory(story);
    return (
      snapshot !== null &&
      detachFolioAIEditSnapshotExternalHyperlinks(snapshot) !== null
    );
  }

  private getNoteStory(
    story: FolioNoteStoryHandle,
  ): Footnote | Endnote | undefined {
    const notes = this.baseDocument.package;
    if (story.type === "footnote") {
      return notes.footnotes?.find(
        (note) => note.id === story.noteId && !isSeparatorFootnote(note),
      );
    }
    return notes.endnotes?.find(
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

  private getHeaderFooterStoryText(source: HeaderFooter): string {
    return normalizeFolioAIBlockText(getHeaderFooterText(source));
  }

  private getNoteStoryText(source: Footnote | Endnote): string {
    const { content } = source;
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
      return this.getHeaderFooterStoryText(source);
    }
    const source = this.getNoteStory(story);
    if (!source) {
      return panic("A listed document story no longer exists", { story });
    }
    return this.getNoteStoryText(source);
  }

  /**
   * `comments` without the threads an edit left anchored nowhere: anchored in
   * a story when it was loaded (or created anchored by this reviewer), and
   * anchored in no story now. See `comment-lifecycle.ts`.
   */
  private keptCommentThreads(
    comments: readonly Comment[],
    mainState: EditorState,
  ): Comment[] {
    const anchoredBefore = new Set(this.loadedMainCommentIds);
    const anchoredNow = anchoredCommentIdsInProseDoc(mainState.doc);
    const addAll = (target: Set<number>, ids: Iterable<number>) => {
      for (const id of ids) {
        target.add(id);
      }
    };
    const referenceStates = noteReferenceStates(mainState.doc);
    for (const handle of this.listStoryHandlesInternal()) {
      if (handle.type === "main") {
        continue;
      }
      const source =
        handle.type === "header" || handle.type === "footer"
          ? this.getHeaderFooterStory(handle)
          : this.getNoteStory(handle);
      addAll(anchoredNow, anchoredCommentIdsInBlocks(source?.content ?? []));
    }
    // A note whose reference is pending deletion keeps the comments anchored in
    // it until the deletion is resolved.
    for (const [key, state] of referenceStates) {
      if (state !== "deleted") continue;
      const handle = noteStoryHandleOf(key);
      if (!handle) continue;
      addAll(
        anchoredNow,
        anchoredCommentIdsInBlocks(this.getNoteStory(handle)?.content ?? []),
      );
    }
    return withoutLostCommentThreads(comments, { anchoredBefore, anchoredNow });
  }

  /**
   * Map each anchored comment id to its anchored text, its containing block id
   * and the story holding it. A comment can be anchored in any story, so every
   * story is read, not only the body.
   */
  private commentAnchors(): Map<number, FolioCommentAnchor> {
    const anchors = new Map<number, FolioCommentAnchor>();
    const add = (doc: PMNode, story: FolioEditableDocumentStoryHandle) => {
      for (const anchor of getCommentAnchorsFromDoc(doc)) {
        if (!anchors.has(anchor.commentId)) {
          anchors.set(anchor.commentId, {
            text: anchor.quote,
            blockId: anchor.blockId,
            story,
          });
        }
      }
    };
    for (const story of this.listStoryHandlesInternal()) {
      const doc =
        story.type === "main"
          ? this.state.doc
          : this.readSecondaryStoryDoc(story);
      if (!doc) {
        continue;
      }
      add(doc, story);
    }
    return anchors;
  }
}

/** @internal Comparison-only access to atomic and fresh reviewer projections. */
export const getFolioDocxComparisonAccess = (
  reviewer: FolioDocxReviewer,
): FolioDocxComparisonAccess => {
  const access = comparisonAccessByReviewer.get(reviewer);
  if (access) return access;
  return panic(
    "Comparison access was requested for an uninitialized DOCX reviewer",
  );
};

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
  /** The engine the review runs on; see {@link FolioDocxReviewerOptions.session}. */
  session?: HeadlessReviewerSession;
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
    ...(options.session !== undefined && { session: options.session }),
  });
  const { applied, skipped } = reviewer.applyOperations(operations, {
    undefinedReferences: "refuse",
    ...(options.mode !== undefined && { mode: options.mode }),
    ...(options.snapshot !== undefined && { snapshot: options.snapshot }),
  });
  const reviewed = await reviewer.toBuffer();
  return { buffer: reviewed, applied, skipped };
};
