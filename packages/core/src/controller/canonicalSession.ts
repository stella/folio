import {
  type Err,
  type Ok,
  panic,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";
import { Fragment, type Node as PMNode } from "prosemirror-model";
import {
  AllSelection,
  type EditorState,
  NodeSelection,
  Selection,
  TextSelection,
  type Transaction,
} from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";

import { hasIllegalXmlCharacters } from "@stll/docx-core";
import {
  lockedControlMessage,
  type LockedControlViolation,
} from "@stll/docx-core/model";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
  combineEdits,
  compileEditorIntent,
  createEditorIntentIdAllocator,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type EditorIntent,
  type EditorIntentMode,
  editorParagraphGroups,
  findStoryBody,
  normalizeForOps,
  OP_STORIES,
  type OpStory,
  packageParagraphIds,
  paragraphIdsIn,
  sameStory,
  suggestedCommentIds,
  suggestionRevisionInfos,
  survivingComments,
  type TextPosition,
  type TouchedBlocks,
  validateOpsDocument,
  withParagraphCopies,
} from "@stll/docx-core/ops";

import {
  splitsGraphemeCluster,
  splitsSurrogatePair,
} from "../ai-edits/character-boundaries";
import { getTrackedChangeGroupsForIds } from "../ai-edits/read";
import {
  CommentMarkerCache,
  type ParagraphCommentMarkers,
  paragraphCommentMarkers,
  visitCommentMarkerItems,
} from "../docx/commentAnchorIndex";
import { ListNormalizationCache } from "../docx/listModelNormalization";
import { cloneDocumentWithParagraphPropertySources } from "../docx/paragraphPropertySource";
import { visitInlineContentSlots } from "../docx/paragraphTraversal";
import { isCellMergeContinuation } from "../docx/tableParser";
import { calculateRowSpans } from "../docx/verticalMergeProjection";
import {
  canonicalCellSelectionRanges,
  canonicalContentPosition,
  canonicalSelectionRange,
} from "../prosemirror/canonicalSelectionRange";
import { expectCommentRangeBoundaryAttrs } from "../prosemirror/commentRangeBoundaryAttrs";
import {
  collectPairedBookmarkIds,
  footnoteToProseDoc,
  headerFooterToProseDoc,
  ProjectionBlockCache,
  toProseDoc,
} from "../prosemirror/conversion/toProseDoc";
import { markPackageChange } from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { runFormattingPatchFromMarks } from "../prosemirror/runFormattingFromMarks";
import {
  CANONICAL_GAP,
  CANONICAL_RULE,
  type CanonicalGap,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";
import type { CanonicalSaveSnapshot } from "../types/canonicalSave";
import type {
  BlockContent,
  Comment,
  Document,
  Paragraph,
  StyleDefinitions,
} from "../types/document";
import type { DeepReadonly, ReadonlyDocument } from "../types/documentModel";
import {
  canonicalNativeCells,
  InlineProjectionCache,
  projectCanonicalInline,
} from "./canonicalInlineProjection";
import { normalizeCanonicalListRendering } from "./canonicalListRendering";
import { lockViolationBetween } from "./canonicalLockGuard";
import {
  type BlockParagraph,
  type ParagraphAddress,
  ParagraphIdIndex,
  type ProjectedBlock,
  sameIdOrder,
  storyFactContributions,
  type StoryProjectionFacts,
} from "./canonicalProjectionBlocks";
import {
  changedWindow,
  unsupportedEditReason,
  unsupportedSeedReason,
} from "./canonicalSeedContent";
import { restoreCanonicalSelection } from "./canonicalSelection";
import {
  type CanonicalSourceRestore,
  canonicalSourceRestores,
  normalizeCanonicalSource,
  withUneditedSourceRestored,
} from "./canonicalSourceNormalization";
import { mapTocSelection } from "./canonicalTocSelection";
import {
  ownedSourceCopies,
  preservePropertySources,
  preservePropertySourcesWholeStories,
} from "./propertySourcePreservation";

const CanonicalSessionErrorBase: TaggedErrorClass<"CanonicalSessionError"> =
  TaggedError("CanonicalSessionError");

export class CanonicalSessionError extends CanonicalSessionErrorBase<{
  message: string;
  reason: "refused" | "noChange";
  ground: CanonicalRefusalGround;
}> {}

const refuse = (message: string) =>
  Result.err(
    new CanonicalSessionError({
      ground: { gap: CANONICAL_GAP.dispatch },
      message,
      reason: "refused",
    }),
  );
/** A change a content control's lock forbids: the document's rule, not a gap. */
const lockRefusal = (violation: LockedControlViolation) =>
  Result.err(
    new CanonicalSessionError({
      ground: { rule: CANONICAL_RULE.contentControlLock },
      message: `Refused: ${lockedControlMessage(violation)}`,
      reason: "refused",
    }),
  );
const noChange = (message: string) =>
  Result.err(
    new CanonicalSessionError({
      ground: { gap: CANONICAL_GAP.dispatch },
      message,
      reason: "noChange",
    }),
  );

/**
 * What the session says for an intent the compiler or the engine refused. A
 * change the engine cannot record belongs to the capability that would record
 * it: review provenance for hyperlinks, and the tracked table records for a
 * suggested row, column or table change. Every other refusal is unclassified.
 */
const intentRefusal = (
  intent: EditorIntent,
  mode: CanonicalSessionMode,
  { reason, message }: { reason: string; message: string },
) => {
  if (reason === DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE)
    return noChange(message);
  if (reason !== DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE)
    return refuse(message);
  let gap: CanonicalGap;
  switch (intent.type) {
    case "setHyperlink":
    case "removeHyperlink":
    case "insertHyperlink":
      gap = CANONICAL_GAP.trackedHyperlinkResolution;
      break;
    case "table":
    case "tableRow":
    case "tableBlock":
      if (mode.type !== "suggesting") return refuse(message);
      gap = CANONICAL_GAP.trackedTableStructure;
      break;
    default:
      return refuse(message);
  }
  return Result.err(
    new CanonicalSessionError({ ground: { gap }, message, reason: "refused" }),
  );
};

const COMMENT_OPERATION_TYPES: ReadonlySet<DocumentOp["type"]> = new Set([
  DOCUMENT_OP_TYPES.CREATE_COMMENT,
  DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
  DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
  DOCUMENT_OP_TYPES.DELETE_COMMENT,
  DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE,
]);

const operationChangesPackage = (op: DocumentOp): boolean => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
    case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
      return true;
    default: {
      if ("story" in op) return !sameStory(op.story, OP_STORIES.MAIN);
      if ("at" in op) return !sameStory(op.at.story, OP_STORIES.MAIN);
      if ("from" in op) return !sameStory(op.from.story, OP_STORIES.MAIN);
      const unreachable: never = op;
      return unreachable;
    }
  }
};

const operationChangesStructure = (op: DocumentOp): boolean => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.DELETE_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_BLOCKS:
    case DOCUMENT_OP_TYPES.SPLIT_BLOCK:
    case DOCUMENT_OP_TYPES.JOIN_BLOCKS:
    case DOCUMENT_OP_TYPES.REPLACE_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_TABLE:
    case DOCUMENT_OP_TYPES.DELETE_TABLE:
    case DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS:
    case DOCUMENT_OP_TYPES.INSERT_ROW:
    case DOCUMENT_OP_TYPES.DELETE_ROW:
    case DOCUMENT_OP_TYPES.SET_TABLE_ROWS:
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
    case DOCUMENT_OP_TYPES.DELETE_COLUMN:
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
    case DOCUMENT_OP_TYPES.SPLIT_CELL:
    case DOCUMENT_OP_TYPES.SET_TABLE:
    case DOCUMENT_OP_TYPES.RESOLVE_REVISION:
    // Table, row and cell properties and the grid sit outside every paragraph,
    // so a save that splices changed paragraphs cannot carry them.
    case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
    case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
    case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
    case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
    // A row's or cell's own revision sits outside every paragraph too.
    case DOCUMENT_OP_TYPES.SET_REVISION_INFO:
      return true;
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
    case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
      return true;
    case DOCUMENT_OP_TYPES.INSERT_TEXT:
    case DOCUMENT_OP_TYPES.INSERT_CONTENT:
    case DOCUMENT_OP_TYPES.DELETE_RANGE:
    case DOCUMENT_OP_TYPES.SPLIT_INLINE:
    case DOCUMENT_OP_TYPES.JOIN_INLINE:
    case DOCUMENT_OP_TYPES.SET_RUN_PROPS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS:
    case DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW:
    case DOCUMENT_OP_TYPES.REPLACE_INLINE:
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
      return false;
    default: {
      const unreachable: never = op;
      return unreachable;
    }
  }
};

const operationChangesBodyProjection = (op: DocumentOp): boolean => {
  switch (op.type) {
    case DOCUMENT_OP_TYPES.CREATE_COMMENT:
    case DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT:
    case DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION:
    case DOCUMENT_OP_TYPES.DELETE_COMMENT:
    case DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE:
    case DOCUMENT_OP_TYPES.CREATE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.REMOVE_HEADER_FOOTER:
    case DOCUMENT_OP_TYPES.ADD_NOTE:
    case DOCUMENT_OP_TYPES.REMOVE_NOTE:
    case DOCUMENT_OP_TYPES.SET_SECTION_PROPS:
    case DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS:
    case DOCUMENT_OP_TYPES.SET_PACKAGE_RESOURCES:
    case DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.DELETE_NUMBERING_INSTANCE:
    case DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT:
      return true;
    default: {
      if ("story" in op) return sameStory(op.story, OP_STORIES.MAIN);
      if ("at" in op) return sameStory(op.at.story, OP_STORIES.MAIN);
      if ("from" in op) return sameStory(op.from.story, OP_STORIES.MAIN);
      const unreachable: never = op;
      return unreachable;
    }
  }
};

export type CanonicalSelection = {
  type: "text" | "inlineNode" | "cell" | "all";
  anchor: TextPosition;
  head: TextPosition;
};
export type CanonicalOrigin = "input" | "undo" | "redo";
/**
 * What a commit does to the session's undo history once published: `opened`
 * a new undo step, `joined` the step the previous input opened (typing on
 * within the grouping window), or walked the history back (`undo`) or
 * forward (`redo`). A host that keeps its own undo stack adds one entry per
 * `opened` commit, so both stacks stay the same length.
 */
export type CanonicalHistoryEffect = "opened" | "joined" | "undo" | "redo";
export type CanonicalSessionMode =
  { type: "editing" } | { type: "suggesting"; author: string };
export const CANONICAL_PROJECTION_META = "folioCanonicalProjection";

/**
 * Which source gap an editor position names where invisible markers (comment
 * ranges, bookmarks, permission ranges, proofing marks) share it: a caret
 * takes the gap typed text takes (the seam's own gap); `afterMarkers` takes the
 * seam's last gap and `beforeMarkers` its first.
 */
export type CanonicalAddressEdge = "caret" | "afterMarkers" | "beforeMarkers";

type CanonicalProjectionOptions = {
  doc: PMNode;
  story: OpStory;
  /** The story's blocks the projection was made from. */
  content: readonly BlockContent[];
  /** One per top-level node of `doc`, in order. */
  blocks: readonly ProjectedBlock[];
  index: ParagraphIdIndex;
  facts: StoryProjectionFacts;
  /**
   * Whether the next projection may keep this one's blocks: each top-level
   * node was projected from the story block at its index alone, and no two
   * source paragraphs share an id.
   */
  reusable: boolean;
};

/** Paragraph gaps bind logical units and source zero-width ordinals to PM positions. */
class CanonicalProjection {
  readonly doc: PMNode;
  readonly story: OpStory;
  readonly content: readonly BlockContent[];
  readonly blocks: readonly ProjectedBlock[];
  readonly index: ParagraphIdIndex;
  readonly facts: StoryProjectionFacts;
  readonly reusable: boolean;
  /** Where each top-level node starts. */
  private readonly starts: readonly number[];
  private readonly addresses: Map<string, ParagraphAddress>;

  constructor(
    options: CanonicalProjectionOptions,
    /** What a projection of an equal document already worked out. */
    held?: {
      starts: readonly number[];
      addresses: Map<string, ParagraphAddress>;
    },
  ) {
    this.doc = options.doc;
    this.story = options.story;
    this.content = options.content;
    this.blocks = options.blocks;
    this.index = options.index;
    this.facts = options.facts;
    this.reusable = options.reusable;
    this.addresses = held?.addresses ?? new Map();
    if (held !== undefined) {
      this.starts = held.starts;
      return;
    }
    const starts: number[] = [];
    let position = 0;
    for (let child = 0; child < this.doc.childCount; child += 1) {
      starts.push(position);
      position += this.doc.child(child).nodeSize;
    }
    this.starts = starts;
  }

  /**
   * The same projection held as `doc`, a document equal to this one's. A
   * commit's transaction ends on a document equal to the projection it staged,
   * sharing the blocks it left; holding that document lets the next commit see
   * the state it was given is the projection without comparing every node.
   */
  heldAs(doc: PMNode): CanonicalProjection {
    return new CanonicalProjection(
      {
        doc,
        story: this.story,
        content: this.content,
        blocks: this.blocks,
        index: this.index,
        facts: this.facts,
        reusable: this.reusable,
      },
      { starts: this.starts, addresses: this.addresses },
    );
  }

  private addressOf(block: number, paragraph: number): ParagraphAddress {
    const entry =
      this.blocks[block]?.paragraphs[paragraph] ??
      panic("Missing canonical paragraph");
    const known = this.addresses.get(entry.blockId);
    if (known !== undefined) return known;
    const address: ParagraphAddress = {
      blockId: entry.blockId,
      container: entry.container ?? this.content,
      start:
        (this.starts[block] ?? panic("Missing block start")) + entry.offset,
      text: entry.inline.text,
      boundaries: entry.inline.boundaries,
      seams: entry.inline.seams,
      node: entry.node,
      source: entry.source,
    };
    this.addresses.set(entry.blockId, address);
    return address;
  }

  /** The paragraph whose content holds `position`, its ends included. */
  private paragraphAround(position: number): ParagraphAddress | undefined {
    // The last top-level node starting before the position.
    let low = 0;
    let high = this.starts.length - 1;
    let block = -1;
    while (low <= high) {
      const middle = (low + high) >> 1;
      if ((this.starts[middle] ?? 0) < position) {
        block = middle;
        low = middle + 1;
      } else high = middle - 1;
    }
    const paragraphs = this.blocks[block]?.paragraphs;
    if (paragraphs === undefined) return undefined;
    const relative = position - (this.starts[block] ?? 0);
    low = 0;
    high = paragraphs.length - 1;
    let found = -1;
    while (low <= high) {
      const middle = (low + high) >> 1;
      if ((paragraphs[middle]?.offset ?? 0) <= relative) {
        found = middle;
        low = middle + 1;
      } else high = middle - 1;
    }
    const paragraph = paragraphs[found];
    if (
      paragraph === undefined ||
      relative > paragraph.offset + paragraph.node.content.size
    )
      return undefined;
    return this.addressOf(block, found);
  }

  addressAt(
    position: number,
    edge: CanonicalAddressEdge = "caret",
  ): Result<TextPosition, CanonicalSessionError> {
    if (!Number.isInteger(position))
      return refuse("The input position is not an integer.");
    position = canonicalContentPosition(this.doc, position);
    const paragraph = this.paragraphAround(position);
    if (paragraph === undefined)
      return refuse("The input is outside a plain paragraph.");
    const relative = position - paragraph.start;
    const offset = paragraph.boundaries.findIndex((gaps) =>
      gaps.some((gap) => gap.position === relative),
    );
    if (offset < 0)
      return refuse("The input splits an indivisible inline element.");
    if (splitsSurrogatePair(paragraph.text, offset)) {
      return refuse("The input would split a surrogate pair.");
    }
    const gaps =
      paragraph.boundaries.at(offset) ?? panic("Missing canonical boundary");
    const index = gaps.findIndex((entry) => entry.position === relative);
    const gap = gaps[index] ?? panic("Missing canonical native gap");
    const seam =
      paragraph.seams.at(offset)?.[index] ?? panic("Missing canonical seam");
    let ordinal = gap.zeroWidthBefore;
    if (edge === "afterMarkers") ordinal = seam.last;
    if (edge === "beforeMarkers") ordinal = seam.first;
    return Result.ok({
      story: this.story,
      blockId: paragraph.blockId,
      offset,
      ...(gaps.length > 1 || ordinal > 0 || seam.first !== seam.last
        ? { zeroWidthBefore: ordinal }
        : {}),
    });
  }

  /**
   * The source gaps a selection from `from` to `to` replaces between. Its
   * edges stay inside the markers beside them, so replacing exactly the
   * words a comment covers keeps the comment on the new text. A selection of
   * a paragraph's whole text takes that paragraph's markers with it: an edge
   * at the very start of a paragraph whose text the selection covers to its
   * end reaches before the markers there, and an edge at the very end of a
   * paragraph the selection covers from its start reaches after them, so a
   * comment on nothing but the replaced text is emptied, as removing the
   * text would empty it.
   */
  rangeEdgesAt(
    from: number,
    to: number,
  ): { start: CanonicalAddressEdge; end: CanonicalAddressEdge } {
    const paragraphOf = (position: number) =>
      this.paragraphAround(canonicalContentPosition(this.doc, position));
    const first = paragraphOf(from);
    const last = paragraphOf(to);
    const startsWhole =
      first !== undefined &&
      canonicalContentPosition(this.doc, from) === first.start &&
      canonicalContentPosition(this.doc, to) >=
        first.start + first.node.content.size;
    const endsWhole =
      last !== undefined &&
      canonicalContentPosition(this.doc, to) ===
        last.start + last.node.content.size &&
      canonicalContentPosition(this.doc, from) <= last.start;
    return {
      start: startsWhole ? "beforeMarkers" : "afterMarkers",
      end: endsWhole ? "afterMarkers" : "beforeMarkers",
    };
  }

  inputAddressAt(
    position: number,
    edge: CanonicalAddressEdge = "caret",
  ): Result<TextPosition, CanonicalSessionError> {
    const address = this.addressAt(position, edge);
    if (address.isErr()) return address;
    const paragraph = this.paragraph(address.value.blockId);
    if (paragraph === undefined) panic("Input lost its paragraph.");
    let destination = position;
    paragraph.node.forEach((node, start) => {
      if (
        node.marks.some((mark) => mark.type.name === "deletion") &&
        destination > paragraph.start + start &&
        destination < paragraph.start + start + node.nodeSize
      )
        destination = paragraph.start + start + node.nodeSize;
    });
    return destination === position
      ? address
      : this.addressAt(destination, edge);
  }

  positionAt(address: TextPosition): Result<number, CanonicalSessionError> {
    const paragraph = this.paragraph(address.blockId);
    if (
      !sameStory(address.story, this.story) ||
      paragraph === undefined ||
      !Number.isInteger(address.offset) ||
      address.offset < 0 ||
      address.offset > paragraph.text.length ||
      splitsSurrogatePair(paragraph.text, address.offset)
    ) {
      return refuse("The canonical selection is outside a plain paragraph.");
    }
    const gaps =
      paragraph.boundaries.at(address.offset) ??
      panic("Missing canonical boundary");
    // An ordinal a seam's editor position stands for, not only the seam's own.
    const seams =
      paragraph.seams.at(address.offset) ?? panic("Missing canonical seams");
    const lastOrdinal = (seams.at(-1) ?? panic("Missing canonical seam")).last;
    // No ordinal stated: after every zero-width leaf at the offset.
    const ordinal = address.zeroWidthBefore ?? lastOrdinal;
    const gap =
      gaps.find(({ zeroWidthBefore }) => zeroWidthBefore === ordinal) ??
      gaps.find((_, index) => {
        const seam = seams[index];
        return (
          seam !== undefined && seam.first <= ordinal && ordinal <= seam.last
        );
      });
    if (!Number.isInteger(ordinal) || ordinal < 0 || ordinal > lastOrdinal)
      return refuse("The canonical selection is outside a paragraph gap.");
    if (gap === undefined)
      return refuse(
        "This source gap is inside a collapsed or erased inline boundary and has no editor position.",
      );
    return Result.ok(paragraph.start + gap.position);
  }

  selectionAt(
    state: EditorState,
  ): Result<CanonicalSelection, CanonicalSessionError> {
    const selection = state.selection;
    let type: CanonicalSelection["type"];
    let anchorPosition = selection.anchor;
    let headPosition = selection.head;
    if (selection instanceof TextSelection) type = "text";
    else if (selection instanceof CellSelection) {
      type = "cell";
      anchorPosition =
        Selection.atStart(
          selection.$anchorCell.nodeAfter ?? panic("Missing anchor cell"),
        ).from +
        selection.$anchorCell.pos +
        1;
      headPosition =
        Selection.atEnd(
          selection.$headCell.nodeAfter ?? panic("Missing head cell"),
        ).to +
        selection.$headCell.pos +
        1;
    } else if (
      selection instanceof NodeSelection &&
      selection.node.isInline &&
      selection.node.isAtom &&
      selection.node.type.name !== "noteMarker"
    )
      type = "inlineNode";
    else if (selection instanceof AllSelection) {
      type = "all";
      const range = canonicalSelectionRange(state);
      anchorPosition = range.from;
      headPosition = range.to;
    } else
      return refuse(
        "Canonical input requires text, an inline atom, or the whole document selection.",
      );
    const anchor = this.addressAt(anchorPosition);
    if (anchor.isErr()) return anchor;
    const head = this.addressAt(headPosition);
    if (head.isErr()) return head;
    return Result.ok({ type, anchor: anchor.value, head: head.value });
  }

  /** The first paragraph the projection addresses, in document order. */
  get firstParagraph(): ParagraphAddress | undefined {
    const block = this.blocks.findIndex(
      ({ paragraphs }) => paragraphs.length > 0,
    );
    return block < 0 ? undefined : this.addressOf(block, 0);
  }

  paragraph(blockId: string): ParagraphAddress | undefined {
    const slot = this.index.get(blockId);
    if (slot === undefined || slot.paragraph < 0) return undefined;
    return this.addressOf(slot.block, slot.paragraph);
  }

  /** The paragraph whose node starts at `before`. */
  paragraphStartingAt(before: number): ParagraphAddress | undefined {
    const paragraph = this.paragraphAround(before + 1);
    return paragraph?.start === before + 1 ? paragraph : undefined;
  }

  nextParagraphInContainer(blockId: string): ParagraphAddress | undefined {
    const slot = this.index.get(blockId);
    if (slot === undefined || slot.paragraph < 0) return undefined;
    const paragraph = this.addressOf(slot.block, slot.paragraph);
    let block = slot.block;
    let next = slot.paragraph + 1;
    while (
      block < this.blocks.length &&
      next >= (this.blocks[block]?.paragraphs.length ?? 0)
    ) {
      block += 1;
      next = 0;
    }
    if (block >= this.blocks.length) return undefined;
    const following = this.addressOf(block, next);
    return paragraph.container === following.container ? following : undefined;
  }
}

type ChangedStoriesOptions = { document: Document; previous: Document };
/** Immutable operations retain content identities for every untouched story. */
const changedStories = ({
  document,
  previous,
}: ChangedStoriesOptions): OpStory[] => {
  const stories: OpStory[] = [];
  if (document.package.document.content !== previous.package.document.content)
    stories.push(OP_STORIES.MAIN);
  for (const kind of ["header", "footer"] as const) {
    const parts =
      kind === "header" ? document.package.headers : document.package.footers;
    const prior =
      kind === "header" ? previous.package.headers : previous.package.footers;
    if (parts === prior) continue;
    for (const [rId, body] of parts ?? []) {
      if (body.content !== prior?.get(rId)?.content)
        stories.push({ kind, rId });
    }
  }
  for (const kind of ["footnote", "endnote"] as const) {
    const notes =
      kind === "footnote"
        ? document.package.footnotes
        : document.package.endnotes;
    const prior =
      kind === "footnote"
        ? previous.package.footnotes
        : previous.package.endnotes;
    if (notes === prior) continue;
    const priorById = new Map(prior?.map((note) => [note.id, note]));
    for (const note of notes ?? []) {
      const original = priorById.get(note.id);
      if (
        note.content !== original?.content ||
        note.noteType !== original?.noteType
      )
        stories.push({ kind, id: note.id });
    }
  }
  return stories;
};

/** A session's projection pieces, reused by every projection of its stories. */
type StoryProjectionCaches = {
  /** The main story's converted blocks. */
  blocks: ProjectionBlockCache;
  /** Each paragraph's gaps, by its source and editor node. */
  inline: InlineProjectionCache;
  /** Each story's normalized list rendering, by block and the counters before it. */
  lists: ListNormalizationCache;
  /** Each paragraph's comment markers, by its source. */
  comments: CommentMarkerCache;
};

export const newStoryProjectionCaches = (): StoryProjectionCaches => ({
  blocks: new ProjectionBlockCache(),
  inline: new InlineProjectionCache(),
  lists: new ListNormalizationCache(),
  comments: new CommentMarkerCache(),
});

type ProjectStoryOptions = {
  document: Document;
  styles: DeepReadonly<StyleDefinitions> | null | undefined;
  story?: OpStory;
  /** What earlier projections converted, kept by its session. */
  caches?: StoryProjectionCaches;
  /** A projection of the story whose top-level blocks this one may keep. */
  previous?: CanonicalProjection;
};

/** Whether a paragraph anchors a text box, which converts to a block of its own. */
const holdsTextBox = (paragraph: Paragraph): boolean => {
  let found = false;
  visitInlineContentSlots(paragraph, ({ item }) => {
    if (
      item.type === "run" &&
      item.content.some(
        (content) =>
          content.type === "shape" && content.shape.textBody !== undefined,
      )
    )
      found = true;
  });
  return found;
};

const structureRefusal = () =>
  new CanonicalSessionError({
    ground: { gap: CANONICAL_GAP.storyContentProjection },
    reason: "refused",
    message: "The projection changed paragraph structure.",
  });

/** A story paragraph's records: where it is held and the top-level block holding it. */
type LocatedSource = {
  source: Paragraph;
  container: readonly BlockContent[];
  top: number;
};

/**
 * Every paragraph a projection addresses or skips under `blocks`, with
 * whether it lies in a merged cell's skipped continuation, in document order.
 */
const visitSourceParagraphs = (
  blocks: readonly BlockContent[],
  visit: (
    source: Paragraph,
    container: readonly BlockContent[],
    continuation: boolean,
  ) => void,
  insideSkippedContinuation = false,
): void => {
  for (const block of blocks)
    visitBlockSources(block, blocks, visit, insideSkippedContinuation);
};

const visitBlockSources = (
  block: BlockContent,
  container: readonly BlockContent[],
  visit: Parameters<typeof visitSourceParagraphs>[1],
  insideSkippedContinuation: boolean,
): void => {
  if (block.type === "paragraph") {
    visit(block, container, insideSkippedContinuation);
    return;
  }
  if (block.type === "blockSdt") {
    visitSourceParagraphs(block.content, visit, insideSkippedContinuation);
    return;
  }
  if (block.type !== "table") return;
  const rowSpans = calculateRowSpans(block);
  for (const [rowIndex, row] of block.rows.entries()) {
    let column = row.formatting?.gridBefore ?? 0;
    for (const cell of row.cells) {
      const skippedContinuation =
        isCellMergeContinuation(cell) &&
        (rowSpans.get(`${rowIndex}-${column}`)?.skip ?? false);
      visitSourceParagraphs(
        cell.content,
        visit,
        insideSkippedContinuation || skippedContinuation,
      );
      column += cell.formatting?.gridSpan ?? 1;
    }
  }
};

/** The source paragraphs of a run of story blocks. */
type StorySources = {
  located: Map<string, LocatedSource>;
  /** Each block's source paragraph ids, from the run's first block. */
  idsByBlock: string[][];
  all: Paragraph[];
  continuation: Set<Paragraph>;
  /** Whether an id repeats in the run or is taken outside it. */
  repeated: boolean;
};

const collectSources = (
  content: readonly BlockContent[],
  start: number,
  end: number,
  takenElsewhere: (paraId: string) => boolean,
): StorySources => {
  const found: StorySources = {
    located: new Map(),
    idsByBlock: [],
    all: [],
    continuation: new Set(),
    repeated: false,
  };
  const at: { top: number; ids: string[] } = { top: start, ids: [] };
  const visit = (
    source: Paragraph,
    container: readonly BlockContent[],
    continuation: boolean,
  ): void => {
    found.all.push(source);
    if (continuation) found.continuation.add(source);
    if (source.paraId === undefined) return;
    if (found.located.has(source.paraId) || takenElsewhere(source.paraId))
      found.repeated = true;
    found.located.set(source.paraId, { source, container, top: at.top });
    at.ids.push(source.paraId);
  };
  for (let top = start; top < end; top += 1) {
    at.top = top;
    at.ids = [];
    found.idsByBlock.push(at.ids);
    visitBlockSources(
      content[top] ?? panic("Missing story block"),
      content,
      visit,
      false,
    );
  }
  return found;
};

/** The comment ranges open as a walk goes, with a copy kept while they stay the same. */
class OpenComments {
  readonly ids: Set<number>;
  private kept: readonly number[];

  constructor(open: readonly number[]) {
    this.ids = new Set(open);
    this.kept = open;
  }

  /** The open ranges now, in their set's order. */
  get current(): readonly number[] {
    let same = this.kept.length === this.ids.size;
    let at = 0;
    for (const id of this.ids) {
      if (!same) break;
      same = this.kept[at] === id;
      at += 1;
    }
    if (!same) this.kept = [...this.ids];
    return this.kept;
  }
}

/** What projecting the paragraphs of one top-level node reads from its story. */
type BlockWalk = {
  facts: StoryProjectionFacts;
  open: OpenComments;
  /** The story's own blocks: a paragraph held there has no container of its own. */
  content: readonly BlockContent[];
  locate: (paraId: string) => LocatedSource | undefined;
  /** Takes the paragraph a standalone text box was converted from; false when there is none. */
  claimTextBoxHost: () => boolean;
  /** Records the source a paragraph is projected from; false when one already was. */
  take: (source: Paragraph) => boolean;
  passComments: (paragraph: Paragraph) => void;
  inline: InlineProjectionCache | undefined;
};

/**
 * The paragraphs of the top-level node at `index`, each start from the
 * node's own, and whether each came from the story block at that index.
 */
const projectBlock = (
  top: PMNode,
  index: number,
  walk: BlockWalk,
): Result<
  { paragraphs: BlockParagraph[]; local: boolean },
  CanonicalSessionError
> => {
  const paragraphs: BlockParagraph[] = [];
  let local = true;
  let failure: CanonicalSessionError | undefined;
  const commentContext = {
    rangedIds: walk.facts.rangedIds,
    pointIds: walk.facts.pointIds,
    openIds: walk.open.ids,
  };
  const visit = (
    node: PMNode,
    offset: number,
    parent: PMNode | null,
  ): boolean => {
    if (failure !== undefined) return false;
    // A text box's paragraphs belong to its shape, not to this story. The
    // session holds no address in them, so input there refuses and the box
    // keeps its source as it was.
    if (node.type.name === "textBox") {
      if (
        node.attrs["_docxPlacement"] === "standalone" &&
        node.attrs["_docxHostParagraph"] != null &&
        !walk.claimTextBoxHost()
      )
        failure = structureRefusal();
      return false;
    }
    // A range opened between blocks covers the paragraphs after it.
    if (node.type.name === "blockCommentRangeBoundary") {
      const { marker } = expectCommentRangeBoundaryAttrs(node);
      if (marker.type === "commentRangeStart") walk.open.ids.add(marker.id);
      else walk.open.ids.delete(marker.id);
      return false;
    }
    if (node.type.name !== "paragraph") return true;
    // A row's w:gridBefore and w:gridAfter render as filler cells whose empty
    // paragraph has no source; there is nothing in them to edit.
    if (parent?.attrs["_omittedGridSlot"] != null) return false;
    const paraId = node.attrs["paraId"];
    const located =
      typeof paraId === "string" ? walk.locate(paraId) : undefined;
    if (typeof paraId !== "string" || located === undefined) {
      failure = structureRefusal();
      return false;
    }
    const { source, container } = located;
    if (!walk.take(source)) {
      failure = structureRefusal();
      return false;
    }
    if (located.top !== index) local = false;
    const inlineArgs = {
      source,
      paragraph: node,
      pairedBookmarkIds: walk.facts.pairedBookmarkIds,
      commentContext,
    };
    const inline =
      walk.inline === undefined
        ? projectCanonicalInline(inlineArgs)
        : walk.inline.project(inlineArgs);
    if (inline.isErr()) {
      failure = new CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.dispatch },
        reason: "refused",
        message: inline.error.message,
      });
      return false;
    }
    paragraphs.push({
      blockId: paraId,
      container: container === walk.content ? undefined : container,
      offset: offset + 1,
      inline: inline.value,
      node,
      source,
    });
    walk.passComments(source);
    return false;
  };
  if (visit(top, 0, null) && top.content.size > 0)
    top.descendants((node, offset, parent) => visit(node, offset + 1, parent));
  if (failure !== undefined) return Result.err(failure);
  return Result.ok({ paragraphs, local });
};

type StoryWalkOptions = {
  doc: PMNode;
  content: readonly BlockContent[];
  story: OpStory;
  markersOf: ParagraphCommentMarkers;
  inline: InlineProjectionCache | undefined;
};

/** Opens and closes the ranges a paragraph's markers start and end. */
const commentPasser =
  (markersOf: ParagraphCommentMarkers, open: OpenComments) =>
  (paragraph: Paragraph): void => {
    for (const item of markersOf(paragraph)) {
      if (item.type === "commentRangeStart") open.ids.add(item.id);
      if (item.type === "commentRangeEnd") open.ids.delete(item.id);
    }
  };

/** Every paragraph's address, walking every top-level node. */
const projectEveryBlock = ({
  doc,
  content,
  story,
  markersOf,
  inline,
}: StoryWalkOptions): Result<CanonicalProjection, CanonicalSessionError> => {
  const pairedBookmarkIds = collectPairedBookmarkIds(content);
  const rangedIds = new Set<number>();
  const referencedIds = new Set<number>();
  visitCommentMarkerItems(content, markersOf, (item) => {
    if (item.type === "commentRangeStart") rangedIds.add(item.id);
    if (item.type === "commentReference") referencedIds.add(item.id);
  });
  const facts: StoryProjectionFacts = {
    pairedBookmarkIds,
    rangedIds,
    pointIds: new Set([...referencedIds].filter((id) => !rangedIds.has(id))),
  };
  const open = new OpenComments([]);
  const passComments = commentPasser(markersOf, open);
  const aligned = doc.childCount === content.length;
  const found = collectSources(content, 0, content.length, () => false);
  let reusable = aligned && !found.repeated;
  const sources = found.located;
  const allSources = new Set(found.all);
  const continuationSources = found.continuation;
  const idsByBlock = found.idsByBlock;
  const projectedSources = new Set<Paragraph>();
  // A paragraph holding only text boxes projects as the boxes alone; the first
  // box of its group stands where the paragraph was.
  const orderedSources = [...allSources];
  const sourceOrder = new Map(
    orderedSources.map((source, index) => [source, index]),
  );
  const textBoxHosts = new Set<Paragraph>();
  let nextSource = 0;
  const walk: BlockWalk = {
    facts,
    open,
    content,
    locate: (paraId) => sources.get(paraId),
    claimTextBoxHost: () => {
      // The host is found by the story's paragraph order, across blocks.
      reusable = false;
      let index = nextSource;
      while (
        index < orderedSources.length &&
        continuationSources.has(orderedSources[index] ?? panic("No source"))
      )
        index += 1;
      const host = orderedSources[index];
      if (host === undefined || !holdsTextBox(host)) return false;
      textBoxHosts.add(host);
      nextSource = index + 1;
      // The host carries no address, but a range it opens or closes still
      // covers the paragraphs after it.
      passComments(host);
      return true;
    },
    take: (source) => {
      if (projectedSources.has(source)) return false;
      projectedSources.add(source);
      nextSource = (sourceOrder.get(source) ?? panic("Unordered source")) + 1;
      return true;
    },
    passComments,
    inline,
  };
  const blocks: ProjectedBlock[] = [];
  for (let index = 0; index < doc.childCount; index += 1) {
    const node = doc.child(index);
    const openBefore = open.current;
    const projected = projectBlock(node, index, walk);
    if (projected.isErr()) return projected;
    if (!projected.value.local) reusable = false;
    blocks.push({
      node,
      source: aligned ? content[index] : undefined,
      paragraphs: projected.value.paragraphs,
      sourceIds: aligned ? (idsByBlock[index] ?? []) : [],
      openBefore,
      openAfter: open.current,
    });
  }
  if (
    projectedSources.size + continuationSources.size + textBoxHosts.size !==
      allSources.size ||
    [...projectedSources].some(
      (source) => continuationSources.has(source) || textBoxHosts.has(source),
    ) ||
    [...textBoxHosts].some((source) => continuationSources.has(source))
  ) {
    return refuse("The projection changed paragraph structure.");
  }
  return Result.ok(
    new CanonicalProjection({
      doc,
      story,
      content,
      blocks,
      index: ParagraphIdIndex.of(blocks),
      facts,
      reusable,
    }),
  );
};

/**
 * The projection keeping every top-level block of `previous` an edit left:
 * the run of blocks from the start and the run to the end whose story block
 * and converted node are both the ones `previous` walked. Only the blocks
 * between are walked, from the comment ranges open where they start. Nothing
 * when the kept blocks would read other story facts than they did, or the
 * walked ones reach outside themselves; the story is then walked in full.
 */
const projectChangedBlocks = (
  previous: CanonicalProjection,
  { doc, content, story, markersOf, inline }: StoryWalkOptions,
): CanonicalProjection | undefined => {
  if (!previous.reusable || !sameStory(previous.story, story)) return undefined;
  const before = previous.blocks;
  const count = content.length;
  if (doc.childCount !== count) return undefined;
  const unchanged = (block: ProjectedBlock | undefined, at: number): boolean =>
    block !== undefined &&
    block.source === content[at] &&
    block.node === doc.child(at);
  let start = 0;
  while (
    start < count &&
    start < before.length &&
    unchanged(before[start], start)
  )
    start += 1;
  let after = 0;
  while (
    after < count - start &&
    after < before.length - start &&
    unchanged(before[before.length - 1 - after], count - 1 - after)
  )
    after += 1;
  const end = before.length - after;
  const newEnd = count - after;
  // The bookmarks paired and the comments ranged across the story stay the
  // same when the walked blocks contribute what the ones they replace did.
  if (
    !sameIdOrder(
      storyFactContributions(previous.content.slice(start, end), markersOf),
      storyFactContributions(content.slice(start, newEnd), markersOf),
    )
  )
    return undefined;
  const removed = new Set<string>();
  for (let index = start; index < end; index += 1)
    for (const id of before[index]?.sourceIds ?? []) removed.add(id);
  const found = collectSources(
    content,
    start,
    newEnd,
    (paraId) =>
      previous.index.get(paraId) !== undefined && !removed.has(paraId),
  );
  if (found.repeated) return undefined;
  const sources = found.located;
  const continuationSources = found.continuation;
  const open = new OpenComments(before[start - 1]?.openAfter ?? []);
  const projectedSources = new Set<Paragraph>();
  const walk: BlockWalk = {
    facts: previous.facts,
    open,
    content,
    locate: (paraId) => sources.get(paraId),
    claimTextBoxHost: () => false,
    take: (source) => {
      if (projectedSources.has(source)) return false;
      projectedSources.add(source);
      return true;
    },
    passComments: commentPasser(markersOf, open),
    inline,
  };
  const window: ProjectedBlock[] = [];
  for (let index = start; index < newEnd; index += 1) {
    const node = doc.child(index);
    const openBefore = open.current;
    const projected = projectBlock(node, index, walk);
    if (projected.isErr() || !projected.value.local) return undefined;
    window.push({
      node,
      source: content[index],
      paragraphs: projected.value.paragraphs,
      sourceIds: found.idsByBlock[index - start] ?? [],
      openBefore,
      openAfter: open.current,
    });
  }
  if (
    projectedSources.size + continuationSources.size !== found.all.length ||
    [...projectedSources].some((source) => continuationSources.has(source))
  )
    return undefined;
  const following = before[end];
  if (
    following !== undefined &&
    !sameIdOrder(open.current, following.openBefore)
  )
    return undefined;
  const blocks = [...before.slice(0, start), ...window, ...before.slice(end)];
  return new CanonicalProjection({
    doc,
    story,
    content,
    blocks,
    index: previous.index.replaced({ start, end, removed, window, blocks }),
    facts: previous.facts,
    reusable: true,
  });
};

const project = ({
  document,
  styles,
  story = OP_STORIES.MAIN,
  caches,
  previous,
}: ProjectStoryOptions): Result<CanonicalProjection, CanonicalSessionError> => {
  const body = findStoryBody(document, story);
  if (!body) return refuse("The canonical story no longer exists.");
  const currentStyles = Object.hasOwn(document.package, "styles")
    ? document.package.styles
    : styles;
  const options = {
    ...(currentStyles == null ? {} : { styles: currentStyles }),
    ...(document.package.theme === undefined
      ? {}
      : { theme: document.package.theme }),
  };
  const converted = Result.try({
    try: () => {
      if (story === OP_STORIES.MAIN)
        return toProseDoc(
          document,
          caches === undefined
            ? options
            : { ...options, blockCache: caches.blocks },
        );
      if (story.kind === "header" || story.kind === "footer")
        return headerFooterToProseDoc(body.content, options);
      return footnoteToProseDoc(body.content, options);
    },
    catch: (cause) =>
      new CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.dispatch },
        reason: "refused",
        message: `Canonical projection failed: ${cause instanceof Error ? cause.message : String(cause)}`,
      }),
  });
  if (converted.isErr()) return converted;
  const walk: StoryWalkOptions = {
    doc: converted.value,
    content: body.content,
    story,
    markersOf: caches?.comments.read ?? paragraphCommentMarkers,
    inline: caches?.inline,
  };
  const kept =
    previous === undefined ? undefined : projectChangedBlocks(previous, walk);
  return kept === undefined ? projectEveryBlock(walk) : Result.ok(kept);
};

const intentStory = (intent: EditorIntent): OpStory => {
  switch (intent.type) {
    case "setHyperlink":
    case "removeHyperlink":
    case "insertHyperlink":
    case "replaceFragment":
    case "moveFragment":
    case "replaceText":
    case "formatRun":
    case "insertAtom":
      return intent.from.story;
    case "generateTOC":
    case "splitParagraph":
    case "formatParagraph":
      return intent.at.story;
    case "setList": {
      const first = intent.items.at(0);
      return first?.at.story ?? panic("A list intent requires a paragraph.");
    }
    case "joinParagraphs":
    case "insertParagraphs":
    case "deleteParagraphs":
      return intent.story;
    case "table":
    case "tableRow":
    case "tableBlock":
      return intent.operation.story;
    default: {
      const unreachable: never = intent;
      return unreachable;
    }
  }
};

/** Input intent, never projection-step adjacency, determines the journal partition. */
export type CanonicalInputSemantic =
  | "typing"
  | "deleteBackward"
  | "deleteForward"
  | "paste"
  | "composition"
  | "replacement"
  | "structure";

const UNDO_GROUP_WINDOW_MS = 500;

const samePosition = (left: TextPosition, right: TextPosition): boolean =>
  sameStory(left.story, right.story) &&
  left.blockId === right.blockId &&
  left.offset === right.offset &&
  (left.zeroWidthBefore ?? 0) === (right.zeroWidthBefore ?? 0);

const sameSelection = (
  left: CanonicalSelection,
  right: CanonicalSelection,
): boolean =>
  left.type === right.type &&
  samePosition(left.anchor, right.anchor) &&
  samePosition(left.head, right.head);

type AppliedJournalEntry = {
  type: "applied";
  ops: readonly DocumentOp[];
  inverse: readonly DocumentOp[];
  preSelection: CanonicalSelection;
  postSelection: CanonicalSelection;
  preDocument: Document;
  postDocument: Document;
  version: number;
  origin: "input";
  semantic: CanonicalInputSemantic;
  grouping: "run" | "isolated" | "compositionCorrection";
  time: number;
  boundary: number;
};

type AppliedJournalGroup = {
  entries: AppliedJournalEntry[];
  undo:
    { type: "entries" } | { type: "replayed"; inverse: readonly DocumentOp[] };
  preSelection: CanonicalSelection;
  postSelection: CanonicalSelection;
};

/** The facts that decide whether an input joins the open undo step. */
type JournalGroupingKey = Pick<
  AppliedJournalEntry,
  | "semantic"
  | "grouping"
  | "time"
  | "boundary"
  | "preSelection"
  | "postSelection"
>;

const continuesGroup = (
  previous: AppliedJournalEntry,
  next: JournalGroupingKey,
): boolean => {
  switch (next.semantic) {
    case "typing":
    case "deleteBackward":
    case "deleteForward":
      return (
        previous.grouping === "run" &&
        next.grouping === "run" &&
        previous.semantic === next.semantic &&
        previous.boundary === next.boundary &&
        next.time >= previous.time &&
        next.time - previous.time <= UNDO_GROUP_WINDOW_MS &&
        sameSelection(previous.postSelection, next.preSelection) &&
        samePosition(next.preSelection.anchor, next.preSelection.head)
      );
    case "composition":
      // A late native final may correct a committed composition. New IME
      // gestures and history/selection changes advance the grouping boundary.
      return (
        next.grouping === "compositionCorrection" &&
        previous.semantic === "composition" &&
        previous.boundary === next.boundary &&
        sameSelection(previous.postSelection, next.preSelection)
      );
    case "paste":
    case "replacement":
    case "structure":
      return false;
    default: {
      const unreachable: never = next.semantic;
      return panic(`Unknown canonical input semantic ${unreachable}`);
    }
  }
};

type UndoneJournalEntry = {
  type: "undone";
  group: AppliedJournalGroup;
  redoOps: readonly DocumentOp[];
};

export type CanonicalCommit = {
  transaction: Transaction;
  document: Document;
  projection: CanonicalProjection;
  bodyProjection: CanonicalProjection;
  selection: CanonicalSelection;
  touched: TouchedBlocks;
  version: number;
  origin: CanonicalOrigin;
  /** What publishing does to the session's undo history. */
  history: CanonicalHistoryEffect;
  /** Revisions a content control lock held back; they stay pending. */
  lockHeld?: readonly number[];
  /** Publish only after every projection/plugin transaction has been accepted. */
  publish: () => Result<void, CanonicalSessionError>;
};

type ProjectionAuthorization = {
  session: CanonicalSession;
  version: number;
  doc: PMNode;
};
const authorizedProjections = new WeakMap<
  Transaction,
  ProjectionAuthorization
>();

/** Metadata is descriptive; transaction identity and session version authorize mutations. */
export const isCanonicalProjectionTransaction = (
  transaction: Transaction,
  session: CanonicalSession,
): boolean => {
  const authorization = authorizedProjections.get(transaction);
  return (
    authorization?.session === session &&
    authorization.version === session.version &&
    transaction.doc.eq(authorization.doc)
  );
};

type PublishCanonicalProjectionOptions = {
  state: EditorState;
  commit: CanonicalCommit;
  session: CanonicalSession;
};

/** The model publishes only after PM and every plugin accept its exact projection. */
export const publishCanonicalProjection = ({
  state,
  commit,
  session,
}: PublishCanonicalProjectionOptions):
  | Err<never, CanonicalSessionError>
  | Ok<
      {
        state: EditorState;
        transactions: readonly Transaction[];
      },
      never
    > => {
  if (!isCanonicalProjectionTransaction(commit.transaction, session))
    return refuse("The staged canonical projection is unauthorized.");
  const applied = Result.try({
    try: () => state.applyTransaction(commit.transaction),
    catch: (cause) =>
      new CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.dispatch },
        reason: "refused",
        message: `Canonical plugin staging failed: ${cause instanceof Error ? cause.message : String(cause)}`,
      }),
  });
  if (applied.isErr()) return applied;
  const staged = applied.value;
  if (
    !staged.transactions.includes(commit.transaction) ||
    !staged.state.doc.eq(commit.projection.doc) ||
    staged.transactions.some(
      (transaction) =>
        transaction.docChanged &&
        !isCanonicalProjectionTransaction(transaction, session),
    )
  )
    return refuse("A plugin refused or changed the canonical projection.");
  const published = commit.publish();
  if (published.isErr()) return published;
  return Result.ok(staged);
};

type StageOptions = {
  state: EditorState;
  ops: readonly DocumentOp[];
  selection: CanonicalSelection;
  origin: CanonicalOrigin;
  history: CanonicalHistoryEffect;
  propertySourceDocument?: Document;
  stagedApplied?: AppliedDocumentOp;
  story?: OpStory;
  previousProjection?: CanonicalProjection;
  onPublish: (inverse: readonly DocumentOp[], version: number) => void;
};

type JournalledOpsOptions = {
  state: EditorState;
  ops: readonly DocumentOp[];
  postSelection: CanonicalSelection;
  stagedApplied?: AppliedDocumentOp;
  semantic?: CanonicalInputSemantic;
};

type CanonicalSessionSeedOptions = {
  document: Document;
  projection: CanonicalProjection;
  styles: DeepReadonly<StyleDefinitions> | null | undefined;
  /** Stories the session reads in its own spelling, given back as read while unedited. */
  sourceRestores?: readonly CanonicalSourceRestore[];
  caches: StoryProjectionCaches;
};

type CanonicalReplaceTextInput = {
  from: number;
  to: number;
  text: string;
  story?: OpStory;
  time?: number;
} & (
  | { semantic: "composition"; compositionPhase?: "correction" }
  | {
      semantic?: Exclude<CanonicalInputSemantic, "composition">;
      compositionPhase?: never;
    }
);

type CachedStoryProjection = {
  content: NonNullable<ReturnType<typeof findStoryBody>>["content"];
  theme: Document["package"]["theme"];
  projection: CanonicalProjection;
};

const storyProjectionKey = (
  story: Exclude<OpStory, typeof OP_STORIES.MAIN>,
): string => {
  switch (story.kind) {
    case "header":
    case "footer":
      return `${story.kind}:${story.rId}`;
    case "footnote":
    case "endnote":
      return `${story.kind}:${story.id}`;
    default: {
      const unreachable: never = story;
      return panic(`Unknown canonical story ${unreachable}`);
    }
  }
};

/** Immutable model authority with a journal staged independently of the PM view. */
class CanonicalSession {
  private currentDocument: Document;
  private currentProjection: CanonicalProjection;
  private readonly sourceBlockIds: ReadonlySet<string>;
  private currentVersion = 0;
  private readonly saveTouched = new Set<string>();
  private saveStructure: CanonicalSaveSnapshot["structure"] = "stable";
  private readonly allocateIntentIds = createEditorIntentIdAllocator();
  private mode: CanonicalSessionMode = { type: "editing" };
  private readonly allocatedBlockIds = new Set<string>();
  private nextBlockId = 1;
  private readonly sourceOwners = new Map<string, Paragraph>();
  private readonly storyProjections = new Map<string, CachedStoryProjection>();
  private currentSelection: CanonicalSelection | null = null;
  private readonly applied: AppliedJournalGroup[] = [];
  private groupingBoundary = 0;
  private lifecycle: { type: "committed" } | { type: "composing" } = {
    type: "committed",
  };
  private readonly undone: UndoneJournalEntry[] = [];
  private readonly styles: DeepReadonly<StyleDefinitions> | null | undefined;
  private readonly sourceRestores: readonly CanonicalSourceRestore[];
  /** What earlier projections converted, reused by every later one. */
  private readonly caches: StoryProjectionCaches;

  constructor({
    document,
    projection,
    styles,
    sourceRestores = [],
    caches,
  }: CanonicalSessionSeedOptions) {
    this.sourceRestores = sourceRestores;
    this.caches = caches;
    this.currentDocument = document;
    this.currentProjection = projection;
    this.sourceBlockIds = new Set(packageParagraphIds(document.package));
    for (const blockId of packageParagraphIds(document.package))
      this.allocatedBlockIds.add(blockId.toUpperCase());
    this.advanceBlockId();
    this.styles = styles == null ? styles : structuredClone(styles);
    this.rememberSources(document.package.document.content);
  }

  /** Render only the committed comment projection while provisional IME blocks document snapshots. */
  getCommittedComments(): Comment[] {
    return structuredClone(survivingComments(this.currentDocument));
  }

  setMode(mode: CanonicalSessionMode): void {
    if (
      this.mode.type !== mode.type ||
      (this.mode.type === "suggesting" &&
        mode.type === "suggesting" &&
        this.mode.author !== mode.author)
    )
      this.breakUndoGroup();
    this.mode =
      mode.type === "editing"
        ? { type: "editing" }
        : { type: "suggesting", author: mode.author };
  }

  hasStyle(styleId: string): boolean {
    return (
      (this.styles ?? this.currentDocument.package.styles)?.styles.some(
        (style) => style.styleId === styleId,
      ) ?? false
    );
  }

  private advanceBlockId(): void {
    while (
      this.allocatedBlockIds.has(
        this.nextBlockId.toString(16).padStart(8, "0").toUpperCase(),
      )
    )
      this.nextBlockId += 1;
  }

  private freshBlockId(): Result<string, CanonicalSessionError> {
    if (this.nextBlockId > 0x7fffffff)
      return refuse("The paragraph identity space is exhausted.");
    return Result.ok(
      this.nextBlockId.toString(16).padStart(8, "0").toUpperCase(),
    );
  }

  /**
   * Paragraph identities for one prepared batch, one group per requested count;
   * none matches a live or retired identity and no two groups overlap. Nothing is
   * reserved until the batch applies, so a batch asks once for all its groups.
   */
  freshBlockIds(
    counts: readonly number[],
  ): Result<string[][], CanonicalSessionError> {
    const total = counts.reduce((sum, count) => sum + count, 0);
    const ids: string[] = [];
    for (let candidate = this.nextBlockId; ids.length < total; candidate += 1) {
      if (candidate > 0x7fffffff)
        return refuse("The paragraph identity space is exhausted.");
      const id = candidate.toString(16).padStart(8, "0").toUpperCase();
      if (!this.allocatedBlockIds.has(id)) ids.push(id);
    }
    const groups: string[][] = [];
    let offset = 0;
    for (const count of counts) {
      groups.push(ids.slice(offset, offset + count));
      offset += count;
    }
    return Result.ok(groups);
  }

  private intentMode(
    document: Document,
    intent: EditorIntent,
  ): EditorIntentMode {
    if (
      this.mode.type === "editing" &&
      !this.intentNeedsIdentityIds(document, intent)
    )
      return { type: "editing", reservedBlockIds: this.allocatedBlockIds };
    const ids = this.allocateIntentIds(document, intent);
    return this.mode.type === "editing"
      ? {
          type: "editing",
          newIds: ids.newIds,
          reservedBlockIds: this.allocatedBlockIds,
        }
      : {
          type: "suggesting",
          reservedBlockIds: this.allocatedBlockIds,
          revision: {
            id: ids.revisionId,
            author: this.mode.author,
            date: new Date().toISOString(),
          },
          newIds: ids.newIds,
        };
  }

  private intentNeedsIdentityIds(
    document: Document,
    intent: EditorIntent,
  ): boolean {
    let blockIds: readonly string[];
    switch (intent.type) {
      case "table":
      case "tableRow":
      case "tableBlock":
      case "insertParagraphs":
      case "deleteParagraphs":
        return true;
      case "generateTOC":
      case "setList":
      case "formatParagraph":
        return false;
      case "setHyperlink":
      case "removeHyperlink":
      case "insertHyperlink":
      case "replaceFragment":
      case "moveFragment":
        return true;
      case "replaceText":
      case "insertAtom":
      case "formatRun":
        if (intent.from.blockId !== intent.to.blockId) return true;
        blockIds = [intent.from.blockId];
        break;
      case "splitParagraph":
        if (intent.to !== undefined && intent.at.blockId !== intent.to.blockId)
          return true;
        blockIds = [intent.at.blockId];
        break;
      case "joinParagraphs":
        blockIds = [intent.blockId, intent.nextBlockId];
        break;
      default: {
        const unreachable: never = intent;
        return unreachable;
      }
    }
    const hasIdentity = (value: unknown): boolean => {
      if (typeof value !== "object" || value === null) return false;
      if (Array.isArray(value)) return value.some(hasIdentity);
      if (
        "id" in value &&
        typeof value.id === "number" &&
        ("author" in value || "sdtType" in value)
      )
        return true;
      return Object.values(value).some(hasIdentity);
    };
    const story = intentStory(intent);
    return blockIds.some((blockId) => {
      const paragraph =
        document === this.currentDocument && story === OP_STORIES.MAIN
          ? this.currentProjection.paragraph(blockId)?.source
          : findStoryBody(document, story)?.content.find(
              (block) => block.type === "paragraph" && block.paraId === blockId,
            );
      return paragraph !== undefined && hasIdentity(paragraph);
    });
  }

  /** Record each paragraph among `blocks` as the owner of its id. */
  private rememberSources(blocks: readonly BlockContent[]): void {
    for (const block of blocks) {
      if (block.type === "paragraph" && block.paraId !== undefined)
        this.sourceOwners.set(block.paraId, block);
    }
  }

  get document(): ReadonlyDocument {
    if (this.isComposing)
      throw new CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.dispatch },
        reason: "refused",
        message: "Composition must finish before taking a snapshot.",
      });
    return this.currentDocument;
  }
  captureSaveSnapshot(): CanonicalSaveSnapshot {
    if (this.isComposing)
      throw new CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.save },
        reason: "refused",
        message:
          "Composition must finish before saving the canonical document.",
      });
    const retainedBlockIdKeys = new Set(
      [
        ...this.sourceBlockIds,
        ...packageParagraphIds(this.currentDocument.package),
      ].map((blockId) => blockId.toUpperCase()),
    );
    return {
      document: withUneditedSourceRestored(
        cloneDocumentWithParagraphPropertySources(this.currentDocument),
        this.sourceRestores,
      ),
      version: this.currentVersion,
      changedBlockIds: [...this.saveTouched].filter((blockId) =>
        retainedBlockIdKeys.has(blockId.toUpperCase()),
      ),
      structure: this.saveStructure,
    };
  }

  get projection(): CanonicalProjection {
    return this.currentProjection;
  }

  /** The caret the last published edit returned, when the editor shows it at `position`. */
  private editCaretAt(
    projection: CanonicalProjection,
    position: number,
  ): TextPosition | undefined {
    const caret = this.currentSelection;
    if (
      caret === null ||
      caret.type !== "text" ||
      caret.anchor.offset !== caret.head.offset ||
      caret.anchor.blockId !== caret.head.blockId ||
      (caret.anchor.zeroWidthBefore ?? null) !==
        (caret.head.zeroWidthBefore ?? null) ||
      !sameStory(caret.head.story, projection.story)
    )
      return undefined;
    const shown = projection.positionAt(caret.head);
    return shown.isOk() && shown.value === position ? caret.head : undefined;
  }
  get selection(): CanonicalSelection | null {
    return this.currentSelection;
  }
  /** Every paragraph identity this session has read or minted, live or retired. */
  get reservedBlockIds(): ReadonlySet<string> {
    return this.allocatedBlockIds;
  }
  get version(): number {
    return this.currentVersion;
  }
  get canUndo(): boolean {
    return this.applied.length > 0;
  }
  get canRedo(): boolean {
    return this.undone.length > 0;
  }

  get isComposing(): boolean {
    return this.lifecycle.type === "composing";
  }

  breakUndoGroup(): void {
    this.groupingBoundary += 1;
  }

  beginComposition(): Result<void, CanonicalSessionError> {
    if (this.isComposing)
      return refuse("A canonical composition is already pending.");
    this.lifecycle = { type: "composing" };
    this.breakUndoGroup();
    return Result.ok(undefined);
  }

  endComposition(): void {
    this.lifecycle = { type: "committed" };
    this.breakUndoGroup();
  }

  private checkState(
    state: EditorState,
    projection = this.projection,
  ): Result<void, CanonicalSessionError> {
    if (this.isComposing)
      return refuse("Composition must finish before committing another edit.");
    if (state.doc !== projection.doc && !state.doc.eq(projection.doc)) {
      return refuse("The input uses a stale canonical projection.");
    }
    return Result.ok(undefined);
  }

  prepareReplace(
    state: EditorState,
    {
      from,
      to,
      text,
      story = OP_STORIES.MAIN,
      semantic = "replacement",
      compositionPhase,
      time = Date.now(),
    }: CanonicalReplaceTextInput,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const projected = this.projectStory(story);
    if (projected.isErr()) return projected;
    const projection = projected.value;
    const checked = this.checkState(state, projection);
    if (checked.isErr()) return checked;
    if (!Number.isFinite(time)) return refuse("The input time is invalid.");
    if (from > to || hasIllegalXmlCharacters(text) || /[\t\r\n]/u.test(text)) {
      return refuse("Canonical input accepts a well-formed text replacement.");
    }
    if (from === to && text.length === 0)
      return noChange("The input makes no text change.");
    const replacesStructuralSelection =
      state.selection instanceof AllSelection &&
      from === state.selection.from &&
      to === state.selection.to;
    if (state.selection instanceof CellSelection) {
      if (from !== state.selection.from || to !== state.selection.to)
        return refuse(
          "A replacement range cannot partially address a cell selection.",
        );
      // Typing over selected cells clears each of them and writes the text
      // into the first, as the editor does.
      const intents: EditorIntent[] = [];
      let first = true;
      for (const range of canonicalCellSelectionRanges(state.selection)) {
        const edges = projection.rangeEdgesAt(range.from, range.to);
        const cellFrom = projection.inputAddressAt(range.from, edges.start);
        if (cellFrom.isErr()) return cellFrom;
        const cellTo = projection.inputAddressAt(range.to, edges.end);
        if (cellTo.isErr()) return cellTo;
        intents.push({
          type: "replaceText",
          from: cellFrom.value,
          to: cellTo.value,
          text: first ? text : "",
        });
        first = false;
      }
      return this.prepareIntents(state, { intents, semantic });
    }
    const range = canonicalSelectionRange(state);
    const inputFrom = replacesStructuralSelection ? range.from : from;
    const inputTo = replacesStructuralSelection ? range.to : to;
    const edges =
      inputFrom === inputTo
        ? { start: "caret" as const, end: "caret" as const }
        : projection.rangeEdgesAt(inputFrom, inputTo);
    // Typing goes on where the last edit left the caret: an editor position
    // can stand for several source gaps between markers, and the edit's own
    // caret names the one right after the text it wrote.
    const editCaret = this.editCaretAt(projection, inputFrom);
    const start =
      inputFrom === inputTo && editCaret !== undefined
        ? Result.ok(editCaret)
        : projection.inputAddressAt(inputFrom, edges.start);
    if (start.isErr()) return start;
    const end =
      inputFrom === inputTo && editCaret !== undefined
        ? Result.ok(editCaret)
        : projection.inputAddressAt(inputTo, edges.end);
    if (end.isErr()) return end;
    const preSelection = projection.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    // Just inside a content control's edge, the editor position, not the
    // source gap, says the text belongs to the control.
    const innerEdge = (position: number, side: "start" | "end") => {
      const resolved = state.doc.resolve(position);
      return (
        resolved.parent.type.name === "sdt" &&
        resolved.parentOffset ===
          (side === "start" ? 0 : resolved.parent.content.size)
      );
    };
    const inside = {
      ...(innerEdge(inputFrom, "start") ? { start: true } : {}),
      ...(innerEdge(inputTo, "end") ? { end: true } : {}),
    };
    // A suggestion over a control's whole text would be recorded as the
    // control's own deletion, which accepting would remove with the new text.
    if (
      this.mode.type === "suggesting" &&
      text.length > 0 &&
      inside.start === true &&
      inside.end === true &&
      state.doc.resolve(inputFrom).parent === state.doc.resolve(inputTo).parent
    )
      return refuse(
        "A suggestion over a content control's whole text is unavailable in this session.",
      );
    const intent = {
      type: "replaceText",
      from: start.value,
      to: end.value,
      text,
      ...(Object.keys(inside).length === 0 ? {} : { inside }),
      ...(state.storedMarks === null
        ? {}
        : {
            runPropsPatch: runFormattingPatchFromMarks(
              state.selection.$from.marks(),
              state.storedMarks,
            ),
          }),
    } as const satisfies EditorIntent;
    const compiled = compileEditorIntent(this.currentDocument, {
      intent,
      mode: this.intentMode(this.currentDocument, intent),
    });
    if (compiled.isErr())
      return compiled.error.reason === DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE
        ? noChange(compiled.error.message)
        : refuse(compiled.error.message);
    const { ops, selection: caret } = compiled.value;
    const postSelection = {
      type: "text",
      anchor: caret,
      head: caret,
    } as const satisfies CanonicalSelection;
    const preDocument = this.currentDocument;
    const isCaret = state.selection.empty;
    const deletion =
      isCaret &&
      text.length === 0 &&
      (semantic === "deleteBackward" || semantic === "deleteForward")
        ? deletionRange(
            state,
            semantic === "deleteBackward" ? "backward" : "forward",
          )
        : null;
    const run =
      isCaret &&
      ((semantic === "typing" &&
        from === to &&
        from === state.selection.head) ||
        (deletion?.isOk() &&
          deletion.value.from === from &&
          deletion.value.to === to));
    const grouping = run ? "run" : "isolated";
    const groupingKey = {
      preSelection: preSelection.value,
      postSelection,
      semantic,
      grouping:
        compositionPhase === "correction" ? "compositionCorrection" : grouping,
      time,
      boundary: this.groupingBoundary,
    } as const satisfies JournalGroupingKey;
    // Decided now: publishing refuses once the version or the grouping
    // boundary moved, so the open step is still the one read here.
    const group = this.applied.at(-1);
    const previous = group?.entries.at(-1);
    const joins =
      group !== undefined &&
      previous !== undefined &&
      continuesGroup(previous, groupingKey);
    return this.stage({
      state,
      story,
      ops,
      selection: postSelection,
      origin: "input",
      history: joins ? "joined" : "opened",
      // The compiler's own application of exactly these ops, when it made one.
      ...(compiled.value.applied === undefined
        ? {}
        : { stagedApplied: compiled.value.applied }),
      onPublish: (inverse, version) => {
        const entry = {
          type: "applied",
          ops,
          inverse,
          ...groupingKey,
          preDocument,
          postDocument: this.currentDocument,
          version,
          origin: "input",
        } as const satisfies AppliedJournalEntry;
        if (group !== undefined && joins) {
          group.entries.push(entry);
          group.postSelection = postSelection;
        } else {
          this.applied.push({
            entries: [entry],
            undo: { type: "entries" },
            preSelection: preSelection.value,
            postSelection,
          });
        }
        this.undone.length = 0;
      },
    });
  }

  /**
   * A structural edit over selected cells: each selected cell is cleared and
   * the edit happens at the start of the first, as the editor's own command
   * does over a cell selection. One commit, so one undo step.
   */
  prepareCellSelectionStructure(
    state: EditorState,
    structure: (at: TextPosition) => EditorIntent,
    {
      resourceOps = [],
      semantic = "structure",
    }: {
      resourceOps?: readonly DocumentOp[];
      semantic?: CanonicalInputSemantic;
    } = {},
  ): Result<CanonicalCommit, CanonicalSessionError> {
    if (!(state.selection instanceof CellSelection))
      return refuse("The structural edit needs a cell selection.");
    const projected = this.projectStory(OP_STORIES.MAIN);
    if (projected.isErr()) return projected;
    const projection = projected.value;
    const intents: EditorIntent[] = [];
    let first: TextPosition | undefined;
    for (const range of canonicalCellSelectionRanges(state.selection)) {
      const edges = projection.rangeEdgesAt(range.from, range.to);
      const cellFrom = projection.inputAddressAt(range.from, edges.start);
      if (cellFrom.isErr()) return cellFrom;
      const cellTo = projection.inputAddressAt(range.to, edges.end);
      if (cellTo.isErr()) return cellTo;
      first ??= cellFrom.value;
      if (
        cellFrom.value.blockId !== cellTo.value.blockId ||
        cellFrom.value.offset !== cellTo.value.offset
      )
        intents.push({
          type: "replaceText",
          from: cellFrom.value,
          to: cellTo.value,
          text: "",
        });
    }
    if (first === undefined) return refuse("The cell selection is empty.");
    intents.push(structure({ ...first, offset: 0, zeroWidthBefore: 0 }));
    return this.prepareIntents(state, { intents, resourceOps, semantic });
  }

  prepareIntent(
    state: EditorState,
    intent: EditorIntent,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    return this.prepareIntents(state, { intents: [intent] });
  }

  prepareIntents(
    state: EditorState,
    {
      intents,
      resourceOps = [],
      semantic = "structure",
    }: {
      intents: readonly EditorIntent[];
      resourceOps?: readonly DocumentOp[];
      semantic?: CanonicalInputSemantic;
    },
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const first = intents.at(0);
    if (first === undefined) return noChange("The intent batch is empty.");
    const story = intentStory(first);
    if (intents.some((intent) => !sameStory(story, intentStory(intent))))
      return refuse("A compound input must address one story.");
    const projected = this.projectStory(story);
    if (projected.isErr()) return projected;
    const checked = this.checkState(state, projected.value);
    if (checked.isErr()) return checked;
    const preSelection = projected.value.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    let document = this.currentDocument;
    let postSelection = preSelection.value;
    const ops: DocumentOp[] = [...resourceOps];
    const edits: AppliedDocumentOp[] = [];
    if (resourceOps.length > 0) {
      const imported = applyDocumentOps(document, resourceOps);
      if (imported.isErr()) return refuse(imported.error.message);
      document = imported.value.document;
      edits.push(imported.value);
    }
    for (const intent of intents) {
      if (
        this.mode.type === "suggesting" &&
        (intent.type === "setHyperlink" ||
          intent.type === "removeHyperlink" ||
          intent.type === "insertHyperlink" ||
          intent.type === "generateTOC")
      ) {
        return Result.err(
          new CanonicalSessionError({
            ground: { gap: CANONICAL_GAP.trackedHyperlinkResolution },
            reason: "refused",
            message:
              "Hyperlink and TOC suggestions require serializable wrapper review provenance.",
          }),
        );
      }
      const compiled = compileEditorIntent(document, {
        intent,
        mode: this.intentMode(document, intent),
        firstBlockId: this.nextBlockId,
      });
      if (compiled.isErr())
        return intentRefusal(intent, this.mode, compiled.error);
      // The compiler's own application of exactly these ops, when it made one.
      const applied =
        compiled.value.applied === undefined
          ? applyDocumentOps(document, compiled.value.ops)
          : Result.ok(compiled.value.applied);
      if (applied.isErr())
        return intentRefusal(intent, this.mode, applied.error);
      document = applied.value.document;
      edits.push(applied.value);
      ops.push(...compiled.value.ops);
      if (intent.type === "generateTOC") {
        const mapping = {
          at: intent.at,
          after: compiled.value.selection,
          ops: compiled.value.ops,
        };
        postSelection = Object.assign({}, postSelection, {
          anchor: mapTocSelection(postSelection.anchor, mapping),
          head: mapTocSelection(postSelection.head, mapping),
        });
      } else if (
        intent.type !== "setHyperlink" &&
        intent.type !== "removeHyperlink" &&
        intent.type !== "formatRun" &&
        intent.type !== "formatParagraph" &&
        intent.type !== "setList"
      ) {
        postSelection = {
          type: "text",
          anchor: compiled.value.selection,
          head: compiled.value.selection,
        };
      }
    }
    const stagedApplied = {
      ...combineEdits(this.currentDocument, edits),
      revisions: edits.flatMap(({ revisions }) => revisions),
    };
    return this.prepareJournalledOps({
      state,
      ops,
      postSelection,
      stagedApplied,
      semantic,
    });
  }

  prepareOps(
    state: EditorState,
    ops: readonly DocumentOp[],
    postSelection: CanonicalSelection,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    return this.prepareJournalledOps({ state, ops, postSelection });
  }

  private prepareJournalledOps({
    state,
    ops,
    postSelection,
    stagedApplied,
    semantic = "structure",
  }: JournalledOpsOptions): Result<CanonicalCommit, CanonicalSessionError> {
    const story = postSelection.anchor.story;
    const projected = this.projectStory(story);
    if (projected.isErr()) return projected;
    const preSelection = projected.value.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    if (ops.length === 0 || stagedApplied?.inverse.length === 0)
      return noChange("The intent makes no document change.");
    const preDocument = this.currentDocument;
    return this.stage({
      state,
      story,
      ops,
      selection: postSelection,
      origin: "input",
      history: "opened",
      ...(stagedApplied === undefined ? {} : { stagedApplied }),
      onPublish: (inverse, version) => {
        const entry = {
          type: "applied",
          ops,
          inverse,
          preSelection: preSelection.value,
          postSelection,
          preDocument,
          postDocument: this.currentDocument,
          version,
          origin: "input",
          semantic,
          grouping: "isolated",
          time: Date.now(),
          boundary: this.groupingBoundary,
        } as const satisfies AppliedJournalEntry;
        this.applied.push({
          entries: [entry],
          undo: { type: "entries" },
          preSelection: preSelection.value,
          postSelection,
        });
        this.undone.length = 0;
      },
    });
  }

  prepareSplit(
    state: EditorState,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const from = this.projection.inputAddressAt(state.selection.from);
    if (from.isErr()) return from;
    const to = this.projection.inputAddressAt(state.selection.to);
    if (to.isErr()) return to;
    const source = this.projection.paragraph(from.value.blockId);
    const styles = this.styles ?? this.currentDocument.package.styles;
    const nextStyle = styles?.styles.find(
      ({ styleId }) => styleId === source?.source.formatting?.styleId,
    )?.next;
    const fresh = this.freshBlockId();
    if (fresh.isErr()) return fresh;
    const intents: EditorIntent[] = [
      {
        type: "splitParagraph",
        at: from.value,
        to: to.value,
        newBlockId: fresh.value,
      },
    ];
    if (
      state.selection.empty &&
      source !== undefined &&
      from.value.offset === source.text.length &&
      nextStyle !== undefined &&
      nextStyle !== source.source.formatting?.styleId
    )
      intents.push({
        type: "formatParagraph",
        at: { ...from.value, offset: 0 },
        patch: { styleId: nextStyle },
      });
    return this.prepareIntents(state, { intents });
  }

  prepareJoin(
    state: EditorState,
    direction: "backward" | "forward",
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    if (!state.selection.empty)
      return refuse("Join requires a collapsed text selection.");
    const at = this.projection.inputAddressAt(state.selection.from);
    if (at.isErr()) return at;
    const groups = editorParagraphGroups(this.currentDocument, at.value.story);
    const index = groups.findIndex(({ paragraphs }) =>
      paragraphs.some((paragraph) => paragraph.paraId === at.value.blockId),
    );
    if (index < 0 || (direction === "backward" && index === 0))
      return noChange("There is no adjacent paragraph to join.");
    const firstGroup = groups.at(direction === "backward" ? index - 1 : index);
    const secondGroup = groups.at(direction === "backward" ? index : index + 1);
    const first = firstGroup?.paragraphs.at(-1);
    const second = secondGroup?.paragraphs.at(0);
    if (!first?.paraId || !second?.paraId)
      return noChange("There is no adjacent paragraph to join.");
    if (
      this.projection.paragraph(first.paraId)?.container !==
      this.projection.paragraph(second.paraId)?.container
    )
      return noChange("There is no adjacent paragraph to join.");
    // The same visible test that routed the key here. An input address moves past
    // tracked-deleted text, so it cannot say whether the caret is at the boundary.
    if (!isCanonicalJoinBoundary(state, direction))
      return refuse("Join requires a caret at a paragraph boundary.");
    return this.prepareIntent(state, {
      type: "joinParagraphs",
      story: at.value.story,
      blockId: first.paraId,
      nextBlockId: second.paraId,
    });
  }

  prepareResolve(
    state: EditorState,
    {
      revisionIds,
      resolution,
      before = [],
    }: {
      revisionIds: readonly number[];
      resolution: "accept" | "reject";
      /** Operations applied ahead of the resolution, in the same commit. */
      before?: readonly DocumentOp[];
    },
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    const preSelection = this.projection.selectionAt(state);
    if (preSelection.isErr()) return preSelection;
    const resolve = (ids: readonly number[]) =>
      [
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: ids,
          decision: resolution,
        },
      ] as const;
    // An id names its whole listed change: one split by a comment or a
    // hyperlink carries several ids, and resolving fewer leaves part pending.
    const changes = getTrackedChangeGroupsForIds(state.doc, revisionIds);
    // A change whose resolution breaks a lock stays pending; the rest resolve.
    let firstHeld: LockedControlViolation | null = null;
    const heldChanges = changes.filter((ids) => {
      const alone = applyDocumentOps(this.currentDocument, resolve(ids));
      if (alone.isErr()) return false;
      const violation = lockViolationBetween(
        this.currentDocument,
        alone.value.document,
      );
      firstHeld ??= violation;
      return violation !== null;
    });
    if (firstHeld !== null && heldChanges.length === changes.length)
      return lockRefusal(firstHeld);
    const lockHeld = heldChanges.flat();
    const resolved = changes.filter((ids) => !heldChanges.includes(ids)).flat();
    const ops = [
      ...before,
      ...(resolved.length === 0 ? [] : resolve(resolved)),
    ];
    // Resolution can retire the selected paragraph; choose a valid result anchor before staging.
    const applied = applyDocumentOps(this.currentDocument, ops);
    if (applied.isErr()) return refuse(applied.error.message);
    const projected = project({
      document: applied.value.document,
      styles: this.styles,
      caches: this.caches,
      previous: this.currentProjection,
    });
    if (projected.isErr()) return projected;
    const surviving = projected.value.paragraph(
      preSelection.value.head.blockId,
    );
    // The first paragraph may sit inside a table that opens the story.
    const first = projected.value.firstParagraph;
    if (first === undefined) return refuse("Resolution left no paragraph.");
    let offset = surviving
      ? Math.min(preSelection.value.head.offset, surviving.text.length)
      : 0;
    if (surviving && splitsSurrogatePair(surviving.text, offset)) offset -= 1;
    const caret = {
      story: OP_STORIES.MAIN,
      blockId: surviving?.blockId ?? first.blockId,
      offset,
    };
    const prepared = this.prepareJournalledOps({
      state,
      ops,
      postSelection: { type: "text", anchor: caret, head: caret },
      stagedApplied: applied.value,
    });
    return prepared.isOk() && lockHeld.length > 0
      ? Result.ok({ ...prepared.value, lockHeld })
      : prepared;
  }

  /**
   * Accept or reject one pending suggestion. Rejecting it rejects its
   * revisions and removes the comments it proposes. Accepting it adopts its
   * revisions as ordinary tracked changes by `author`, keeping their ids, as
   * the editor does: the change then waits for review like any other. Its
   * comments stay as ordinary comments.
   */
  prepareSuggestion(
    state: EditorState,
    {
      suggestionId,
      resolution,
      author,
      date,
    }: {
      suggestionId: string;
      resolution: "accept" | "reject";
      author: string;
      date: string;
    },
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const infos = suggestionRevisionInfos(
      this.currentDocument,
      OP_STORIES.MAIN,
      suggestionId,
    );
    const commentIds = suggestedCommentIds(this.currentDocument, suggestionId);
    if (infos.length === 0 && commentIds.length === 0)
      return noChange(`No pending suggestion ${suggestionId}.`);
    if (resolution === "reject")
      return this.prepareResolve(state, {
        revisionIds: infos.map(({ id }) => id),
        resolution: "reject",
        // The comments go first, while the text they cover is still there.
        before: commentIds.map((id) => ({
          type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
          id,
          scope: "thread",
        })),
      });
    const checked = this.checkState(state);
    if (checked.isErr()) return checked;
    const selection = this.projection.selectionAt(state);
    if (selection.isErr()) return selection;
    const comments = this.currentDocument.package.document.comments ?? [];
    const ops: DocumentOp[] = [];
    if (infos.length > 0)
      ops.push({
        type: DOCUMENT_OP_TYPES.SET_REVISION_INFO,
        story: OP_STORIES.MAIN,
        infos: infos.map(({ id }) => ({ id, author, date })),
      });
    for (const comment of comments)
      if (commentIds.includes(comment.id))
        ops.push({
          type: DOCUMENT_OP_TYPES.UPDATE_COMMENT_CONTENT,
          id: comment.id,
          content: comment.content,
          patch: { suggestionId: null },
        });
    return this.prepareJournalledOps({
      state,
      ops,
      postSelection: selection.value,
    });
  }

  projectStory(
    story: OpStory,
  ): Result<CanonicalProjection, CanonicalSessionError> {
    if (story === OP_STORIES.MAIN) return Result.ok(this.projection);
    const key = storyProjectionKey(story);
    const body = findStoryBody(this.currentDocument, story);
    if (body === undefined) {
      this.storyProjections.delete(key);
      return refuse("The canonical story no longer exists.");
    }
    const cached = this.storyProjections.get(key);
    const theme = this.currentDocument.package.theme;
    if (cached?.content === body.content && cached.theme === theme)
      return Result.ok(cached.projection);
    const projected = project({
      document: this.currentDocument,
      styles: this.styles,
      caches: this.caches,
      story,
      ...(cached === undefined ? {} : { previous: cached.projection }),
    });
    if (projected.isOk())
      this.storyProjections.set(key, {
        content: body.content,
        theme,
        projection: projected.value,
      });
    return projected;
  }

  prepareOperations(
    state: EditorState,
    ops: readonly DocumentOp[],
  ): Result<CanonicalCommit, CanonicalSessionError> {
    if (ops.length === 0) return refuse("The operation batch is empty.");
    const selected = this.projection.selectionAt(state);
    if (selected.isErr()) return selected;
    let preparedOps = ops;
    if (
      ops.every(
        (op) =>
          op.type === DOCUMENT_OP_TYPES.ADD_NOTE ||
          op.type === DOCUMENT_OP_TYPES.REMOVE_NOTE,
      )
    ) {
      const tracked: DocumentOp[] = [];
      let document = this.currentDocument;
      for (const op of ops) {
        if (
          op.type !== DOCUMENT_OP_TYPES.ADD_NOTE &&
          op.type !== DOCUMENT_OP_TYPES.REMOVE_NOTE
        )
          panic("A note operation batch contains only note operations.");
        const allocation = this.allocateIntentIds(document, {
          type: "replaceFragment",
          from: { ...op.at, offset: 0 },
          to: { ...op.at, offset: Number.MAX_SAFE_INTEGER },
          paragraphs: [
            {
              type: "paragraph",
              content: [
                {
                  type: "run",
                  content: [{ type: "break", breakType: "page" }],
                },
              ],
            },
          ],
          openStart: 0,
          openEnd: 0,
        });
        const next = {
          ...op,
          ...(this.mode.type === "suggesting"
            ? {
                revision: {
                  id: allocation.revisionId,
                  author: this.mode.author,
                  date: new Date().toISOString(),
                },
              }
            : {}),
          newIds: allocation.newIds,
        } as const;
        if (ops.length > 1) {
          const applied = applyDocumentOp(
            cloneDocumentWithParagraphPropertySources(document),
            structuredClone(next),
          );
          if (applied.isErr()) return refuse(applied.error.message);
          document = applied.value.document;
        }
        tracked.push(next);
      }
      preparedOps = tracked;
    }
    const addedNotes = preparedOps.filter(
      (op) => op.type === DOCUMENT_OP_TYPES.ADD_NOTE,
    );
    // The caret follows the reference inserted at the selection's end, past every
    // reference the batch inserts at or before it in the same paragraph.
    const atSelection = addedNotes.find(({ at }) =>
      [selected.value.anchor, selected.value.head].some(
        (end) =>
          sameStory(at.story, end.story) &&
          at.blockId === end.blockId &&
          at.offset === end.offset,
      ),
    );
    const caret =
      atSelection === undefined
        ? undefined
        : {
            ...atSelection.at,
            offset:
              atSelection.at.offset +
              addedNotes.filter(
                ({ at }) =>
                  sameStory(at.story, atSelection.at.story) &&
                  at.blockId === atSelection.at.blockId &&
                  at.offset <= atSelection.at.offset,
              ).length,
          };
    const postSelection =
      addedNotes.length === preparedOps.length && caret !== undefined
        ? { type: "text" as const, anchor: caret, head: caret }
        : selected.value;
    const preDocument = this.currentDocument;
    return this.stage({
      state,
      ops: preparedOps,
      selection: postSelection,
      origin: "input",
      history: "opened",
      onPublish: (inverse, version) => {
        this.applied.push({
          entries: [
            {
              type: "applied",
              ops: preparedOps,
              inverse,
              preSelection: selected.value,
              postSelection,
              preDocument,
              postDocument: this.currentDocument,
              version,
              origin: "input",
              semantic: "structure",
              grouping: "isolated",
              time: Date.now(),
              boundary: this.groupingBoundary,
            },
          ],
          undo: { type: "entries" },
          preSelection: selected.value,
          postSelection,
        });
        this.undone.length = 0;
      },
    });
  }

  prepareUndo(
    state: EditorState,
    story: OpStory = OP_STORIES.MAIN,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const group = this.applied.at(-1);
    if (group === undefined)
      return refuse("There is no canonical edit to undo.");
    const propertySourceDocument = group.entries.at(0)?.preDocument;
    return this.stage({
      state,
      story,
      ops:
        group.undo.type === "replayed"
          ? group.undo.inverse
          : group.entries.toReversed().flatMap((entry) => entry.inverse),
      selection: group.preSelection,
      origin: "undo",
      history: "undo",
      ...(propertySourceDocument === undefined
        ? {}
        : { propertySourceDocument }),
      onPublish: (redoOps) => {
        this.applied.pop();
        this.undone.push({ type: "undone", group, redoOps });
        this.breakUndoGroup();
      },
    });
  }

  prepareRedo(
    state: EditorState,
    story: OpStory = OP_STORIES.MAIN,
  ): Result<CanonicalCommit, CanonicalSessionError> {
    const undone = this.undone.at(-1);
    if (undone === undefined)
      return refuse("There is no canonical edit to redo.");
    const propertySourceDocument = undone.group.entries.at(-1)?.postDocument;
    return this.stage({
      state,
      story,
      ops: undone.redoOps,
      selection: undone.group.postSelection,
      origin: "redo",
      history: "redo",
      ...(propertySourceDocument === undefined
        ? {}
        : { propertySourceDocument }),
      onPublish: (inverse) => {
        this.undone.pop();
        this.applied.push({
          entries: undone.group.entries,
          undo: { type: "replayed", inverse },
          preSelection: undone.group.preSelection,
          postSelection: undone.group.postSelection,
        });
        this.breakUndoGroup();
      },
    });
  }

  private stage({
    state,
    ops,
    selection,
    origin,
    history,
    onPublish,
    propertySourceDocument,
    stagedApplied,
    story = OP_STORIES.MAIN,
    previousProjection,
  }: StageOptions): Result<CanonicalCommit, CanonicalSessionError> {
    const previous =
      previousProjection === undefined
        ? this.projectStory(story)
        : Result.ok(previousProjection);
    if (previous.isErr()) return previous;
    const checked = this.checkState(state, previous.value);
    if (checked.isErr()) return checked;
    const applied =
      stagedApplied === undefined
        ? applyDocumentOps(this.currentDocument, ops)
        : Result.ok(stagedApplied);
    if (applied.isErr()) return refuse(applied.error.message);
    if (applied.value.inverse.length === 0)
      return noChange("The intent makes no document change.");
    // Judged on the result, so every path into the session meets one rule.
    const locked = lockViolationBetween(
      this.currentDocument,
      applied.value.document,
    );
    if (locked !== null) return lockRefusal(locked);
    if (propertySourceDocument !== undefined)
      applied.value.document = preservePropertySources({
        target: applied.value.document,
        source: propertySourceDocument,
        indentation: "restore",
      });
    const normalized = normalizeCanonicalListRendering(
      applied.value.document,
      this.caches.lists,
    );
    applied.value.document = normalized.document;
    applied.value.inverse = normalized.inverse.concat(applied.value.inverse);
    const changed = changedStories({
      document: applied.value.document,
      previous: this.currentDocument,
    });
    if (
      unsupportedEditReason({
        previous: this.currentDocument,
        document: applied.value.document,
        stories: changed,
      }) !== null
    )
      return Result.err(
        new CanonicalSessionError({
          ground: { gap: CANONICAL_GAP.storyContentProjection },
          reason: "refused",
          message:
            "The operations produce unsupported canonical story content.",
        }),
      );
    applied.value.document = preservePropertySources({
      target: applied.value.document,
      source: this.currentDocument,
      stories: changed,
    });
    if (propertySourceDocument !== undefined)
      applied.value.document = preservePropertySources({
        target: applied.value.document,
        source: propertySourceDocument,
      });
    // Every main-story paragraph the session holds is its id's owner, so
    // only the blocks this edit put in place can take an owner's sources.
    applied.value.document = withParagraphCopies({
      document: applied.value.document,
      story: OP_STORIES.MAIN,
      copies: ownedSourceCopies(
        this.sourceOwners,
        changedWindow(
          this.currentDocument.package.document.content,
          applied.value.document.package.document.content,
        ).added,
      ),
    });
    const remembered = changedWindow(
      this.currentDocument.package.document.content,
      applied.value.document.package.document.content,
    ).added;
    const projectedStory = findStoryBody(applied.value.document, story)
      ? story
      : OP_STORIES.MAIN;
    const projected = project({
      document: applied.value.document,
      styles: this.styles,
      caches: this.caches,
      story: projectedStory,
      previous: sameStory(projectedStory, OP_STORIES.MAIN)
        ? this.currentProjection
        : previous.value,
    });
    if (projected.isErr()) return projected;
    const bodyProjection = (() => {
      if (sameStory(projectedStory, OP_STORIES.MAIN)) return projected;
      if (ops.some(operationChangesBodyProjection))
        return project({
          document: applied.value.document,
          styles: this.styles,
          caches: this.caches,
          previous: this.currentProjection,
        });
      return Result.ok(this.currentProjection);
    })();
    if (bodyProjection.isErr()) return bodyProjection;
    const transaction = state.tr;
    const before = state.doc;
    const after = projected.value.doc;
    let prefix = 0;
    let from = 0;
    while (
      prefix < before.childCount &&
      prefix < after.childCount &&
      before.child(prefix).eq(after.child(prefix))
    ) {
      from += before.child(prefix).nodeSize;
      prefix += 1;
    }
    let oldEnd = before.childCount;
    let newEnd = after.childCount;
    let to = before.content.size;
    while (
      oldEnd > prefix &&
      newEnd > prefix &&
      before.child(oldEnd - 1).eq(after.child(newEnd - 1))
    ) {
      oldEnd -= 1;
      newEnd -= 1;
      to -= before.child(oldEnd).nodeSize;
    }
    if (oldEnd !== prefix || newEnd !== prefix) {
      const replacement: PMNode[] = [];
      for (let index = prefix; index < newEnd; index += 1)
        replacement.push(after.child(index));
      transaction.replaceWith(from, to, Fragment.fromArray(replacement));
    }
    if (!transaction.doc.eq(projected.value.doc))
      transaction.replaceWith(
        0,
        transaction.doc.content.size,
        projected.value.doc.content,
      );
    // The document node's own attributes hold what the last section starts with.
    for (const [name, value] of Object.entries(projected.value.doc.attrs)) {
      if (transaction.doc.attrs[name] !== value) {
        transaction.setDocAttribute(name, value);
      }
    }
    if (
      sameStory(selection.anchor.story, projectedStory) &&
      sameStory(selection.head.story, projectedStory)
    ) {
      if (
        !restoreCanonicalSelection({
          transaction,
          projection: projected.value,
          selection,
          unavailable: { type: "near", anchor: state.selection.anchor },
        })
      )
        return refuse(
          "Canonical history cannot restore the selected inline atom.",
        );
    }
    // The transaction now ends on a document equal to the projection.
    const held = projected.value.heldAs(transaction.doc);
    const heldBody =
      bodyProjection.value === projected.value ? held : bodyProjection.value;
    transaction.setMeta(CANONICAL_PROJECTION_META, {
      type: "canonical",
      origin,
      version: this.version + 1,
    });
    if (ops.some(operationChangesPackage)) markPackageChange(transaction);
    transaction.setMeta("addToHistory", false);
    const baseVersion = this.version;
    const baseBoundary = this.groupingBoundary;
    authorizedProjections.set(transaction, {
      session: this,
      version: baseVersion,
      doc: projected.value.doc,
    });
    return Result.ok({
      transaction,
      document: applied.value.document,
      projection: held,
      bodyProjection: heldBody,
      selection,
      touched: applied.value.touched,
      version: baseVersion + 1,
      origin,
      history,
      publish: () => {
        if (this.isComposing || this.groupingBoundary !== baseBoundary)
          return refuse(
            "The staged canonical commit crossed a composition or selection boundary.",
          );
        if (this.version !== baseVersion)
          return refuse("The staged canonical commit is stale.");
        if (!isCanonicalProjectionTransaction(transaction, this)) {
          return refuse(
            "The staged canonical projection was changed after preparation.",
          );
        }
        for (const blockId of applied.value.touched.inserted)
          this.allocatedBlockIds.add(blockId.toUpperCase());
        for (const op of ops) {
          if (op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK)
            this.allocatedBlockIds.add(op.newBlockId.toUpperCase());
          if (op.type === DOCUMENT_OP_TYPES.ADD_NOTE) {
            for (const blockId of paragraphIdsIn(op.note.content))
              this.allocatedBlockIds.add(blockId.toUpperCase());
          }
        }
        this.advanceBlockId();
        this.rememberSources(remembered);
        if (projectedStory !== OP_STORIES.MAIN) {
          const body = findStoryBody(applied.value.document, projectedStory);
          if (body === undefined)
            panic("A staged story projection lost its source.");
          this.storyProjections.set(storyProjectionKey(projectedStory), {
            content: body.content,
            theme: applied.value.document.package.theme,
            projection: held,
          });
        }
        this.currentSelection = selection;
        this.currentDocument = applied.value.document;
        this.currentProjection = heldBody;
        // Save changes are cumulative against originalBuffer, including undo/redo.
        // They are published with the model, never with a staged projection.
        for (const ids of [
          applied.value.touched.modified,
          applied.value.touched.inserted,
          applied.value.touched.removed,
        ]) {
          for (const id of ids) this.saveTouched.add(id);
        }
        // Comment operations touch only the comments part and keep the paragraph
        // splice path; every other operation that inserts or removes blocks, or
        // changes package parts or a non-body story, invalidates it.
        const bodyOps = ops.filter(
          (op) => !COMMENT_OPERATION_TYPES.has(op.type),
        );
        if (
          bodyOps.some(operationChangesStructure) ||
          bodyOps.some(operationChangesPackage) ||
          (bodyOps.length > 0 &&
            (applied.value.touched.inserted.length > 0 ||
              applied.value.touched.removed.length > 0))
        )
          this.saveStructure = "changed";
        this.currentVersion = baseVersion + 1;
        onPublish(applied.value.inverse, this.currentVersion);
        return Result.ok(undefined);
      },
    });
  }
}

export type { CanonicalProjection, CanonicalSession };
/** Project one story, reusing what the caches kept when they are given. */
export { project as projectCanonicalStory };

export const createCanonicalSession = (
  document: ReadonlyDocument,
  styles?: DeepReadonly<StyleDefinitions> | null,
): Result<CanonicalSession, CanonicalSessionError> => {
  const owned = cloneDocumentWithParagraphPropertySources(document);
  const source = normalizeCanonicalSource(owned);
  const unsupported = unsupportedSeedReason(source.document);
  if (unsupported !== null) return refuse(unsupported);
  const identified = preservePropertySourcesWholeStories({
    target: normalizeForOps(source.document),
    source: owned,
  });
  const validated = validateOpsDocument(identified);
  if (validated.isErr()) return refuse(validated.error.message);
  const caches = newStoryProjectionCaches();
  const normalized = preservePropertySourcesWholeStories({
    target: normalizeCanonicalListRendering(identified, caches.lists).document,
    source: owned,
  });
  const projected = project({ document: normalized, styles, caches });
  if (projected.isErr()) return projected;
  return Result.ok(
    new CanonicalSession({
      document: normalized,
      projection: projected.value,
      styles,
      sourceRestores: canonicalSourceRestores(normalized, source.normalized),
      caches,
    }),
  );
};

const hasDeletedParagraphMark = (node: PMNode): boolean => {
  const mark: unknown = node.attrs["pPrMark"];
  return (
    typeof mark === "object" &&
    mark !== null &&
    "kind" in mark &&
    (mark.kind === "del" || mark.kind === "moveFrom")
  );
};

/** Visible graphemes concatenate across retained deletions and deleted paragraph marks. */
const visibleDeletionContext = (state: EditorState) => {
  const { selection } = state;
  const paragraphDepth = selection.$from.depth;
  const containerDepth = paragraphDepth - 1;
  if (
    containerDepth < 0 ||
    selection.$from.node(paragraphDepth).type.name !== "paragraph"
  )
    return refuse("Canonical deletion requires a caret inside a paragraph.");
  const container = selection.$from.node(containerDepth);
  const containerStart = selection.$from.start(containerDepth);
  const index = selection.$from.index(containerDepth);
  let first = index;
  while (first > 0 && hasDeletedParagraphMark(container.child(first - 1)))
    first -= 1;
  let last = index;
  while (
    last + 1 < container.childCount &&
    hasDeletedParagraphMark(container.child(last))
  )
    last += 1;
  let text = "";
  const physicalStarts: number[] = [];
  const physicalEnds: number[] = [];
  let visibleOffset = 0;
  let issue: CanonicalSessionError | undefined;
  container.forEach((paragraph, paragraphOffset, paragraphIndex) => {
    if (paragraphIndex < first || paragraphIndex > last) return;
    const native = canonicalNativeCells(paragraph);
    if (native.isErr()) {
      issue = new CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.dispatch },
        reason: "refused",
        message: native.error.message,
      });
      return;
    }
    for (const cell of native.value) {
      if (
        cell.type === "zeroWidth" ||
        cell.node.marks.some((mark) => mark.type.name === "deletion")
      )
        continue;
      const value = cell.type === "text" ? cell.text : "\uFFFC";
      const from = containerStart + paragraphOffset + 1 + cell.from;
      const to = containerStart + paragraphOffset + 1 + cell.to;
      if (
        cell.type !== "text" &&
        selection.from > from &&
        selection.from < to
      ) {
        issue = new CanonicalSessionError({
          ground: { gap: CANONICAL_GAP.dispatch },
          reason: "refused",
          message: "Character deletion cannot start inside an inline atom.",
        });
        return;
      }
      for (let unit = 0; unit < value.length; unit += 1) {
        const position = from + (cell.type === "text" ? unit : 0);
        physicalStarts.push(position);
        text += value.charAt(unit);
        physicalEnds.push(cell.type === "text" ? position + 1 : to);
        if (position < selection.from) visibleOffset += 1;
      }
    }
  });
  return issue === undefined
    ? Result.ok({ text, physicalStarts, physicalEnds, visibleOffset })
    : Result.err(issue);
};

export const isCanonicalJoinBoundary = (
  state: EditorState,
  direction: "backward" | "forward",
): boolean => {
  if (!(state.selection instanceof TextSelection) || !state.selection.empty)
    return false;
  const context = visibleDeletionContext(state);
  if (context.isErr()) return false;
  const { text, visibleOffset } = context.value;
  return visibleOffset === (direction === "backward" ? 0 : text.length);
};

/** Native character deletion consumes one whole visible grapheme. */
export const deletionRange = (
  state: EditorState,
  direction: "backward" | "forward",
): Result<{ from: number; to: number }, CanonicalSessionError> => {
  const { selection } = state;
  if (
    selection instanceof NodeSelection &&
    selection.node.isInline &&
    selection.node.isAtom &&
    selection.node.type.name !== "noteMarker"
  )
    return Result.ok({ from: selection.from, to: selection.to });
  if (selection instanceof CellSelection)
    return Result.ok({ from: selection.from, to: selection.to });
  if (selection instanceof AllSelection)
    return Result.ok(canonicalSelectionRange(state));
  if (!(selection instanceof TextSelection))
    return refuse(
      "Canonical deletion requires text, an inline atom, or the whole document selection.",
    );
  if (!selection.empty)
    return Result.ok({ from: selection.from, to: selection.to });
  const parent = selection.$from.parent;
  const offset = selection.$from.parentOffset;
  const adjacent =
    direction === "backward"
      ? parent.childBefore(offset).node
      : parent.childAfter(offset).node;
  if (adjacent?.type.name === "noteMarker")
    return refuse("Automatic note marks cannot be deleted as characters.");
  const context = visibleDeletionContext(state);
  if (context.isErr()) return context;
  const { text, physicalStarts, physicalEnds, visibleOffset } = context.value;
  if (splitsSurrogatePair(text, visibleOffset))
    return refuse("The deletion would split a surrogate pair.");
  if (splitsGraphemeCluster(text, visibleOffset))
    return refuse(
      "Character deletion requires a caret at a grapheme boundary.",
    );
  if (
    (direction === "backward" && visibleOffset === 0) ||
    (direction === "forward" && visibleOffset === text.length)
  )
    return refuse("There is no visible character to delete in this paragraph.");
  let boundary = visibleOffset + (direction === "backward" ? -1 : 1);
  while (splitsGraphemeCluster(text, boundary))
    boundary += direction === "backward" ? -1 : 1;
  const from = physicalStarts.at(
    direction === "backward" ? boundary : visibleOffset,
  );
  const to = physicalEnds.at(
    (direction === "backward" ? visibleOffset : boundary) - 1,
  );
  if (from === undefined || to === undefined)
    panic("Canonical deletion lost a grapheme boundary.");
  return Result.ok({ from, to });
};
