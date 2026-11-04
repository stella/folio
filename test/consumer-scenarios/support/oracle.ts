/**
 * The requested-outcome oracle: what an applied operation must have done,
 * worked out from the document before it and the request alone, never from
 * folio's apply code. A receipt that says `applied` is checked against it
 * after a save and a reopen: in `direct` mode as saved, in `tracked-changes`
 * mode with every change accepted (and rejecting every change must give the
 * document before back). A refused operation must leave the document as it
 * was.
 *
 * The model keeps block texts and requested properties in reading order, plus
 * an independent cell grid and main-story links on unchanged paragraphs. A batch
 * resolves every operation against the document as it was read (see the
 * batch-claims contract), so a batch's expectations compose in order over
 * the same pre-state, and an operation the engine refused contributes
 * nothing. An operation the model cannot predict (a type marked `null` in
 * `EXPECTATIONS`, or a target it cannot place) leaves that batch's text
 * uncompared; supported geometry and links are still checked.
 * FOLIO_ORACLE_GAPS=1 prints each gap.
 */

import assert from "node:assert/strict";

import type { FolioContentStatedNumbering } from "@stll/folio-core/compare/content-types";
import {
  FOLIO_DOCUMENT_OPERATION_TYPES,
  type FolioDocumentOperationType,
  type FolioDocumentStoryHandle,
  getFolioDocumentOperationRemovedBlockIds,
  getTrackedFolioDocumentOperationRemovedBlockIds,
  inspectDocumentStylesFromDocx,
  parseDocx,
} from "@stll/folio-core/server";

import { sequentialGroups } from "./batch-replay-order.ts";
import { recordFeatureHit, recordHit, type StepKind } from "./coverage.ts";
import { openReviewer, toArrayBuffer } from "./documents.ts";
import {
  operationSelection,
  targetFeatureSignature,
} from "./feature-coverage.ts";
import {
  captureLinks,
  comparePreservedLinks,
  type LinkSnapshot,
} from "./link-oracle.ts";
import {
  type AppliedOperation,
  coreBatch,
  type Mode,
  type Operation,
} from "./operations.ts";
import {
  applyTableOperation,
  compareTableGeometry,
  modelFromRows,
  removeTextBoxTables,
  type TableModel,
  UnsupportedTableExpectation,
} from "./table-oracle.ts";
import {
  blocksOfStory,
  type Feature,
  type FeatureIndex,
  featureIndex,
  storyKindOf,
  touchesSurrogate,
} from "./targets.ts";

type Reviewer = Awaited<ReturnType<typeof openReviewer>>;
type ParsedDocument = Awaited<ReturnType<typeof parseDocx>>;
type ParsedBlock = ParsedDocument["package"]["document"]["content"][number];
type ParsedInline = Extract<
  ParsedBlock,
  { type: "paragraph" }
>["content"][number];
type ParsedNumbering = NonNullable<
  NonNullable<
    Extract<ParsedBlock, { type: "paragraph" }>["formatting"]
  >["numPr"]
>;
/** Expected numbering uses input numeric facts, without invoking the production allocator. */
type Numbering =
  | Exclude<ParsedNumbering, { kind: "reference" }>
  | Pick<
      Extract<ParsedNumbering, { kind: "reference" }>,
      "kind" | "numId" | "ilvl"
    >
  | undefined;
type NumberingFacts = {
  styles: Map<string, Numbering>;
  direct: Map<string, Numbering>;
  markerLevels: Map<number, Set<number>>;
};

/** Fold stated slots independently of the engine's numbering resolver. */
const numberingOver = (base: Numbering, stated: Numbering): Numbering => {
  if (stated === undefined) return base;
  switch (stated.kind) {
    case "none":
      return stated;
    case "reference":
      return {
        ...stated,
        ilvl: stated.ilvl ?? (base?.kind === "none" ? undefined : base?.ilvl),
      };
    case "levelOnly":
      if (base?.kind === "reference") return { ...base, ilvl: stated.ilvl };
      return base?.kind === "none" ? base : stated;
    default: {
      const unhandled: never = stated;
      throw new Error(`Unhandled numbering ${JSON.stringify(unhandled)}`);
    }
  }
};

type LiveNumberingLevel = ReturnType<
  Reviewer["readNumberingDefinitions"]
>[number];

/**
 * The marker levels of the main story's lists that only their paragraphs
 * define so far (lists its operations started, pending or applied). The
 * reviewer reads each one whole, so a reference to any of its levels, and a
 * paragraph anywhere among the other insertions beside its anchor, is
 * classified by the list's own definition.
 */
const addLiveListMarkerLevels = (
  facts: NumberingFacts,
  definitions: readonly LiveNumberingLevel[],
): void => {
  const lists = new Map<number, Map<number, LiveNumberingLevel>>();
  for (const definition of definitions) {
    if (facts.markerLevels.has(definition.numId)) continue;
    const levels =
      lists.get(definition.numId) ?? new Map<number, LiveNumberingLevel>();
    levels.set(definition.level, definition);
    lists.set(definition.numId, levels);
  }
  for (const [numId, levels] of lists) {
    facts.markerLevels.set(
      numId,
      new Set(
        [...levels.values()]
          .filter((level) =>
            showsMarker(
              {
                text: level.levelText,
                format: level.format,
                legal: false,
                hidden: false,
              },
              (ilvl) => levels.get(ilvl)?.format,
            ),
          )
          .map(({ level }) => level),
      ),
    );
  }
};

/**
 * A header, footer or note story starts its lists against the package alone,
 * so its new ids can repeat another story's and the reviewer's numbering
 * read (the main story's) does not define them. The level a pending
 * insertion beside its anchor states is all that marks such a list.
 */
const addPendingInsertionListLevels = (
  facts: NumberingFacts,
  reviewer: Reviewer,
  blocks: NonNullable<ReturnType<Reviewer["snapshotStory"]>>["blocks"],
): void => {
  const pendingInsertionBlockIds = new Set<string>();
  for (const { operation } of reviewer.exportPendingSuggestions()) {
    if (
      operation.type !== "insertAfterBlock" &&
      operation.type !== "insertBeforeBlock"
    )
      continue;
    if (operation.numbering?.kind !== "newList") continue;
    let anchorIndex = blocks.findIndex(({ id }) => id === operation.blockId);
    if (anchorIndex < 0) continue;
    const direction = operation.type === "insertAfterBlock" ? 1 : -1;
    const anchorTable = blocks[anchorIndex]?.table;
    if (anchorTable) {
      while (
        blocks.at(anchorIndex + direction)?.table?.outerTableIndex ===
        anchorTable.outerTableIndex
      ) {
        anchorIndex += direction;
      }
    }
    const blockAt = (index: number) =>
      index >= 0 ? blocks.at(index) : undefined;
    let insertedIndex = anchorIndex + direction;
    while (pendingInsertionBlockIds.has(blockAt(insertedIndex)?.id ?? "")) {
      insertedIndex += direction;
    }
    const inserted = blockAt(insertedIndex) ?? null;
    if (
      !inserted ||
      inserted.kind === "diagnostic" ||
      inserted.statedNumbering.kind !== "reference"
    )
      continue;
    pendingInsertionBlockIds.add(inserted.id);
    const level = inserted.statedNumbering.ilvl ?? 0;
    const levels =
      facts.markerLevels.get(inserted.statedNumbering.numId) ??
      new Set<number>();
    levels.add(level);
    facts.markerLevels.set(inserted.statedNumbering.numId, levels);
  }
};

/** Capture the authored story, including proposals omitted from the saved model. */
const liveNumberingFacts = (
  reviewer: Reviewer,
  story: Story,
): NumberingFacts => {
  const facts = numberingFactsFromDocument(reviewer.toDocument());
  const snapshot = reviewer.snapshotStory(story);
  if (!snapshot) return facts;
  const { anchors, blocks } = snapshot;
  if (story.type === "main") {
    addLiveListMarkerLevels(facts, reviewer.readNumberingDefinitions());
  } else {
    addPendingInsertionListLevels(facts, reviewer, blocks);
  }
  for (const block of blocks) {
    if (block.kind === "diagnostic") continue;
    const anchor = anchors[block.id];
    assert.ok(anchor, `Live block ${block.id} has no anchor`);
    facts.direct.set(
      block.id,
      block.statedNumbering.kind === "inherit"
        ? undefined
        : block.statedNumbering,
    );
  }
  return facts;
};

type MarkerLevel = {
  text: string;
  format: string;
  legal: boolean;
  hidden: boolean;
};

/** Whether a numbering level shows a marker, given the formats of its list's other levels. */
const showsMarker = (
  { text, format, legal, hidden }: MarkerLevel,
  formatAt: (ilvl: number) => string | undefined,
): boolean => {
  if (hidden) return false;
  // A legal level counts in decimal whatever its own format says.
  if (text.length === 0) return legal || format !== "none";
  // Only marker presence matters here; counter values and numbering glyphs do not.
  return (
    text
      .replace(/%([1-9])/gu, (_, slot: string) =>
        !legal && formatAt(Number(slot) - 1) === "none" ? "" : "1",
      )
      .trim().length > 0
  );
};

type ParsedParagraph = Extract<ParsedBlock, { type: "paragraph" }>;

/** Every paragraph of `document`'s stories, text boxes and cells included, in document order. */
const visitParsedParagraphs = (
  document: ParsedDocument,
  visitParagraph: (paragraph: ParsedParagraph) => void,
): void => {
  const visitInline = (node: ParsedInline): void => {
    switch (node.type) {
      case "run":
        for (const content of node.content) {
          if (content.type === "shape" && content.shape.textBody)
            visit(content.shape.textBody.content);
        }
        break;
      case "hyperlink":
        for (const child of node.children) visitInline(child);
        break;
      case "complexField":
        for (const run of [...node.fieldCode, ...node.fieldResult])
          visitInline(run);
        break;
      case "simpleField":
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
      case "inlineSdt":
      case "inlineWrapper":
        for (const child of node.content) visitInline(child);
        break;
      case "bookmarkStart":
      case "bookmarkEnd":
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "mathEquation":
      case "preservedInline":
        break;
      default: {
        const unhandled: never = node;
        throw new Error(`Unhandled inline ${JSON.stringify(unhandled)}`);
      }
    }
  };
  const visit = (blocks: readonly ParsedBlock[]): void => {
    for (const block of blocks) {
      switch (block.type) {
        case "paragraph": {
          for (const content of block.content) visitInline(content);
          visitParagraph(block);
          break;
        }
        case "table":
          for (const row of block.rows)
            for (const cell of row.cells) visit(cell.content);
          break;
        case "blockSdt":
        case "blockCustomXml":
          visit(block.content);
          break;
        case "preservedBlock":
        case "bookmarkStart":
        case "bookmarkEnd":
        case "moveFromRangeStart":
        case "moveFromRangeEnd":
        case "moveToRangeStart":
        case "moveToRangeEnd":
          break;
        default: {
          const unhandled: never = block;
          throw new Error(`Unhandled block ${JSON.stringify(unhandled)}`);
        }
      }
    }
  };
  visit(document.package.document.content);
  for (const parts of [document.package.headers, document.package.footers]) {
    for (const part of parts?.values() ?? []) visit(part.content);
  }
  for (const parts of [document.package.footnotes, document.package.endnotes]) {
    for (const part of parts ?? []) visit(part.content);
  }
};

/** Package definitions and direct provenance, never an effective reader sample. */
const numberingFactsFromDocument = (
  document: ParsedDocument,
): NumberingFacts => {
  const definitions = new Map(
    (document.package.styles?.styles ?? []).map((style) => [
      style.styleId,
      style,
    ]),
  );
  const styles = new Map<string, Numbering>();
  const resolve = (id: string, path = new Set<string>()): Numbering => {
    if (styles.has(id)) return styles.get(id);
    if (path.has(id)) throw new Error(`Cyclic style inheritance at ${id}`);
    path.add(id);
    const style = definitions.get(id);
    const value = numberingOver(
      style?.basedOn ? resolve(style.basedOn, path) : undefined,
      style?.pPr?.numPr,
    );
    styles.set(id, value);
    return value;
  };
  for (const id of definitions.keys()) resolve(id);
  const direct = new Map<string, Numbering>();
  visitParsedParagraphs(document, (paragraph) => {
    if (paragraph.paraId !== undefined)
      direct.set(paragraph.paraId, paragraph.formatting?.numPr);
  });
  const markerLevels = new Map<number, Set<number>>();
  const abstractNums = document.package.numbering?.abstractNums ?? [];
  for (const instance of document.package.numbering?.nums ?? []) {
    let abstract = abstractNums.find(
      (definition) => definition.abstractNumId === instance.abstractNumId,
    );
    // An empty definition linked to a numbering style takes its levels from the definition that style names.
    const link = abstract?.numStyleLink;
    if (link !== undefined && abstract?.levels.length === 0) {
      abstract =
        abstractNums.find(
          (definition) =>
            definition.styleLink === link && definition.levels.length > 0,
        ) ?? abstract;
    }
    const levels = new Map(
      (abstract?.levels ?? []).map((level) => [level.ilvl, level]),
    );
    for (const override of instance.levelOverrides ?? []) {
      if (override.lvl) levels.set(override.ilvl, override.lvl);
    }
    markerLevels.set(
      instance.numId,
      new Set(
        [...levels.values()]
          .filter((level) =>
            showsMarker(
              {
                text: level.lvlText,
                format: level.numFmt,
                legal: level.isLgl === true,
                hidden: Boolean(level.rPr?.hidden),
              },
              (ilvl) => levels.get(ilvl)?.numFmt,
            ),
          )
          .map((level) => level.ilvl),
      ),
    );
  }
  return { styles, direct, markerLevels };
};

export const numberingFactsOf = async (
  bytes: Uint8Array,
): Promise<NumberingFacts> => {
  const facts = numberingFactsFromDocument(
    await parseDocx(toArrayBuffer(bytes), { preloadFonts: false }),
  );
  // Unsaved positional identities come from the same owner as every reader anchor.
  const reviewer = await openReviewer(bytes);
  for (const { handle } of reviewer.listStories()) {
    const snapshot = reviewer.snapshotStory(handle);
    assert.ok(snapshot, "A discovered story must have a snapshot");
    const { anchors, blocks } = snapshot;
    for (const block of blocks) {
      if (block.kind === "diagnostic" || facts.direct.has(block.id)) continue;
      const anchor = anchors[block.id];
      assert.ok(anchor, `Block ${block.id} must have an anchor`);
      facts.direct.set(
        block.id,
        block.statedNumbering.kind === "inherit"
          ? undefined
          : block.statedNumbering,
      );
    }
  }
  return facts;
};

type TableLocation = {
  outerTableIndex: number;
  tableIndex: number;
  rowIndex: number;
  cellIndex: number;
  gridColumnIndex: number;
  columnSpan: number;
  rowSpan: number;
};

type Run = {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  directFormatting?: Record<string, unknown>;
};

/** One block as a reader sees it, with the fields the oracle compares. */
export type Row = {
  id: string;
  text: string;
  kind: string;
  styleId?: string;
  headingLevel?: number;
  /** The outline level the paragraph states itself, over its style's. */
  directOutlineLevel?:
    { kind: "heading"; level: number } | { kind: "bodyText" };
  /** The authored numbering override reported by the reader snapshot. */
  statedNumbering?: FolioContentStatedNumbering;
  /** Effective display level, derived from listReference by rowsOf. */
  listLevel?: number;
  listReference?: { numId: number; level: number };
  displayLabel?: string;
  directAlignment?: string;
  directSpacing?: Record<string, unknown>;
  table?: TableLocation;
  previewRuns?: readonly Run[];
};

/** The fields every block keeps unless an operation asked to change them. */
const KEPT_FIELDS = ["kind", "styleId", "headingLevel", "listLevel"] as const;

/** A field the request leaves to the document (a style's own outline or numbering). */
const ANY = Symbol("any");
type Fields = {
  [Key in (typeof KEPT_FIELDS)[number] | "directAlignment"]?:
    Row[Key] | typeof ANY;
};

type ListLineage =
  | {
      type: "reference";
      numId: number;
      source: "document" | "operation";
      basis: "direct" | "style";
      /** Stated as a level alone: the id still comes from the paragraph's style. */
      partial?: true;
    }
  | { type: "newList"; operationId: string }
  | {
      type: "none";
      source: "document" | "operation";
      basis: "direct" | "style";
      partial?: true;
    }
  | { type: "unknown"; source: "document" | "operation" };

/**
 * Direct numbering a copy of the paragraph keeps whole. A level stated alone
 * (`w:ilvl` without `w:numId`) takes its id from the style of the paragraph it
 * is on, so it follows the copy's style and is not carried over as written.
 */
const hasFixedDirectNumbering = (lineage: ListLineage): boolean =>
  hasDirectNumbering(lineage) &&
  !("partial" in lineage && lineage.partial === true);

const hasDirectNumbering = (lineage: ListLineage): boolean => {
  switch (lineage.type) {
    case "newList":
      return true;
    case "reference":
    case "none":
      return lineage.basis === "direct";
    case "unknown":
      return false;
    default: {
      const unhandled: never = lineage;
      throw new Error(`Unhandled list lineage ${JSON.stringify(unhandled)}`);
    }
  }
};

type Story = FolioDocumentStoryHandle;
const MAIN: Story = { type: "main" };
type NoteStory = Extract<Story, { type: "footnote" | "endnote" }>;
const noteStoryKey = ({ type, noteId }: NoteStory): string =>
  `${type}:${String(noteId)}`;

/**
 * Operations that leave every note reference in their target block where it
 * is: formatting, comments and insertions beside the block. Any other edit of
 * a block holding a reference leaves that note's comments unpredicted.
 */
export const KEEPS_NOTE_REFERENCES: ReadonlySet<string> = new Set([
  "formatRange",
  "setBlockParagraphProperties",
  "commentOnRange",
  "commentOnBlock",
  "insertAfterBlock",
  "insertBeforeBlock",
]);

/**
 * Update `references` (the note references each live block holds) for one
 * operation that removed `removal`, and return the note stories whose last
 * reference it removed. A merge moves the next block's content, and the note
 * references in it, into the surviving block, so those notes stay.
 */
export const retireNoteReferences = (
  references: Map<string, readonly Story[]>,
  operationType: string,
  blockId: string,
  removal: readonly string[],
): string[] => {
  const candidates = new Set<string>();
  for (const removedBlockId of removal) {
    const held = references.get(removedBlockId) ?? [];
    references.delete(removedBlockId);
    if (operationType === "mergeBlockWithNext" && removedBlockId !== blockId) {
      references.set(blockId, [...(references.get(blockId) ?? []), ...held]);
      continue;
    }
    for (const story of held) {
      if (story.type === "footnote" || story.type === "endnote")
        candidates.add(noteStoryKey(story));
    }
  }
  // A note stays while any live block still references it.
  for (const held of references.values()) {
    for (const story of held) {
      if (story.type === "footnote" || story.type === "endnote") {
        candidates.delete(noteStoryKey(story));
      }
    }
  }
  return [...candidates];
};

/** The blocks of `story` (the body by default) a reader lists. */
export const rowsOf = (reviewer: Reviewer, story: Story = MAIN): Row[] => {
  const rows: Row[] = [];
  for (const block of blocksOfStory(reviewer, story)) {
    switch (block.kind) {
      case "paragraph":
      case "heading":
      case "listItem": {
        const row: Row = { id: block.id, kind: block.kind, text: block.text };
        if (block.styleId !== undefined) row.styleId = block.styleId;
        if (block.headingLevel !== undefined)
          row.headingLevel = block.headingLevel;
        if (block.directOutlineLevel !== undefined)
          row.directOutlineLevel = block.directOutlineLevel;
        row.statedNumbering = block.statedNumbering;
        if (block.directAlignment !== undefined)
          row.directAlignment = block.directAlignment;
        if (block.directSpacing !== undefined)
          row.directSpacing = { ...block.directSpacing };
        if (block.listReference !== undefined) {
          row.listReference = block.listReference;
          row.listLevel = block.listReference.level;
        }
        if (block.displayLabel !== undefined)
          row.displayLabel = block.displayLabel;
        if (block.table !== undefined) row.table = block.table;
        if (block.previewRuns !== undefined) {
          row.previewRuns = block.previewRuns.map((run) => {
            const preview = {
              text: run.text,
              bold: run.bold,
              italic: run.italic,
              underline: run.underline,
            };
            if (run.directFormatting === undefined) return preview;
            return {
              ...preview,
              directFormatting: { ...run.directFormatting },
            };
          });
        }
        rows.push(row);
        break;
      }
      case "diagnostic":
        rows.push({ id: block.id, kind: block.kind, text: block.text });
        break;
      default: {
        const unreachable: never = block;
        throw new Error(`Unhandled content block: ${unreachable}`);
      }
    }
  }
  return rows;
};

const save = async (reviewer: Reviewer): Promise<Uint8Array> =>
  new Uint8Array(await reviewer.toBuffer());

/**
 * Whether a story is still in a document. A note goes with its reference:
 * accepting the deletion of the text that holds the reference removes the
 * note, whatever it holds by then.
 */
type StoryPresence = "present" | "removed";

type ResolvedState = {
  rows: Row[];
  story: StoryPresence;
  bytes: Uint8Array;
  comments: Comment[];
  links: LinkSnapshot;
};

/**
 * `bytes` opened, every change resolved one way, saved and reopened: the
 * blocks a reader is left with, whether the story is still there, and the
 * saved package.
 */
export const resolvedState = async (
  bytes: Uint8Array,
  resolution: "accept" | "reject",
  story: Story = MAIN,
): Promise<ResolvedState> => {
  const reviewer = await openReviewer(bytes);
  if (resolution === "accept") reviewer.acceptAll();
  else reviewer.rejectAll();
  const saved = await save(reviewer);
  const reopened = await openReviewer(saved);
  return {
    rows: rowsOf(reopened, story),
    story: reopened.readStory(story) === null ? "removed" : "present",
    bytes: saved,
    comments: commentsOf(reopened),
    links: captureLinks(reopened),
  };
};

type Comment = {
  id: number;
  text: string;
  anchor: string;
  blockId: string | null;
};
const commentsOf = (reviewer: Reviewer): Comment[] =>
  reviewer.getComments().map((comment) => ({
    id: comment.id,
    text: comment.text,
    anchor: comment.anchoredText ?? "",
    blockId: comment.blockId ?? null,
  }));

// ---------------------------------------------------------------------------
// The model
// ---------------------------------------------------------------------------

type TextEdit = { start: number; end: number; replace: string };

/** A block of the pre-state, or one an operation adds, with what the batch did to it. */
type ModelRow = {
  /** The pre-state block; absent for a block an operation adds. */
  pre?: Row;
  text: string;
  fields: Fields;
  listLineage: ListLineage;
  /** Extra checks on the result's block: a message when it fails. */
  checks: ((row: Row) => string | null)[];
  /** Text edits in the pre-state text's coordinates. */
  edits: TextEdit[];
  /** Split points in the pre-state text's coordinates, with what the break consumes. */
  splits: { offset: number; consumed: number; second?: Fields }[];
  mergeSeparator?: string;
  removed: boolean;
  /**
   * The block ends in a paragraph mark pending deletion: a reader lists it
   * and the block after it apart, accepting joins them. What an operation
   * on either does after the join is not modelled yet.
   */
  pendingJoin?: boolean;
  /**
   * A whole paragraph pending deletion, which a reader lists as a blank
   * block: accepting removes it, mark and all, unless an operation wrote
   * into it.
   */
  pendingDeletion?: boolean;
  /**
   * A block the reviewer still lists but the resolved pre-state does not
   * hold: nothing read from that pre-state describes it.
   */
  outsidePreState?: boolean;
  /**
   * A block inserted before or after the table its anchor sits in. Rows and
   * cells the batch adds to that table go inside it, so they stay between
   * the table's own blocks and this one, whatever the batch order.
   */
  besideTable?: boolean;
  /** A block of a table row the batch deletes, rather than one emptied alone. */
  rowDeleted?: boolean;
  /** A text box paragraph removed with the paragraph the box is drawn in. */
  removedWithTextBox?: boolean;
  /** For a block inserted next to a text box paragraph: that paragraph, whose box holds it. */
  inTextBoxOf?: ModelRow;
  before: ModelRow[];
  after: ModelRow[];
  /** Formatting that must hold over `[start, end)` of the result's text. */
  formats: {
    start: number;
    end: number;
    property: "bold" | "italic" | "underline";
  }[];
};

export type Model = {
  rows: ModelRow[];
  /** Live source paragraphs whose pending marks run into an accepted row. */
  pendingJoins: Map<ModelRow, { sources: readonly Row[]; destination: Row }>;
  /** Outside text carried to a new destination by an earlier table removal. */
  preservedJoinPrefixes: Map<ModelRow, TextEdit>;
  mode: Mode;
  tables: TableModel;
  tableGaps: string[];
  /** Existing paragraphs reveal style-derived kind, levels and effective run formatting. */
  styleExamples: Map<string, Row>;
  numberingFacts?: NumberingFacts;
  /** Text of the live blocks an operation can name, before the batch. */
  liveTextById: ReadonlyMap<string, string>;
  /** Live blocks explicitly deleted by applied operations, including pending joins. */
  deletedBlockIds: Set<string>;
  /**
   * The live document's numbering, for blocks the accepted package no longer
   * holds (a paragraph pending deletion): what such a block states is what a
   * copy of it inherits, and a list only it used is defined there alone.
   */
  liveNumberingFacts?: NumberingFacts;
  /** Note stories whose references were removed by applied structural operations. */
  removedNoteStories: Set<string>;
  /**
   * Note stories whose references an operation other than a removal may have
   * moved or deleted; their comments are not predicted either way.
   */
  uncertainNoteStories: Set<string>;
  /** Note story containing each comment-addressable block in the pre-state. */
  noteStoryByBlockId: ReadonlyMap<string, string>;
  /** Paragraph style ids the saved package must define. */
  styles: Set<string>;
  /** Comments that must exist, with their anchored text when it is known. */
  comments: {
    text: string;
    anchor?: () => string | undefined;
    removed?: () => boolean;
  }[];
  /** Operations the oracle has no expectation for; the check notes them. */
  unmodelled: string[];
  /** Blocks that are paragraphs of a text box drawn in the block before them. */
  inTextBox?: ReadonlySet<string>;
  /**
   * The story in the resolved pre-state. A removed one (a note whose
   * reference's deletion is pending) stays removed once accepted: nothing
   * an operation puts in it survives.
   */
  story: StoryPresence;
};

const fieldsOf = (row: Row): Fields => {
  const fields: Fields = {};
  for (const key of KEPT_FIELDS)
    (fields as Record<string, unknown>)[key] = row[key];
  if (row.directAlignment !== undefined)
    fields.directAlignment = row.directAlignment;
  return fields;
};

const listLineageOf = (
  numbering: Numbering | typeof UNKNOWN_NUMBERING,
  source: "document" | "operation",
  basis: "direct" | "style",
): ListLineage => {
  if (numbering === UNKNOWN_NUMBERING) return { type: "unknown", source };
  if (numbering?.kind === "reference")
    return { type: "reference", numId: numbering.numId, source, basis };
  if (numbering?.kind === "none") return { type: "none", source, basis };
  return { type: "unknown", source };
};

type StyleListLineageOptions = {
  model: Model;
  styleId: string | null;
  source: ModelRow;
  /** Whether the paragraph keeps its direct numbering; `inherit` drops it. */
  keepsDirect?: boolean;
};

const styleListLineage = ({
  model,
  styleId,
  source,
  keepsDirect = true,
}: StyleListLineageOptions): ListLineage => {
  const numbering = requestedStyleNumbering(
    model,
    styleId,
    keepsDirect ? (liveParagraph(source) ?? source.pre) : undefined,
  );
  if (model.numberingFacts !== undefined && numbering !== UNKNOWN_NUMBERING) {
    if (
      numbering === undefined ||
      (numbering.kind === "reference" &&
        numberingKind(model, numbering) !== "listItem")
    ) {
      return { type: "none", source: "operation", basis: "style" };
    }
  }
  return listLineageOf(numbering, "operation", "style");
};

/**
 * Untouched pre-state rows need style facts to distinguish inherited list
 * markers from direct numbering. Operation-written lineage already describes
 * the current paragraph and must not be recomputed from its pre-state.
 */
const inheritedListFields = (model: Model, row: Row): Fields => {
  const fields = fieldsOf(row);
  const current = model.rows.find((candidate) => candidate.pre?.id === row.id);
  if (current) {
    switch (current.listLineage.type) {
      case "newList":
        return fields;
      case "reference":
      case "none":
      case "unknown":
        if (current.listLineage.source === "operation") {
          // An operation that left the block numbering nothing leaves a
          // copy of it a plain paragraph, whatever marker the pre-state had.
          return current.listLineage.type === "none" &&
            fields.kind === "listItem"
            ? { ...fields, kind: "paragraph", listLevel: undefined }
            : fields;
        }
        break;
      default: {
        const unhandled: never = current.listLineage;
        throw new Error(`Unhandled list lineage ${JSON.stringify(unhandled)}`);
      }
    }
  }
  if (!model.numberingFacts) return fields;
  if (row.kind !== "listItem" && row.kind !== "heading") return fields;
  const numbering = requestedStyleNumbering(model, row.styleId ?? null, row);
  if (numbering === UNKNOWN_NUMBERING) return fields;
  const listItem = numberingKind(model, numbering) === "listItem";
  // A copy takes the numbering a heading states, not the marker it shows: a
  // heading whose mark is pending deletion paints its original numbering,
  // which accepting retires with it.
  if (row.kind === "heading") {
    return {
      ...fields,
      listLevel:
        listItem && numbering?.kind === "reference"
          ? (numbering.ilvl ?? 0)
          : undefined,
    };
  }
  if (listItem) return fields;
  return { ...fields, kind: "paragraph", listLevel: undefined };
};

const modelRow = (text: string, fields: Fields = {}, pre?: Row): ModelRow => {
  let listLineage: ListLineage = { type: "unknown", source: "document" };
  if (pre?.listReference) {
    listLineage = {
      type: "reference",
      numId: pre.listReference.numId,
      source: "document",
      basis: pre.statedNumbering?.kind === "inherit" ? "style" : "direct",
      ...(pre.statedNumbering?.kind === "levelOnly"
        ? { partial: true as const }
        : {}),
    };
  } else if (pre?.statedNumbering?.kind === "none") {
    listLineage = { type: "none", source: "document", basis: "direct" };
  } else if (pre?.statedNumbering?.kind === "inherit") {
    listLineage = { type: "none", source: "document", basis: "style" };
  }
  return {
    ...(pre ? { pre } : {}),
    text,
    fields,
    listLineage,
    checks: [],
    edits: [],
    splits: [],
    removed: false,
    before: [],
    after: [],
    formats: [],
  };
};

/**
 * Which live blocks with words that accepting removes run on into the next
 * block. `marked`: the paragraphs whose own mark is a pending deletion or a
 * move's source, read from the saved live package; any other such block
 * drops words the model does not place. `assumed`: every one of them (for
 * hand-built rows with no package to read).
 */
export type JoinSources =
  { type: "marked"; ids: ReadonlySet<string> } | { type: "assumed" };

const ASSUMED_JOIN_SOURCES: JoinSources = { type: "assumed" };

const joinsOnAccept = (sources: JoinSources, id: string): boolean => {
  switch (sources.type) {
    case "marked":
      return sources.ids.has(id);
    case "assumed":
      return true;
    default: {
      const unhandled: never = sources;
      throw new Error(`Unhandled join sources ${JSON.stringify(unhandled)}`);
    }
  }
};

/** Paragraph marks accepting removes: a pending deletion, or a move's source. */
const JOINING_MARK_KINDS: ReadonlySet<string> = new Set(["del", "moveFrom"]);

/** The paraIds of the paragraphs of `bytes` whose mark accepting removes. */
const markedJoinSourcesOf = async (bytes: Uint8Array): Promise<JoinSources> => {
  const ids = new Set<string>();
  visitParsedParagraphs(
    await parseDocx(toArrayBuffer(bytes), { preloadFonts: false }),
    (paragraph) => {
      if (
        paragraph.paraId !== undefined &&
        paragraph.pPrMark !== undefined &&
        JOINING_MARK_KINDS.has(paragraph.pPrMark.kind)
      )
        ids.add(paragraph.paraId);
    },
  );
  return { type: "marked", ids };
};

/**
 * The model of `rows`. `live`, the reviewer's own blocks, adds the blocks a
 * reader still lists but accepting removes (a whole paragraph pending
 * deletion reads as a blank block): an operation may name one, and what it
 * asks for is then placed where the reader saw that block.
 */
export const modelOf = (
  rows: readonly Row[],
  live: readonly Row[] = rows,
  joinSources: JoinSources = ASSUMED_JOIN_SOURCES,
): Model => {
  const model: Model = {
    rows: rows.map((row) => modelRow(row.text, fieldsOf(row), row)),
    pendingJoins: new Map(),
    preservedJoinPrefixes: new Map(),
    mode: "direct",
    tables: modelFromRows(rows),
    tableGaps: [],
    styleExamples: new Map(
      rows.flatMap((row) =>
        row.styleId !== undefined && row.text.length > 0
          ? [[row.styleId, row] as const]
          : [],
      ),
    ),
    liveTextById: new Map(live.map(({ id, text }) => [id, text])),
    deletedBlockIds: new Set(),
    removedNoteStories: new Set(),
    uncertainNoteStories: new Set(),
    noteStoryByBlockId: new Map(),
    styles: new Set(),
    comments: [],
    unmodelled: [],
    story: "present",
  };
  const known = new Set(rows.map((row) => row.id));
  const indexOf = (id: string | undefined): number =>
    id === undefined
      ? -1
      : model.rows.findIndex((candidate) => candidate.pre?.id === id);
  let previous: string | undefined;
  // Blocks whose mark accepting removes, waiting for the block they run into.
  let joining = false;
  let joiningSources: Row[] = [];
  for (const row of live) {
    if (known.has(row.id)) {
      // The paragraph left by a join is the one whose mark stays: the block
      // the joined ones run into, which keeps its identity.
      const survivor = model.rows[indexOf(row.id)];
      if (survivor?.pre) {
        // Text and block properties in `rows` are the accepted projection,
        // but operations target the live snapshot. A pending table-row
        // deletion can renumber the accepted projection's surviving rows;
        // keep the live table coordinates for source-anchor placement.
        survivor.pre = {
          ...survivor.pre,
          ...(row.table !== undefined && { table: row.table }),
        };
        if (joining) {
          survivor.pendingJoin = true;
          model.pendingJoins.set(survivor, {
            sources: joiningSources,
            destination: row,
          });
        }
      }
      joining = false;
      joiningSources = [];
      previous = row.id;
      continue;
    }
    known.add(row.id);
    const ghost = modelRow("", {}, row);
    ghost.outsidePreState = true;
    if (row.text.length > 0 && !joinsOnAccept(joinSources, row.id)) {
      // Accepting drops words whose paragraph mark stays (a row or cell
      // pending deletion whose runs are not marked deleted, say). Where they
      // go is not modelled; no join is guessed for them.
      ghost.removed = true;
      model.unmodelled.push(
        `block ${row.id} has words accepting drops with no pending paragraph-mark deletion`,
      );
      model.rows.splice(indexOf(previous) + 1, 0, ghost);
      previous = row.id;
      continue;
    }
    if (row.text.length > 0) {
      // Runs on into the block after it once accepted.
      joining = true;
      joiningSources.push(row);
      ghost.pendingJoin = true;
      ghost.removed = true;
      model.rows.push(ghost);
      continue;
    }
    ghost.pendingDeletion = true;
    // Words joining a later block run on through it: what lands beside it
    // decides where they stop.
    if (joining) ghost.pendingJoin = true;
    model.rows.splice(indexOf(previous) + 1, 0, ghost);
    previous = row.id;
  }
  if (joining) {
    // Nothing after it stays: it can only have joined the block before.
    const head = model.rows[indexOf(previous)];
    if (head) head.pendingJoin = true;
  }
  return model;
};

/** Why an operation could not be modelled; the check reports it as an oracle gap. */
class Unmodelled extends Error {}

const target = (
  model: Model,
  blockId: unknown,
  { adjacent = false } = {},
): ModelRow => {
  const row = model.rows.find((candidate) => candidate.pre?.id === blockId);
  if (!row)
    throw new Unmodelled(
      `block ${String(blockId)} is not in the resolved pre-state`,
    );
  if (row.pendingJoin)
    throw new Unmodelled(`block ${String(blockId)} has a pending join`);
  // An insertion only places a block beside its anchor, which may go.
  if (row.removed && !adjacent) {
    throw new Error(
      `applied an operation on block ${String(blockId)}, which an earlier one of the batch removed`,
    );
  }
  return row;
};

const sameParagraphContainer = (left: Row, right: Row): boolean => {
  if (left.table === undefined || right.table === undefined) {
    return left.table === undefined && right.table === undefined;
  }
  return (
    left.table.tableIndex === right.table.tableIndex &&
    left.table.rowIndex === right.table.rowIndex &&
    left.table.cellIndex === right.table.cellIndex
  );
};

/** Whether block `id` is in `rows` and the last of its story or table cell there. */
const endsContainer = (rows: readonly Row[], id: string): boolean => {
  const index = rows.findIndex((row) => row.id === id);
  const row = rows[index];
  return (
    row !== undefined &&
    rows.slice(index + 1).every((later) => !sameParagraphContainer(row, later))
  );
};

/**
 * Whether a deletion of block `id` kept it. A story or table cell ends with a
 * paragraph, so deleting the last block of one (`before`, the rows before the
 * batch) may only empty it; it is kept when it is still in `after`, empty and
 * still the last of its story or cell. A paragraph property edit replayed
 * after such a deletion formats that block; after any other, the block is gone.
 */
export const deletionKeptBlock = ({
  before,
  after,
  id,
}: {
  before: readonly Row[];
  after: readonly Row[];
  id: string;
}): boolean =>
  endsContainer(before, id) &&
  after.find((row) => row.id === id)?.text === "" &&
  endsContainer(after, id);

/** How the kept block `row` differs from what a paragraph property edit asked of it. */
const keptBlockProblems = (
  requested: ParagraphPropertiesRequest,
  row: Row,
): string[] => {
  const problems: string[] = [];
  const name = `the emptied block ${row.id}`;
  for (const [key, value] of Object.entries(requested.fields)) {
    if (value === ANY) continue;
    const actualValue = row[key as keyof Row];
    if (JSON.stringify(actualValue) !== JSON.stringify(value)) {
      problems.push(
        `${name}: ${key} is ${JSON.stringify(actualValue)}, expected ${JSON.stringify(value)}\n    after ${JSON.stringify(row)}`,
      );
    }
  }
  for (const check of requested.checks) {
    const problem = check(row);
    if (problem !== null) problems.push(`${name}: ${problem}`);
  }
  const lineage = requested.listLineage;
  switch (lineage?.type) {
    case "reference":
      if (row.listReference?.numId !== lineage.numId) {
        problems.push(
          `${name}: numbering list reference ${JSON.stringify(row.listReference)} differs from modeled reference ${lineage.numId}`,
        );
      }
      break;
    case "none":
      if (row.listReference !== undefined) {
        problems.push(
          `${name}: unexpected numbering list reference ${JSON.stringify(row.listReference)}`,
        );
      }
      break;
    case "newList":
      if (row.listReference === undefined)
        problems.push(`${name}: new list has no list reference`);
      break;
    case "unknown":
    case undefined:
      break;
    default: {
      const unhandled: never = lineage;
      throw new Error(`Unhandled list lineage ${JSON.stringify(unhandled)}`);
    }
  }
  return problems;
};

const isContainerCarrier = (model: Model, row: ModelRow): boolean => {
  const pre = row.pre;
  if (pre === undefined) return false;
  const index = model.rows.indexOf(row);
  return model.rows
    .slice(index + 1)
    .every(
      (candidate) =>
        candidate.pre === undefined ||
        !sameParagraphContainer(pre, candidate.pre),
    );
};

const tableOf = (row: ModelRow): TableLocation => {
  const table = row.pre?.table;
  if (!table) throw new Unmodelled("the anchor is not in a table");
  return table;
};

const tableRows = (model: Model, tableIndex: number): ModelRow[] =>
  model.rows.filter((row) => row.pre?.table?.tableIndex === tableIndex);

/** Table removal owns live cells, including comments, rather than joined outside text. */
const removeTableRows = (
  model: Model,
  removedRows: readonly ModelRow[],
): void => {
  const newlyRemoved = removedRows.filter((row) => !row.removed);
  for (const row of removedRows) {
    row.removed = true;
    if (row.pre) model.deletedBlockIds.add(row.pre.id);
  }
  for (const survivor of newlyRemoved) {
    let prefix = model.preservedJoinPrefixes.get(survivor)?.replace ?? "";
    const ownership = model.pendingJoins.get(survivor);
    // Ownership transfers once; later region removals can include this row again.
    model.preservedJoinPrefixes.delete(survivor);
    model.pendingJoins.delete(survivor);
    if (ownership) {
      const { sources, destination } = ownership;
      const joinedText =
        sources.map(({ text }) => text).join("") + destination.text;
      if (survivor.text !== joinedText) {
        throw new Error(
          `Pending join ownership does not reconstruct block ${destination.id}`,
        );
      }
      prefix += sources
        .filter((source) => !model.deletedBlockIds.has(source.id))
        .map(({ text }) => text)
        .join("");
    }
    if (prefix.length === 0) continue;
    const next = model.rows
      .slice(model.rows.indexOf(survivor) + 1)
      .find((row) => !row.removed);
    if (next) {
      const existing = model.preservedJoinPrefixes.get(next);
      if (existing) existing.replace += prefix;
      else {
        const edit = { start: 0, end: 0, replace: prefix };
        next.edits.push(edit);
        model.preservedJoinPrefixes.set(next, edit);
      }
      continue;
    }
    const first = ownership?.sources.find(
      (source) => !model.deletedBlockIds.has(source.id),
    );
    model.rows.push(
      modelRow(prefix, first ? fieldsOf(first) : survivor.fields),
    );
  }
};

/** Removing a table region may target the destination of a pending outside join. */
const tableRemovalTarget = (model: Model, blockId: unknown): ModelRow => {
  const row = model.rows.find((candidate) => candidate.pre?.id === blockId);
  if (!row)
    throw new Error(`Table removal anchor ${String(blockId)} is absent`);
  if (row.removed)
    throw new Error(
      `Table removal anchor ${String(blockId)} was already removed`,
    );
  tableOf(row);
  return row;
};

const hasMergedCells = (model: Model, tableIndex: number): boolean =>
  tableRows(model, tableIndex).some(
    (row) =>
      (row.pre?.table?.rowSpan ?? 1) > 1 ||
      (row.pre?.table?.columnSpan ?? 1) > 1,
  );

const paragraphsOf = (text: string): string[] =>
  text.split(/\r\n|\r|\n/u).filter((line) => line.length > 0);

/** The paragraph's own numbering is not in the numbering source: nothing is expected of it. */
const UNKNOWN_NUMBERING = Symbol("unknown numbering");
type RequestedNumbering = Numbering | typeof UNKNOWN_NUMBERING;

/**
 * Whether the numbering source may lack the paragraph. The saved modes read
 * the accepted package, which no longer holds a paragraph that accepting
 * removes. Suggested mode captures every live paragraph, including proposals.
 */
const outsideNumberingSource = (model: Model, id: string): boolean =>
  model.rows.some((row) => row.outsidePreState === true && row.pre?.id === id);

const requestedStyleNumbering = (
  model: Model,
  styleId: string | null,
  row?: Row,
): RequestedNumbering => {
  const facts = model.numberingFacts;
  const style = styleId === null ? undefined : facts?.styles.get(styleId);
  if (!row) return style;
  let direct: Numbering;
  if (facts) {
    if (!facts.direct.has(row.id)) {
      // Any other paragraph must be there.
      assert.ok(
        outsideNumberingSource(model, row.id),
        `Numbering provenance missing for ${row.id}`,
      );
      const live = model.liveNumberingFacts;
      if (live === undefined || !live.direct.has(row.id))
        return UNKNOWN_NUMBERING;
      direct = live.direct.get(row.id);
    } else {
      direct = facts.direct.get(row.id);
    }
  } else if (row.statedNumbering?.kind === "reference") {
    direct = row.statedNumbering;
  } else if (row.statedNumbering?.kind === "none") {
    direct = row.statedNumbering;
  }
  return numberingOver(style, direct);
};

/**
 * The lineage of a document row as it stands now. A list marker the row shows
 * only through a pending property change numbers nothing a copy inherits, so
 * the authored numbering decides, as for the fields a copy takes.
 */
const currentLineage = (model: Model, row: ModelRow): ListLineage => {
  const { listLineage } = row;
  if (
    listLineage.type !== "reference" ||
    listLineage.source !== "document" ||
    !model.numberingFacts ||
    !row.pre
  ) {
    return listLineage;
  }
  const numbering = requestedStyleNumbering(
    model,
    row.pre.styleId ?? null,
    row.pre,
  );
  if (numbering === UNKNOWN_NUMBERING) return listLineage;
  if (
    numbering?.kind === "reference" &&
    numberingKind(model, numbering) === "listItem"
  ) {
    return { ...listLineage, numId: numbering.numId };
  }
  return { type: "none", source: "document", basis: listLineage.basis };
};

type OperationListLineageOptions = {
  model: Model;
  request: Record<string, unknown>;
  operationId: string;
  source: ModelRow;
  inheritsFormatting: boolean;
  /** The request sets the properties of an existing paragraph, not of a copy. */
  replacesDirect?: boolean;
};

const statedNumId = (numbering: object): number => {
  assert.ok(
    "numId" in numbering && typeof numbering["numId"] === "number",
    "An applied reference request must name a numeric list id",
  );
  return numbering["numId"];
};

const statedLevel = (numbering: object): number | undefined =>
  "ilvl" in numbering && typeof numbering.ilvl === "number"
    ? numbering.ilvl
    : undefined;

/** A reference whose level shows no marker numbers nothing, however it is stated. */
const renderedLineage = (model: Model, numbering: Numbering): ListLineage => {
  if (numbering?.kind !== "reference") {
    return { type: "none", source: "operation", basis: "direct" };
  }
  if (!model.numberingFacts) return { type: "unknown", source: "operation" };
  return numberingKind(model, numbering) === "listItem"
    ? {
        type: "reference",
        numId: numbering.numId,
        source: "operation",
        basis: "direct",
      }
    : { type: "none", source: "operation", basis: "direct" };
};

type BaseNumberingOptions = {
  model: Model;
  request: Record<string, unknown>;
  source: ModelRow;
  inheritsFormatting: boolean;
  replacesDirect: boolean;
};

/** The numbering a request is stated over: the style's, plus the direct numbering the block keeps. */
const baseNumbering = ({
  model,
  request,
  source,
  inheritsFormatting,
  replacesDirect,
}: BaseNumberingOptions): RequestedNumbering => {
  const requested = request["styleId"];
  let styleId: string | null;
  if (typeof requested === "string" || requested === null) {
    styleId = requested;
  } else if (typeof source.fields.styleId === "string") {
    styleId = source.fields.styleId;
  } else {
    styleId = source.pre?.styleId ?? null;
  }
  // A stated numbering replaces the paragraph's own `w:numPr`, so only the
  // style's numbering is left for a partial override to keep.
  if (!inheritsFormatting || replacesDirect) {
    return requestedStyleNumbering(model, styleId);
  }
  const lineage = source.listLineage;
  if (lineage.type === "newList") return UNKNOWN_NUMBERING;
  if (lineage.source === "operation") {
    return lineage.type === "none" ? { kind: "none" } : UNKNOWN_NUMBERING;
  }
  return requestedStyleNumbering(
    model,
    styleId,
    liveParagraph(source) ?? source.pre,
  );
};

const operationListLineage = ({
  model,
  request,
  operationId,
  source,
  inheritsFormatting,
  replacesDirect = false,
}: OperationListLineageOptions): ListLineage => {
  const numbering = request["numbering"];
  if ("numbering" in request) {
    assert.ok(
      numbering !== null && typeof numbering === "object",
      "An applied numbering request must be an object",
    );
    assert.ok(
      "kind" in numbering,
      "An applied numbering request must name a kind",
    );
    switch (numbering["kind"]) {
      case "newList":
        return { type: "newList", operationId };
      case "none":
        return { type: "none", source: "operation", basis: "direct" };
      case "reference":
      case "levelOnly": {
        const level = statedLevel(numbering);
        let stated: Numbering;
        if (numbering["kind"] === "reference") {
          stated = {
            kind: "reference",
            numId: statedNumId(numbering),
            ...(level === undefined ? {} : { ilvl: level }),
          };
        } else {
          assert.ok(level !== undefined, "A level request must name a level");
          stated = { kind: "levelOnly", ilvl: level };
        }
        const base = baseNumbering({
          model,
          request,
          source,
          inheritsFormatting,
          replacesDirect,
        });
        const partially = (lineage: ListLineage): ListLineage =>
          stated?.kind === "levelOnly" &&
          (lineage.type === "reference" || lineage.type === "none")
            ? { ...lineage, partial: true }
            : lineage;
        if (base === UNKNOWN_NUMBERING) {
          return stated?.kind === "reference" && stated.ilvl !== undefined
            ? renderedLineage(model, stated)
            : { type: "unknown", source: "operation" };
        }
        return partially(renderedLineage(model, numberingOver(base, stated)));
      }
      case "inherit": {
        const styleId = request["styleId"];
        if (typeof styleId === "string" || styleId === null) {
          return styleListLineage({
            model,
            styleId,
            source,
            keepsDirect: false,
          });
        }
        return styleListLineage({
          model,
          styleId:
            typeof source.fields.styleId === "string"
              ? source.fields.styleId
              : (source.pre?.styleId ?? null),
          source,
          keepsDirect: false,
        });
      }
      default:
        assert.fail(
          `Unexpected applied numbering kind ${JSON.stringify(numbering["kind"])}`,
        );
    }
  }
  if ("styleId" in request) {
    const styleId = request["styleId"];
    assert.ok(
      typeof styleId === "string" || styleId === null,
      "An applied style id must be a string or null",
    );
    if (inheritsFormatting && hasFixedDirectNumbering(source.listLineage)) {
      return source.listLineage;
    }
    return styleListLineage({
      model,
      styleId,
      source,
      keepsDirect: inheritsFormatting,
    });
  }
  if (inheritsFormatting) return currentLineage(model, source);
  return { type: "none", source: "operation", basis: "direct" };
};

const numberingFields = (
  model: Model,
  styleId: string | null,
  row?: Row,
): Fields => {
  const numbering = requestedStyleNumbering(model, styleId, row);
  if (numbering === UNKNOWN_NUMBERING) return { listLevel: ANY };
  return { listLevel: shownReference(model, numbering)?.level };
};

/** A retained reference alone is not a rendered marker: its level must exist and be visible. */
const numberingKind = (model: Model, numbering: Numbering): Fields["kind"] => {
  if (numbering?.kind !== "reference") return "paragraph";
  assert.ok(
    model.numberingFacts,
    "Numbering definitions are required to classify a reference",
  );
  // A list only a block pending deletion used is gone from the accepted
  // package; the live document still defines it, and a copy stating it makes
  // the save define it again.
  const levels =
    model.numberingFacts.markerLevels.get(numbering.numId) ??
    model.liveNumberingFacts?.markerLevels.get(numbering.numId);
  return levels?.has(numbering.ilvl ?? 0) ? "listItem" : "paragraph";
};

/**
 * The list reference a reader shows for `numbering`: one only where its level
 * renders a marker. Every check of a row's `listReference` derives it here.
 */
const shownReference = (
  model: Model,
  numbering: Numbering,
): { numId: number; level: number } | undefined =>
  numbering?.kind === "reference" &&
  numberingKind(model, numbering) === "listItem"
    ? { numId: numbering.numId, level: numbering.ilvl ?? 0 }
    : undefined;

/**
 * The fields a style id asks for on `pre` (absent: a new block, formatted
 * like `inherited`, the block it was inserted beside). A style may number its
 * paragraphs, and a direct outline level outlives a restyle, so the kind and
 * levels are the request's only for a plain paragraph made a `HeadingN`. A
 * new block keeps the outline level its anchor states, and only an example
 * that states none of its own shows its style's.
 */
const styleFields = (
  model: Model,
  styleId: string | null,
  pre: Row | undefined,
  inherited?: Row,
): Fields => {
  if (styleId !== null) model.styles.add(styleId);
  const example =
    styleId === null ? undefined : model.styleExamples.get(styleId);
  const styleExample =
    example?.directOutlineLevel === undefined ? example : undefined;
  const heading =
    styleId === null
      ? undefined
      : /^Heading(?<level>[1-9])$/u.exec(styleId)?.groups?.["level"];
  const plain =
    pre !== undefined &&
    pre.kind !== "heading" &&
    pre.directOutlineLevel === undefined;
  const stated =
    pre === undefined && styleId !== null
      ? inherited?.directOutlineLevel
      : undefined;
  let headingLevel: Fields["headingLevel"] = ANY;
  if (stated)
    headingLevel = stated.kind === "heading" ? stated.level + 1 : undefined;
  else if (styleExample && pre?.headingLevel === undefined) {
    headingLevel = styleExample.headingLevel;
  } else if (heading && plain) headingLevel = Number(heading);
  const numberedFields = numberingFields(model, styleId, pre ?? inherited);
  // A sample's classification may come from different direct numbering or marker visibility.
  let exampleKind = styleExample?.kind;
  const numbering = requestedStyleNumbering(model, styleId, pre ?? inherited);
  if (
    exampleKind !== undefined &&
    exampleKind !== "heading" &&
    numbering !== UNKNOWN_NUMBERING
  ) {
    const kind = numberingKind(model, numbering);
    if (kind !== ANY) exampleKind = kind;
  }
  // Whether the block is a list item then depends on numbering that is not known.
  const undecided =
    exampleKind === "listItem" && numberedFields.listLevel === ANY;
  return {
    styleId: styleId ?? undefined,
    kind:
      stated || undecided ? ANY : (exampleKind ?? (heading ? "heading" : ANY)),
    headingLevel,
    ...numberedFields,
  };
};

const isBullet = (label: string | undefined): boolean =>
  label !== undefined && label.length > 0 && !/[\p{L}\p{N}]/u.test(label);

type ParagraphRequest = Record<string, unknown>;

/** Checks that the numbering a restyle derives survives; direct numbering stated after it replaces them. */
const STYLE_NUMBERING_CHECKS = new WeakSet<ModelRow["checks"][number]>();

/** Drop the style-derived numbering checks once a request states its own numbering. */
const retireStyleNumberingChecks = (
  row: ModelRow,
  request: ParagraphRequest,
): void => {
  const stated = request["numbering"];
  if (!isRecord(stated) || stated["kind"] === "inherit") return;
  row.checks = row.checks.filter((check) => !STYLE_NUMBERING_CHECKS.has(check));
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

type ParagraphRequestOptions = {
  model: Model;
  request: ParagraphRequest;
  pre: Row | undefined;
  source: ModelRow;
  inherited?: Row;
  numbering: "deriveStyle" | "preserveDirect";
};

/** Fields and checks for the paragraph properties an operation sets on `pre` (absent: a new block). */
const paragraphRequest = ({
  model,
  request,
  pre,
  source,
  inherited,
  numbering: numberingMode,
}: ParagraphRequestOptions): { fields: Fields; checks: ModelRow["checks"] } => {
  const fields: Fields = {};
  const checks: ModelRow["checks"] = [];
  if ("styleId" in request) {
    const styleId = request["styleId"] as string | null;
    Object.assign(fields, styleFields(model, styleId, pre, inherited));
    const numbering = requestedStyleNumbering(model, styleId, pre ?? inherited);
    if (
      numberingMode === "deriveStyle" &&
      !("numbering" in request) &&
      numbering !== UNKNOWN_NUMBERING
    ) {
      const reference = shownReference(model, numbering);
      const styleNumberingCheck: ModelRow["checks"][number] = (row) => {
        if (reference) {
          return row.listReference?.numId === reference.numId
            ? null
            : `style/direct numbering ${reference.numId} was lost or changed`;
        }
        return row.listReference === undefined
          ? null
          : "the requested style introduced unexpected numbering";
      };
      STYLE_NUMBERING_CHECKS.add(styleNumberingCheck);
      checks.push(styleNumberingCheck);
    }
    const example =
      styleId === null ? undefined : model.styleExamples.get(styleId);
    if (
      model.mode !== "suggested" &&
      example &&
      !example.previewRuns?.some((run) => run.directFormatting !== undefined) &&
      !pre?.previewRuns?.some((run) => run.directFormatting !== undefined)
    ) {
      for (const property of ["bold", "italic", "underline"] as const) {
        const flags =
          formattingAt({ row: example, property }) ??
          Array(example.text.length).fill(false);
        if (!flags.every((flag) => flag === flags[0])) continue;
        const expected = flags[0] === true;
        checks.push((row) => {
          const actual =
            formattingAt({ row, property }) ??
            Array(row.text.length).fill(false);
          return actual.every(
            (flag, offset) =>
              flag === expected ||
              source.formats.some(
                (format) =>
                  format.property === property &&
                  offset >= format.start &&
                  offset < format.end,
              ),
          )
            ? null
            : `effective ${property} differs from style ${styleId}`;
        });
      }
    }
  }
  if ("alignment" in request) {
    fields.directAlignment =
      (request["alignment"] as string | null) ?? undefined;
  }
  if ("spacing" in request) {
    const spacing = request["spacing"] as Record<string, unknown> | null;
    checks.push((row) => {
      if (spacing === null) {
        return row.directSpacing === undefined
          ? null
          : `direct spacing ${JSON.stringify(row.directSpacing)} was not cleared`;
      }
      const missing = Object.entries(spacing).filter(
        ([key, value]) => row.directSpacing?.[key] !== value,
      );
      return missing.length === 0
        ? null
        : `direct spacing ${JSON.stringify(row.directSpacing)} lacks ${JSON.stringify(spacing)}`;
    });
  }
  if ("numbering" in request) {
    const numbering = request["numbering"] as Record<string, unknown>;
    const preKind = pre?.kind ?? inherited?.kind;
    const requestedStyleId = request["styleId"];
    assert.ok(
      requestedStyleId === undefined ||
        requestedStyleId === null ||
        typeof requestedStyleId === "string",
      "A requested style id must be a string or null",
    );
    const styleId =
      requestedStyleId === undefined
        ? (pre?.styleId ?? inherited?.styleId ?? null)
        : requestedStyleId;
    const styleNumbering = requestedStyleNumbering(model, styleId);
    if (styleNumbering === UNKNOWN_NUMBERING) {
      fields.listLevel = ANY;
    } else if (numbering["kind"] === "inherit") {
      const effective = styleNumbering;
      const reference = shownReference(model, effective);
      fields.listLevel = reference?.level;
      if ((pre ?? inherited) && preKind !== "heading")
        fields.kind = numberingKind(model, effective);
      checks.push((row) => {
        if (reference) {
          return row.listReference?.numId === reference.numId
            ? null
            : `style numbering ${reference.numId} was lost or changed`;
        }
        return row.listReference === undefined
          ? null
          : `style inheritance introduced unexpected numbering ${JSON.stringify(row.listReference)}`;
      });
    } else if (numbering["kind"] === "none") {
      fields.listLevel = undefined;
      checks.push((row) =>
        row.listReference === undefined
          ? null
          : `still numbered ${JSON.stringify(row.listReference)}`,
      );
      if (preKind === "listItem") fields.kind = "paragraph";
    } else if (numbering["kind"] === "newList") {
      const kind = numbering["format"];
      fields.listLevel =
        typeof numbering["level"] === "number" ? numbering["level"] : 0;
      if ((pre ?? inherited) && preKind !== "heading") fields.kind = "listItem";
      const used = new Set(
        model.rows.flatMap((row) =>
          row.pre?.listReference ? [row.pre.listReference.numId] : [],
        ),
      );
      checks.push((row) => {
        if (!row.listReference) return "not numbered";
        if (used.has(row.listReference.numId))
          return `joined list ${row.listReference.numId} instead of starting one`;
        if (kind === "bullet" && !isBullet(row.displayLabel))
          return `a new bullet list reads "${row.displayLabel}"`;
        if (kind === "numbered" && isBullet(row.displayLabel))
          return `a new numbered list reads "${row.displayLabel}"`;
        return null;
      });
    } else if (numbering["kind"] === "reference") {
      const numId = numbering["numId"];
      assert.ok(
        typeof numId === "number",
        "A reference request must name a numeric list id",
      );
      const effective = numberingOver(styleNumbering, {
        kind: "reference",
        numId,
        ...(typeof numbering["ilvl"] === "number"
          ? { ilvl: numbering["ilvl"] }
          : {}),
      });
      const reference = shownReference(model, effective);
      fields.listLevel = reference?.level;
      if ((pre ?? inherited) && preKind !== "heading")
        fields.kind = numberingKind(model, effective);
      if (reference)
        checks.push((row) =>
          row.listReference?.numId === reference.numId &&
          row.listReference?.level === reference.level
            ? null
            : `numbered ${JSON.stringify(row.listReference)}, not ${JSON.stringify(reference)}`,
        );
    } else if (numbering["kind"] === "levelOnly") {
      const effective = numberingOver(styleNumbering, {
        kind: "levelOnly",
        ilvl: numbering["ilvl"] as number,
      });
      const reference = shownReference(model, effective);
      fields.listLevel = reference?.level;
      if ((pre ?? inherited) && preKind !== "heading")
        fields.kind = numberingKind(model, effective);
      checks.push((row) => {
        if (reference) {
          return row.listReference?.numId === reference.numId &&
            row.listReference?.level === reference.level
            ? null
            : `numbered ${JSON.stringify(row.listReference)}, not ${JSON.stringify(reference)}`;
        }
        return row.listReference === undefined
          ? null
          : `level-only override unexpectedly numbered ${JSON.stringify(row.listReference)}`;
      });
    }
  }
  return { fields, checks };
};

/**
 * A comment the operation adds. `anchor` reads its anchored text once the
 * whole batch is modelled, or nothing when the batch changed that text.
 */
const addComment = (
  model: Model,
  operation: Operation,
  anchor?: () => string | undefined,
  removed?: () => boolean,
): void => {
  const comment = operation["comment"] as { text?: string } | undefined;
  if (typeof comment?.text === "string") {
    model.comments.push({
      text: comment.text,
      ...(anchor === undefined ? {} : { anchor }),
      ...(removed === undefined ? {} : { removed }),
    });
  }
};

const coveredRangeWasRemoved = (
  row: ModelRow,
  start: number,
  end: number,
): boolean =>
  start < end &&
  Array.from({ length: end - start }, (_, index) => start + index).every(
    (offset) =>
      row.edits.some(
        ({ start: cutStart, end: cutEnd, replace }) =>
          replace === "" && cutStart <= offset && offset < cutEnd,
      ),
  );

/** Whether the row's text is still the pre-state's, so offsets into it hold. */
const untouchedText = (row: ModelRow): boolean =>
  row.edits.length === 0 &&
  row.splits.length === 0 &&
  row.mergeSeparator === undefined;

type Expect = (model: Model, operation: Operation) => void;

/** What a paragraph property edit asks of its block's result. */
type ParagraphPropertiesRequest = {
  fields: Fields;
  checks: ModelRow["checks"];
  /** The list the block reads, when the edit names a style or numbering. */
  listLineage?: ListLineage;
};

/** Give `row` the paragraph properties `operation` sets; what that asks of the result. */
const setParagraphProperties = (
  model: Model,
  row: ModelRow,
  operation: Operation,
): ParagraphPropertiesRequest => {
  const properties = operation["properties"];
  if (isRecord(properties)) {
    row.listLineage = operationListLineage({
      model,
      request: properties,
      operationId: String(operation["id"]),
      source: row,
      inheritsFormatting: true,
      replacesDirect: true,
    });
  }
  const { fields, checks } = paragraphRequestForLiveRow({
    model,
    request: properties as ParagraphRequest,
    row,
  });
  if (isRecord(properties)) retireStyleNumberingChecks(row, properties);
  Object.assign(row.fields, fields);
  row.checks.push(...checks);
  return {
    fields,
    checks,
    ...(isRecord(properties) &&
      ("styleId" in properties || "numbering" in properties) && {
        listLineage: row.listLineage,
      }),
  };
};

/**
 * The paragraphs of the text boxes drawn in `row`, which a reader lists just
 * after it; none for a paragraph inside a text box.
 */
const textBoxRowsOf = (model: Model, row: ModelRow): ModelRow[] => {
  const boxed = (candidate: ModelRow): boolean =>
    candidate.pre !== undefined &&
    model.inTextBox?.has(candidate.pre.id) === true;
  if (boxed(row)) return [];
  const rows: ModelRow[] = [];
  for (const candidate of model.rows.slice(model.rows.indexOf(row) + 1)) {
    if (!boxed(candidate)) break;
    rows.push(candidate);
  }
  return rows;
};

/**
 * A text box's paragraphs belong to the paragraph it is drawn in: a block
 * inserted after that paragraph follows them.
 */
const pastTextBoxes = (model: Model, row: ModelRow): ModelRow =>
  textBoxRowsOf(model, row).at(-1) ?? row;

/**
 * Where a block inserted next to `row` goes: next to the row itself, or, for
 * a row inside a table, next to the outermost table (the documented rule:
 * an insertion anchored in a cell lands beside the table, not in the cell).
 */
const insertionAnchor = (
  model: Model,
  row: ModelRow,
  position: "before" | "after",
): ModelRow => {
  if (position === "after") row = pastTextBoxes(model, row);
  const table = row.pre?.table;
  if (!table) return row;
  const rows = model.rows.filter(
    (candidate) =>
      candidate.pre?.table?.outerTableIndex === table.outerTableIndex,
  );
  return (position === "before" ? rows[0] : rows.at(-1)) ?? row;
};

const liveParagraph = (row: ModelRow): Row | undefined => {
  if (!row.pre) return undefined;
  const live = { ...row.pre };
  for (const [key, value] of Object.entries(row.fields)) {
    if (value !== ANY) Object.assign(live, { [key]: value });
  }
  return live;
};

type LiveParagraphRequestOptions = {
  model: Model;
  request: ParagraphRequest;
  row: ModelRow;
  inherited?: Row;
  /** The request makes a new block beside `row` rather than editing it. */
  insertion?: boolean;
};

const paragraphRequestForLiveRow = ({
  model,
  request,
  row,
  inherited,
  insertion = false,
}: LiveParagraphRequestOptions): ReturnType<typeof paragraphRequest> => {
  // A new block is not the row it sits beside: it copies that row's live
  // properties only when it inherits them, and has no paragraph of its own.
  const live = insertion ? inherited : liveParagraph(row);
  const preserveDirect =
    "styleId" in request &&
    !("numbering" in request) &&
    (!insertion || inherited !== undefined) &&
    hasFixedDirectNumbering(row.listLineage);
  const result = paragraphRequest({
    model,
    request,
    pre: insertion ? undefined : live,
    source: row,
    inherited,
    numbering: preserveDirect ? "preserveDirect" : "deriveStyle",
  });
  if (preserveDirect && live) {
    // Direct numbering outlives the restyle, but a heading style still makes
    // the block a heading.
    if (result.fields.kind !== "heading") result.fields.kind = live.kind;
    result.fields.listLevel = live.listLevel;
  }
  return result;
};
/**
 * `named` when it is a text box paragraph: a block inserted next to it goes
 * in that box, and goes with it when the paragraph the box is drawn in is
 * deleted.
 */
const textBoxHolder = (model: Model, named: ModelRow): ModelRow | undefined =>
  named.pre !== undefined && model.inTextBox?.has(named.pre.id) === true
    ? named
    : undefined;

const insertBlock =
  (position: "before" | "after"): Expect =>
  (model, operation) => {
    const named = target(model, operation["blockId"], { adjacent: true });
    const anchor = insertionAnchor(model, named, position);
    const inherited =
      operation["inheritFormatting"] === false
        ? undefined
        : liveParagraph(named);
    const { fields: requested, checks } = paragraphRequestForLiveRow({
      model,
      request: operation,
      row: named,
      inherited,
      insertion: true,
    });
    const fields = {
      ...(inherited ? inheritedListFields(model, inherited) : {}),
      ...requested,
    };
    const scope = operation["formattingScope"] ?? "firstParagraph";
    const texts =
      operation["lineBreakMode"] === "inline"
        ? [String(operation["text"])]
        : paragraphsOf(String(operation["text"]));
    // `""` inserts a blank paragraph, which no reader lists.
    const holder = textBoxHolder(model, named);
    const rows = texts.map((text, index) => {
      const row = modelRow(text);
      if (index === 0 || scope === "allParagraphs") {
        row.fields = fields;
        row.checks = checks;
        row.listLineage = operationListLineage({
          model,
          request: operation,
          operationId: String(operation["id"]),
          source: named,
          inheritsFormatting: operation["inheritFormatting"] !== false,
        });
      }
      if (holder) row.inTextBoxOf = holder;
      return row;
    });
    pushBesideBlocks(model, anchor, position, rows);
    addComment(model, operation);
  };

/**
 * Add `cells`, new blocks of a table, to the gap on `side` of one of its
 * blocks: after any the batch put there before them, and nearer the table
 * than a block inserted beside the table.
 */
const addToTable = (
  row: ModelRow,
  side: "before" | "after",
  cells: readonly ModelRow[],
) => {
  const gap = row[side];
  const beside = gap.filter((entry) => entry.besideTable === true);
  const inside = [
    ...gap.filter((entry) => entry.besideTable !== true),
    ...cells,
  ];
  gap.splice(
    0,
    gap.length,
    ...(side === "after" ? [...inside, ...beside] : [...beside, ...inside]),
  );
};

/**
 * The blocks inserted between `row` and its neighbour on `position`'s side.
 * An insertion after one block and one before the next land in one gap, in
 * batch order.
 */
const gapOf = (
  model: Model,
  row: ModelRow,
  position: "before" | "after",
): ModelRow[] => {
  if (position === "after") return row.after;
  const previous = model.rows[model.rows.indexOf(row) - 1];
  return previous && !previous.pendingJoin ? previous.after : row.before;
};

/**
 * Put `rows`, blocks inserted outside every table, in the gap on `position`'s
 * side of `anchor`. When a table's own block borders that gap, they are
 * beside the table: rows and cells the batch adds to it stay nearer it.
 */
const pushBesideBlocks = (
  model: Model,
  anchor: ModelRow,
  position: "before" | "after",
  rows: readonly ModelRow[],
) => {
  const gap = gapOf(model, anchor, position);
  const previous = model.rows[model.rows.indexOf(anchor) - 1];
  const bordering =
    position === "after" || gap === anchor.before ? anchor : previous;
  const nearTable = bordering?.pre?.table !== undefined;
  for (const row of rows) if (nearTable) row.besideTable = true;
  gap.push(...rows);
};

const rangeOf = (
  operation: Operation,
): { blockId: string; start: number; end: number } => {
  const range = operation["range"] as {
    blockId: string;
    startOffset: number;
    endOffset: number;
  };
  return {
    blockId: range.blockId,
    start: range.startOffset,
    end: range.endOffset,
  };
};

/**
 * One block-text expectation per public operation type. `null` leaves text
 * unmodelled for that type; the separate table model still checks its grid.
 * A new operation type fails the completeness scenario until it is classified.
 */
export const EXPECTATIONS = {
  replaceInBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const find = String(operation["find"]);
    const text = row.pre?.text ?? row.text;
    const start = text.indexOf(find);
    if (start === -1)
      throw new Error(`applied, but "${find}" is not in "${text}"`);
    row.edits.push({
      start,
      end: start + find.length,
      replace: String(operation["replace"]),
    });
    addComment(model, operation);
  },
  replaceRange: (model, operation) => {
    const { blockId, start, end } = rangeOf(operation);
    target(model, blockId).edits.push({
      start,
      end,
      replace: String(operation["replace"]),
    });
    addComment(model, operation);
  },
  formatRange: (model, operation) => {
    const { blockId, start, end } = rangeOf(operation);
    const row = target(model, blockId);
    const formatting = operation["formatting"] as Record<string, unknown>;
    for (const property of ["bold", "italic", "underline"] as const) {
      if (formatting[property] === true)
        row.formats.push({ start, end, property });
    }
  },
  commentOnRange: (model, operation) => {
    const { blockId, start, end } = rangeOf(operation);
    const row = target(model, blockId);
    addComment(
      model,
      operation,
      () =>
        untouchedText(row) && !row.removed
          ? row.text.slice(start, end)
          : undefined,
      () => coveredRangeWasRemoved(row, start, end),
    );
  },
  commentOnBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const quote = operation["quote"];
    const anchored = typeof quote === "string" ? quote : undefined;
    const start = anchored === undefined ? 0 : row.text.indexOf(anchored);
    const end = start + (anchored?.length ?? row.text.length);
    addComment(
      model,
      operation,
      () =>
        untouchedText(row) && !row.removed ? (anchored ?? row.text) : undefined,
      () => start >= 0 && coveredRangeWasRemoved(row, start, end),
    );
  },
  insertAfterBlock: insertBlock("after"),
  insertBeforeBlock: insertBlock("before"),
  replaceBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    row.edits.push({
      start: 0,
      end: row.text.length,
      replace: String(operation["text"]),
    });
    if ("styleId" in operation) {
      row.listLineage = operationListLineage({
        model,
        request: operation,
        operationId: String(operation["id"]),
        source: row,
        inheritsFormatting: true,
      });
      const { fields, checks } = paragraphRequestForLiveRow({
        model,
        request: { styleId: operation["styleId"] },
        row,
      });
      Object.assign(row.fields, fields);
      row.checks.push(...checks);
    }
    addComment(model, operation);
  },
  deleteBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    row.removed = true;
    // The text boxes drawn in the paragraph go with it once the deletion
    // resolves; the suggested live view still lists their paragraphs.
    if (model.mode === "suggested") return;
    const boxedRows = textBoxRowsOf(model, row);
    for (const boxed of boxedRows) {
      boxed.removed = true;
      boxed.removedWithTextBox = true;
    }
    removeTextBoxTables(
      model.tables,
      boxedRows.flatMap(({ pre }) => (pre === undefined ? [] : [pre.id])),
    );
  },
  splitBlock: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const offset = Number(operation["offset"]);
    const separator =
      typeof operation["separator"] === "string" ? operation["separator"] : "";
    const first = operation["firstParagraphProperties"] as
      ParagraphRequest | undefined;
    const second = operation["secondParagraphProperties"] as
      ParagraphRequest | undefined;
    if (
      (first && ("styleId" in first || "numbering" in first)) ||
      (second && ("styleId" in second || "numbering" in second))
    ) {
      row.listLineage = { type: "unknown", source: "operation" };
    }
    if (first)
      Object.assign(
        row.fields,
        paragraphRequest({
          model,
          request: first,
          pre: liveParagraph(row),
          source: row,
          numbering: "deriveStyle",
        }).fields,
      );
    row.splits.push({
      offset,
      consumed: separator.length,
      ...(second
        ? {
            second: {
              ...row.fields,
              ...paragraphRequest({
                model,
                request: second,
                pre: liveParagraph(row),
                source: row,
                numbering: "deriveStyle",
              }).fields,
            },
          }
        : {}),
    });
  },
  mergeBlockWithNext: (model, operation) => {
    const row = target(model, operation["blockId"]);
    const next = model.rows[model.rows.indexOf(row) + 1];
    if (row.pendingDeletion || next?.pendingDeletion) {
      // Whether the join keeps that paragraph's pending deletion is not modelled yet.
      throw new Unmodelled("a join with a paragraph pending deletion");
    }
    row.mergeSeparator =
      typeof operation["separator"] === "string" ? operation["separator"] : "";
    if (
      model.mode !== "direct" &&
      next?.removed &&
      isContainerCarrier(model, next)
    ) {
      // The batch runs the deletion first, so the join meets a paragraph with
      // no text left and has nothing to separate, as when no block survives.
      next.removed = false;
      next.edits.push({ start: 0, end: next.text.length, replace: "" });
      row.mergeSeparator = "";
    }
    const merged = operation["mergedParagraphProperties"] as
      ParagraphRequest | undefined;
    if (merged) {
      if ("styleId" in merged || "numbering" in merged) {
        row.listLineage = operationListLineage({
          model,
          request: merged,
          operationId: String(operation["id"]),
          source: row,
          inheritsFormatting: true,
          replacesDirect: true,
        });
      }
      Object.assign(
        row.fields,
        paragraphRequestForLiveRow({ model, request: merged, row }).fields,
      );
    }
  },
  setBlockParagraphProperties: (model, operation) => {
    setParagraphProperties(
      model,
      target(model, operation["blockId"]),
      operation,
    );
  },
  insertTable: (model, operation) => {
    const position = operation["position"] === "before" ? "before" : "after";
    const named = target(model, operation["blockId"], { adjacent: true });
    // A whole table is a peer of the outermost table its anchor sits in.
    const anchor = insertionAnchor(model, named, position);
    const rows = operation["rows"] as string[][];
    // A cell's line break starts a paragraph there; a blank one stays blank.
    const texts = rows.flatMap((cells) =>
      cells.flatMap((cell) => cell.split(/\r\n|\r|\n/u)),
    );
    const holder = textBoxHolder(model, named);
    pushBesideBlocks(
      model,
      anchor,
      position,
      texts.map((text) => {
        const row = modelRow(text);
        if (holder) row.inTextBoxOf = holder;
        return row;
      }),
    );
  },
  deleteTable: (model, operation) => {
    const { tableIndex } = tableOf(
      tableRemovalTarget(model, operation["blockId"]),
    );
    const rows = tableRows(model, tableIndex);
    for (const row of rows) row.rowDeleted = true;
    removeTableRows(model, rows);
  },
  insertTableRow: (model, operation) => {
    // A row may go in beside one the batch deletes.
    const anchor = tableOf(
      target(model, operation["blockId"], { adjacent: true }),
    );
    const inRow = tableRows(model, anchor.tableIndex).filter(
      (row) => row.pre?.table?.rowIndex === anchor.rowIndex,
    );
    const texts = (
      (operation["cellTexts"] as string[] | undefined) ?? []
    ).flatMap((text) => text.split(/\r\n|\r|\n/u));
    const added = texts.map((text) => modelRow(text));
    // Every supplied text lands in one new row, one cell after another.
    const cells: ModelRow[] = added.filter((row) => row.text.length > 0);
    for (const cell of cells) {
      cell.checks.push((row) => (row.table ? null : "is not in a table"));
    }
    const position = operation["position"] === "before" ? "before" : "after";
    const edge = position === "before" ? inRow[0] : inRow.at(-1);
    if (!edge) throw new Unmodelled("the anchor row has no blocks");
    addToTable(edge, position, added);
  },
  deleteTableRow: (model, operation) => {
    const anchor = tableOf(tableRemovalTarget(model, operation["blockId"]));
    if (hasMergedCells(model, anchor.tableIndex))
      throw new Unmodelled("merged cells");
    const rows = tableRows(model, anchor.tableIndex).filter(
      (row) => row.pre?.table?.rowIndex === anchor.rowIndex,
    );
    for (const row of rows) row.rowDeleted = true;
    removeTableRows(model, rows);
  },
  insertTableColumn: (model, operation) => {
    const anchor = tableOf(target(model, operation["blockId"]));
    if (hasMergedCells(model, anchor.tableIndex))
      throw new Unmodelled("merged cells");
    const texts = (operation["cellTexts"] as string[] | undefined) ?? [];
    const after = operation["position"] !== "before";
    const rows = tableRows(model, anchor.tableIndex);
    const rowIndexes = [
      ...new Set(rows.map((row) => row.pre?.table?.rowIndex ?? 0)),
    ].sort((left, right) => left - right);
    texts.forEach((text, index) => {
      const rowIndex = rowIndexes[index];
      if (rowIndex === undefined)
        throw new Error("more cell texts than rows were applied");
      const inRow = rows.filter((row) => row.pre?.table?.rowIndex === rowIndex);
      // The new cell follows the anchor's column (or precedes it) in each row.
      const column = anchor.gridColumnIndex;
      const cell = modelRow(text);
      // A row pending deletion keeps its place in the live table and takes
      // its new cell along: accepting the deletion removes the row. Emptied
      // cells alone do not make one; a row stays with every cell retired.
      if (
        inRow.length > 0 &&
        inRow.every(
          (row) =>
            row.outsidePreState === true ||
            (model.mode !== "direct" && row.rowDeleted === true),
        )
      )
        cell.removed = true;
      const lastBefore = inRow.findLast((row) =>
        after
          ? (row.pre?.table?.gridColumnIndex ?? 0) <= column
          : (row.pre?.table?.gridColumnIndex ?? 0) < column,
      );
      if (lastBefore) addToTable(lastBefore, "after", [cell]);
      else if (inRow[0]) addToTable(inRow[0], "before", [cell]);
    });
  },
  deleteTableColumn: (model, operation) => {
    const anchor = tableOf(tableRemovalTarget(model, operation["blockId"]));
    if (hasMergedCells(model, anchor.tableIndex))
      throw new Unmodelled("merged cells");
    removeTableRows(
      model,
      tableRows(model, anchor.tableIndex).filter(
        (row) => row.pre?.table?.gridColumnIndex === anchor.gridColumnIndex,
      ),
    );
  },
  insertSignatureTable: (model, operation) => {
    const position = operation["position"] === "before" ? "before" : "after";
    const named = target(model, operation["blockId"], { adjacent: true });
    const anchor = insertionAnchor(model, named, position);
    const parties = operation["parties"] as {
      name: string;
      signatory?: string;
      title?: string;
    }[];
    const texts = parties.flatMap(({ name, signatory, title }) => {
      const lines = [name, "", "", "_".repeat(28)];
      if (signatory) lines.push(signatory);
      if (title) lines.push(title);
      return lines;
    });
    // Like a block, a signature table anchored in a cell lands beside the
    // outermost table, so rows the batch adds to that table stay nearer it.
    pushBesideBlocks(
      model,
      anchor,
      position,
      texts.map((text) => modelRow(text)),
    );
    addComment(model, operation);
  },
  mergeTableCells: null,
  splitTableCell: null,
} as const satisfies Record<FolioDocumentOperationType, Expect | null>;

/** The public operation types, for the completeness scenario. */
export const OPERATION_TYPES: readonly string[] =
  FOLIO_DOCUMENT_OPERATION_TYPES;

/** Apply one applied operation's expectation to `model`. */
export const expectOperation = (model: Model, operation: Operation): void => {
  if (
    operation.type === "deleteBlock" &&
    typeof operation["blockId"] === "string"
  ) {
    model.deletedBlockIds.add(operation["blockId"]);
  }
  if (
    model.mode === "suggested" &&
    ["deleteTable", "deleteTableRow", "deleteTableColumn"].includes(
      operation.type,
    )
  ) {
    model.tableGaps.push(
      `${operation.type}: a pending deletion remains in the live grid`,
    );
  } else if (model.tableGaps.length === 0) {
    try {
      applyTableOperation(model.tables, operation);
    } catch (error) {
      if (error instanceof UnsupportedTableExpectation) {
        model.tableGaps.push(`${operation.type}: ${error.message}`);
      } else {
        throw error;
      }
    }
  }
  const expectation = (
    EXPECTATIONS as Record<string, Expect | null | undefined>
  )[operation.type];
  if (expectation === undefined) {
    throw new Error(
      `the oracle has no entry for operation type ${operation.type}`,
    );
  }
  if (expectation === null) {
    model.unmodelled.push(`${operation.type}: no block-text expectation yet`);
    return;
  }
  try {
    expectation(model, operation);
  } catch (error) {
    if (error instanceof Unmodelled) {
      model.unmodelled.push(`${operation.type}: ${error.message}`);
      return;
    }
    throw error;
  }
};

// ---------------------------------------------------------------------------
// Materializing and comparing
// ---------------------------------------------------------------------------

type Expected = {
  text: string;
  fields: Fields;
  checks: ModelRow["checks"];
  formats: ModelRow["formats"];
  /** The pre-state block, while its text is unchanged: formatting outside a range keeps. */
  pre?: Row;
  listLineage: ListLineage;
};

const applyEdits = (
  text: string,
  edits: readonly TextEdit[],
  splits: ModelRow["splits"],
) => {
  // Everything in pre-state coordinates, applied from the end backwards.
  type Cut = {
    start: number;
    end: number;
    replace: string;
    split?: ModelRow["splits"][number];
  };
  const cuts: Cut[] = [
    ...edits.map((edit) => ({ ...edit })),
    ...splits.map((split) => ({
      start: split.offset,
      end: split.offset + split.consumed,
      replace: "",
      split,
    })),
  ].sort((left, right) => right.start - left.start || right.end - left.end);
  const segments: { text: string; split?: ModelRow["splits"][number] }[] = [];
  const tail = text;
  let current = "";
  let cursor = text.length;
  for (const cut of cuts) {
    if (cut.end > cursor)
      throw new Error(`two applied edits overlap in "${text}"`);
    current = cut.replace + tail.slice(cut.end, cursor) + current;
    cursor = cut.start;
    if (cut.split) {
      segments.unshift({ text: current, split: cut.split });
      current = "";
    }
  }
  current = tail.slice(0, cursor) + current;
  segments.unshift({ text: current });
  return segments;
};

const materialize = (model: Model): Expected[] => {
  if (model.story === "removed") return [];
  const out: Expected[] = [];
  type Container = TableLocation | undefined;
  let joinNext: { separator: string; container: Container } | undefined;
  const sameContainer = (left: Container, right: Container): boolean =>
    left === undefined
      ? right === undefined
      : right !== undefined &&
        left.tableIndex === right.tableIndex &&
        left.rowIndex === right.rowIndex &&
        left.cellIndex === right.cellIndex;
  const push = (entry: Expected, container: Container) => {
    if (joinNext !== undefined) {
      const previous = out.at(-1);
      if (previous && sameContainer(joinNext.container, container)) {
        previous.text = `${previous.text}${joinNext.separator}${entry.text}`;
        // A join keeps the first paragraph's properties; formatting offsets no longer hold.
        previous.formats = [];
        joinNext = undefined;
        return;
      }
      if (previous) previous.text += joinNext.separator;
      joinNext = undefined;
    }
    out.push(entry);
  };
  const emit = (row: ModelRow) => {
    for (const inserted of row.before) emit(inserted);
    const gone =
      row.removed ||
      row.inTextBoxOf?.removedWithTextBox === true ||
      (row.pendingDeletion === true && row.edits.length === 0);
    if (!gone) {
      const segments = applyEdits(row.text, row.edits, row.splits);
      const whole = untouchedText(row);
      segments.forEach((segment, index) =>
        push(
          {
            text: segment.text,
            fields:
              index === 0 ? row.fields : (segment.split?.second ?? row.fields),
            checks: row.checks,
            formats: whole ? row.formats : [],
            listLineage: row.listLineage,
            ...(whole && row.pre ? { pre: row.pre } : {}),
          },
          row.pre?.table,
        ),
      );
      if (row.mergeSeparator !== undefined) {
        joinNext = { separator: row.mergeSeparator, container: row.pre?.table };
      }
    }
    for (const inserted of row.after) emit(inserted);
  };
  for (const row of model.rows) emit(row);
  return out;
};

/** A reader lists no blank block; neither side is compared on one. */
const visible = <T extends { text: string }>(rows: readonly T[]): T[] =>
  rows.filter((row) => row.text.length > 0);

type FormattingAtOptions = {
  row: Row;
  property: "bold" | "italic" | "underline";
  source?: "effective" | "direct";
};

const formattingAt = ({
  row,
  property,
  source = "effective",
}: FormattingAtOptions): (boolean | undefined)[] | null => {
  const runs = row.previewRuns;
  if (!runs || runs.map((run) => run.text).join("") !== row.text) return null;
  return runs.flatMap((run) => {
    const direct = run.directFormatting?.[property];
    const directFlag = typeof direct === "boolean" ? direct : undefined;
    const flag = source === "effective" ? run[property] === true : directFlag;
    return Array.from({ length: run.text.length }, () => flag);
  });
};

/** Compare `actual` with what `model` expects; every mismatch, as text. */
export const compareWithModel = (
  model: Model,
  actual: readonly Row[],
): string[] => {
  const expected = visible(materialize(model));
  const got = visible(actual);
  const problems: string[] = [];
  const expectedTexts = expected.map((row) => row.text);
  const gotTexts = got.map((row) => row.text);
  if (JSON.stringify(expectedTexts) !== JSON.stringify(gotTexts)) {
    problems.push(
      `block texts differ:\n    expected ${JSON.stringify(expectedTexts)}\n    got      ${JSON.stringify(gotTexts)}`,
    );
    return problems;
  }
  expected.forEach((entry, index) => {
    const row = got[index] as Row;
    for (const [key, value] of Object.entries(entry.fields)) {
      if (value === ANY) continue;
      const actualValue = row[key as keyof Row];
      if (JSON.stringify(actualValue) !== JSON.stringify(value)) {
        problems.push(
          `"${row.text}": ${key} is ${JSON.stringify(actualValue)}, expected ${JSON.stringify(value)}\n    before ${JSON.stringify(entry.pre)}\n    after ${JSON.stringify(row)}`,
        );
      }
    }
    for (const check of entry.checks) {
      const problem = check(row);
      if (problem !== null) problems.push(`"${row.text}": ${problem}`);
    }
    for (const format of entry.formats) {
      const flags = formattingAt({ row, property: format.property });
      if (flags === null) continue;
      const restyled = entry.fields.styleId !== entry.pre?.styleId;
      const requestedFlags = restyled
        ? formattingAt({ row, property: format.property, source: "direct" })
        : flags;
      const missing = requestedFlags
        ?.slice(format.start, format.end)
        .some((flag) => flag !== true);
      if (missing) {
        problems.push(
          `"${row.text}": [${format.start}, ${format.end}) is not all ${format.property}`,
        );
      }
      // A restyle may change inherited values; the range must still preserve direct provenance.
      const source = restyled ? "direct" : "effective";
      const outsideFlags = formattingAt({
        row,
        property: format.property,
        source,
      });
      const before =
        entry.pre &&
        formattingAt({ row: entry.pre, property: format.property, source });
      const inside = (offset: number) =>
        entry.formats.some(
          (other) =>
            other.property === format.property &&
            offset >= other.start &&
            offset < other.end,
        );
      const changed =
        before && outsideFlags
          ? outsideFlags.findIndex(
              (flag, offset) => !inside(offset) && flag !== before[offset],
            )
          : -1;
      if (changed !== -1) {
        problems.push(
          `"${row.text}": ${format.property} changed at ${changed}, outside what was asked\n    before runs ${JSON.stringify(entry.pre?.previewRuns)}\n    after runs ${JSON.stringify(row.previewRuns)}`,
        );
      }
    }
  });
  const newListIds = new Map<string, number>();
  const referenceIds = new Set<number>();
  for (const entry of expected) {
    if (entry.listLineage.type === "reference")
      referenceIds.add(entry.listLineage.numId);
  }
  expected.forEach((entry, index) => {
    const row = got[index];
    if (!row) return;
    switch (entry.listLineage.type) {
      case "reference":
        if (row.listReference?.numId !== entry.listLineage.numId) {
          problems.push(
            `"${row.text}": numbering list reference ${JSON.stringify(row.listReference)} differs from modeled reference ${entry.listLineage.numId}`,
          );
        }
        break;
      case "none":
        if (row.listReference !== undefined) {
          problems.push(
            `"${row.text}": unexpected numbering list reference ${JSON.stringify(row.listReference)}`,
          );
        }
        break;
      case "newList": {
        const numId = row.listReference?.numId;
        if (numId === undefined) {
          problems.push(`"${row.text}": new list has no list reference`);
          break;
        }
        const previous = newListIds.get(entry.listLineage.operationId);
        if (previous !== undefined && previous !== numId) {
          problems.push(
            `"${row.text}": new list ${entry.listLineage.operationId} has inconsistent references ${previous} and ${numId}`,
          );
        }
        newListIds.set(entry.listLineage.operationId, numId);
        break;
      }
      case "unknown":
        break;
      default: {
        const unhandled: never = entry.listLineage;
        throw new Error(`Unhandled list lineage ${JSON.stringify(unhandled)}`);
      }
    }
  });
  const claimedNewListIds = new Set<number>();
  for (const [operationId, numId] of newListIds) {
    if (referenceIds.has(numId) || claimedNewListIds.has(numId)) {
      problems.push(
        `new list ${operationId} aliases another modeled list reference ${numId}`,
      );
    }
    claimedNewListIds.add(numId);
  }
  return problems;
};

type AnchorDisposition = "kept" | "removed" | "ambiguous";

type RemovedLiveRun = { text: string; firstLength: number };

/** Whether the batch leaves no live block `id`: it deletes or removes it. */
const removedByBatch = (model: Model, id: string): boolean => {
  if (model.deletedBlockIds.has(id)) return true;
  const row = model.rows.find((candidate) => candidate.pre?.id === id);
  return (
    row !== undefined &&
    ((row.removed && row.pendingJoin !== true) || row.pendingDeletion === true)
  );
};

/**
 * The live text of `blockId` and of the blocks after it that the batch also
 * removes, joined as a comment quote joins the blocks its range spans (with
 * no separator), and where the first block's text ends. A comment whose range
 * runs from one removed block into the next goes with them.
 */
const removedLiveRun = (
  model: Model,
  blockId: string,
): RemovedLiveRun | undefined => {
  let run: RemovedLiveRun | undefined;
  for (const [id, liveText] of model.liveTextById) {
    if (run === undefined) {
      if (id === blockId)
        run = { text: liveText, firstLength: liveText.length };
      continue;
    }
    if (!removedByBatch(model, id)) break;
    run.text += liveText;
  }
  return run;
};

/** Whether `quote` starts inside the run's first block, where the comment is. */
const quotedFromFirstBlock = (run: RemovedLiveRun, quote: string): boolean => {
  const start = run.text.indexOf(quote);
  return start !== -1 && start < Math.max(run.firstLength, 1);
};

/**
 * Decide whether the model removed the anchor, accounting for repeated text.
 * `acceptedAnchor` is the comment's quote in the accepted document: a live
 * quote reads pending deleted text that the live block text leaves out.
 */
const anchorDisposition = (
  model: Model,
  entry: Comment,
  acceptedAnchor: string = entry.anchor,
): AnchorDisposition => {
  if (entry.blockId === null) return "kept";
  const row = model.rows.find(
    (candidate) => candidate.pre?.id === entry.blockId,
  );
  const explicitlyDeleted = model.deletedBlockIds.has(entry.blockId);
  if (explicitlyDeleted || (row?.removed && !row.pendingJoin)) {
    const removed = removedLiveRun(model, entry.blockId);
    return entry.anchor === "" ||
      (removed !== undefined &&
        (quotedFromFirstBlock(removed, entry.anchor) ||
          (acceptedAnchor !== "" &&
            quotedFromFirstBlock(removed, acceptedAnchor))))
      ? "removed"
      : "kept";
  }
  // A pending join is already absent from the accepted pre-state; its live
  // block has not been removed by this batch.
  if (!row?.pre || row.removed || entry.anchor === "") return "kept";
  const starts: number[] = [];
  for (
    let start = 0;
    start <= row.pre.text.length - entry.anchor.length;
    start += 1
  ) {
    if (row.pre.text.startsWith(entry.anchor, start)) starts.push(start);
  }
  if (starts.length === 0) return "kept";
  const removed = [
    ...row.edits
      .filter(({ replace }) => replace.length === 0)
      .map(({ start, end }) => ({ start, end })),
    ...row.splits.map(({ offset, consumed }) => ({
      start: offset,
      end: offset + consumed,
    })),
  ];
  const removedPlacements = starts.map((start) =>
    Array.from(
      { length: entry.anchor.length },
      (_, index) => start + index,
    ).every((offset) =>
      removed.some((cut) => cut.start <= offset && offset < cut.end),
    ),
  );
  if (removedPlacements.every(Boolean)) return "removed";
  return removedPlacements.some(Boolean) ? "ambiguous" : "kept";
};

/** A full-block anchor has an exact replacement when one edit rewrites the block. */
const expectedKeptAnchor = (
  model: Model,
  entry: Comment,
): string | undefined => {
  const row = model.rows.find(
    (candidate) => candidate.pre?.id === entry.blockId,
  );
  if (!row?.pre || row.removed || row.pendingJoin || row.splits.length > 0)
    return undefined;
  if (
    entry.anchor === "" ||
    entry.anchor !== row.pre.text ||
    row.edits.length !== 1
  ) {
    return undefined;
  }
  const edit = row.edits[0];
  if (
    !edit ||
    edit.start !== 0 ||
    edit.end !== row.pre.text.length ||
    edit.replace === ""
  ) {
    return undefined;
  }
  return edit.replace.includes("\n") ? undefined : edit.replace;
};

/** Pending markup may interleave the old text and insertion in one range. */
const containsInOrder = (text: string, expected: string): boolean => {
  let offset = 0;
  for (const character of expected) {
    const index = text.indexOf(character, offset);
    if (index === -1) return false;
    offset = index + character.length;
  }
  return true;
};

const pendingAnchorMatches = (
  actual: string,
  old: string,
  replacement: string,
): boolean => {
  if (!containsInOrder(actual, old) || !containsInOrder(actual, replacement))
    return false;
  // A retained character can serve both texts. With disjoint characters, the
  // deleted text must precede the insertion in the live tracked range.
  const replacementCharacters = new Set(replacement);
  return [...old].some((character) => replacementCharacters.has(character))
    ? true
    : containsInOrder(actual, old + replacement);
};

export const compareComments = (
  model: Model,
  before: readonly Comment[],
  after: readonly Comment[],
  afterRows: readonly Row[],
  liveBefore: readonly Comment[] = before,
  mode: Mode = "direct",
): string[] => {
  const problems: string[] = [];
  const remaining = [...after];
  const kept: Comment[] = [];
  const gone: { entry: Comment; located: Comment }[] = [];
  const storyBlockIds = new Set(
    model.rows.flatMap((row) => (row.pre === undefined ? [] : [row.pre.id])),
  );
  for (const entry of before) {
    const live = liveBefore.find(({ id }) => id === entry.id);
    // The accepted quote belongs to the accepted model. Use a live quote
    // only when a pending paragraph join moved its anchor to another block.
    const located = live && live.blockId !== entry.blockId ? live : entry;
    let disposition =
      mode === "suggested"
        ? "kept"
        : anchorDisposition(model, located, entry.anchor);
    // A suggested removal stays pending and rejectable, so its comments stay.
    if (
      mode !== "suggested" &&
      ((model.story === "removed" &&
        located.blockId !== null &&
        storyBlockIds.has(located.blockId)) ||
        (located.blockId !== null &&
          model.removedNoteStories.has(
            model.noteStoryByBlockId.get(located.blockId) ?? "",
          )))
    )
      disposition = "removed";
    if (
      mode !== "suggested" &&
      located.blockId !== null &&
      model.uncertainNoteStories.has(
        model.noteStoryByBlockId.get(located.blockId) ?? "",
      )
    )
      disposition = "ambiguous";
    if (disposition === "removed") gone.push({ entry, located });
    else if (disposition === "kept") kept.push(entry);
  }
  for (const comment of kept) {
    const anchor = expectedKeptAnchor(model, comment);
    const index = remaining.findIndex(
      ({ id, text, anchor: actualAnchor }) =>
        id === comment.id &&
        text === comment.text &&
        (anchor === undefined ||
          (mode === "suggested"
            ? pendingAnchorMatches(actualAnchor, comment.anchor, anchor)
            : actualAnchor === anchor)),
    );
    if (index === -1) {
      problems.push(
        `no comment ${JSON.stringify(comment)} among ${JSON.stringify(after)}`,
      );
    } else {
      remaining.splice(index, 1);
    }
  }
  const beforeIds = new Set(before.map(({ id }) => id));
  for (const comment of model.comments) {
    if (comment.removed?.()) continue;
    const anchor = comment.anchor?.();
    const index = remaining.findIndex(
      (candidate) =>
        !beforeIds.has(candidate.id) &&
        candidate.text === comment.text &&
        (anchor === undefined || candidate.anchor === anchor),
    );
    if (index === -1) {
      problems.push(
        `no comment ${JSON.stringify({ text: comment.text, anchor })} among ${JSON.stringify(after)}`,
      );
    } else {
      remaining.splice(index, 1);
    }
  }
  // A new comment the model removed must not survive among the unclaimed ones.
  for (const comment of model.comments) {
    if (!comment.removed?.()) continue;
    const anchor = comment.anchor?.();
    const stale = remaining.find(
      (candidate) =>
        !beforeIds.has(candidate.id) &&
        candidate.text === comment.text &&
        (anchor === undefined || candidate.anchor === anchor),
    );
    if (stale !== undefined) {
      problems.push(
        `removed comment ${JSON.stringify({ text: comment.text, anchor })} remains among ${JSON.stringify(after)}`,
      );
    }
  }
  // Only what no expectation claimed can be the removed comment itself.
  const survivingBlockIds = new Set(afterRows.map(({ id }) => id));
  for (const { entry, located } of gone) {
    const survivor = remaining.find(({ id }) => id === entry.id);
    // A block id locates the first paragraph containing an anchor; it does not
    // own the whole range. Splits and joins can move the surviving part of a
    // range to another paragraph before this batch deletes the old locator.
    // In that case the comment is still anchored and must survive. A thread
    // left on the deleted locator, moved to another block the batch also
    // deleted, or left in no block at all, is the orphan this check is meant
    // to catch.
    if (
      survivor &&
      (survivor.blockId === located.blockId ||
        survivor.blockId === null ||
        !survivingBlockIds.has(survivor.blockId))
    ) {
      problems.push(
        `comment ${JSON.stringify(entry)} outlived the block it anchored`,
      );
    }
  }
  return problems;
};

const compareStyles = async (
  model: Model,
  bytes: Uint8Array,
): Promise<string[]> => {
  if (model.styles.size === 0) return [];
  const catalog = await inspectDocumentStylesFromDocx(toArrayBuffer(bytes));
  const paragraphStyles = new Set(
    catalog.styles
      .filter((style) => style.type === "paragraph")
      .map((style) => style.styleId),
  );
  return [...model.styles]
    .filter((styleId) => !paragraphStyles.has(styleId))
    .map(
      (styleId) => `the saved package defines no paragraph style "${styleId}"`,
    );
};

// ---------------------------------------------------------------------------
// The check around one apply
// ---------------------------------------------------------------------------

/**
 * The state an apply is checked against. `direct` and `tracked-changes`:
 * saved, reopened and every change accepted, so pending changes from before
 * resolve the same way on both sides. `suggested`: the live reviewer, whose
 * suggestions stay out of the package until accepted.
 */
export type Pre = {
  mode: Mode;
  live: string;
  rows: Row[];
  /** Whether the story of `rows` is in that state. */
  resolvedStory: StoryPresence;
  /** The reviewer's own blocks, when `rows` are another view of them. */
  liveRows: Row[];
  /** Which live blocks accepting runs on into the next one. */
  joinSources: JoinSources;
  /** Immutable pre-state; suggestions have not entered saved package bytes yet. */
  numberingSource:
    | { type: "package"; bytes: Uint8Array; live: NumberingFacts }
    | { type: "live"; facts: NumberingFacts };
  comments: Comment[];
  links: LinkSnapshot;
  liveComments: Comment[];
  rejected?: Row[];
  /** The story the operations target; the body by default. */
  story: Story;
  /** The session the step runs in, for the coverage ledger. */
  step: StepKind;
  /** What each block of the story is, for the coverage ledger. */
  targets: FeatureIndex;
};

export type CaptureOptions = { story?: Story; step?: StepKind };

const liveState = (reviewer: Reviewer, story: Story = MAIN): string =>
  JSON.stringify({
    blocks: reviewer.getContent(),
    ...(story.type === "main" ? {} : { story: rowsOf(reviewer, story) }),
    comments: reviewer.getComments().length,
  });

const project = (rows: readonly Row[]) =>
  visible(rows).map(({ text, kind, styleId, headingLevel, listLevel }) => ({
    text,
    kind,
    styleId,
    headingLevel,
    listLevel,
  }));

/** Per-character effective formatting survives a pending paragraph join. */
const inlineProject = (rows: readonly Row[]) =>
  visible(rows).flatMap((row) =>
    (row.previewRuns ?? [{ text: row.text }]).flatMap((run) =>
      Array.from(run.text, (character) => ({
        character,
        bold: run.bold === true,
        italic: run.italic === true,
        underline: run.underline === true,
      })),
    ),
  );

/** Authored run properties survive style inheritance changes on acceptance. */
const authoredInlineProject = (rows: readonly Row[]) =>
  visible(rows).flatMap((row) =>
    (row.previewRuns ?? [{ text: row.text }]).flatMap((run) =>
      Array.from(run.text, (character) => ({
        character,
        directFormatting: run.directFormatting,
      })),
    ),
  );

/** Reader-visible paragraph, table and effective run properties, without unstable ids. */
const fullProject = (rows: readonly Row[]) =>
  visible(rows).map((row) => ({
    text: row.text,
    kind: row.kind,
    styleId: row.styleId,
    headingLevel: row.headingLevel,
    listLevel: row.listLevel,
    directAlignment: row.directAlignment,
    directSpacing: row.directSpacing,
    table: row.table,
    runs: inlineProject([row]),
  }));

const suggestedProject = (rows: readonly Row[]) =>
  visible(rows).map((row) => ({
    text: row.text,
    kind: row.kind,
    styleId: row.styleId,
    headingLevel: row.headingLevel,
    listLevel: row.listLevel,
    directAlignment: row.directAlignment,
    directSpacing: row.directSpacing,
    table: row.table,
    runs: authoredInlineProject([row]),
  }));

/** Capture what the oracle needs before an apply. */
export const capture = async (
  reviewer: Reviewer,
  mode: Mode,
  { story = MAIN, step = "fresh" }: CaptureOptions = {},
): Promise<Pre> => {
  const live = liveState(reviewer, story);
  const liveComments = commentsOf(reviewer);
  const context = { story, step, targets: featureIndex(reviewer, story) };
  if (mode === "suggested") {
    const rows = rowsOf(reviewer, story);
    return {
      mode,
      live,
      rows,
      resolvedStory: "present",
      liveRows: rows,
      // The suggested model reads the live rows themselves: no block is
      // missing from them, so none joins another.
      joinSources: { type: "marked", ids: new Set() },
      numberingSource: {
        type: "live",
        facts: liveNumberingFacts(reviewer, story),
      },
      comments: liveComments,
      liveComments,
      links: captureLinks(reviewer),
      ...context,
    };
  }
  const bytes = await save(reviewer);
  const accepted = await resolvedState(bytes, "accept", story);
  return {
    mode,
    live,
    rows: accepted.rows,
    resolvedStory: accepted.story,
    liveRows: rowsOf(reviewer, story),
    joinSources: await markedJoinSourcesOf(bytes),
    numberingSource: {
      type: "package",
      bytes: accepted.bytes,
      live: liveNumberingFacts(reviewer, story),
    },
    comments: accepted.comments,
    links: accepted.links,
    liveComments,
    ...(mode === "tracked-changes"
      ? { rejected: (await resolvedState(bytes, "reject", story)).rows }
      : {}),
    ...context,
  };
};

export type Outcome = {
  /** The operations the receipt says applied, in request order. */
  applied: readonly AppliedOperation[];
  /** Every operation asked for, applied or not (the ledger counts refusals); `applied` if absent. */
  attempted?: readonly Operation[];
};

const LEDGER_MODES: Record<Mode, string> = {
  direct: "direct",
  "tracked-changes": "tracked",
  suggested: "suggested",
};

/** The offsets an operation acts at, in its block's text, when it names any. */
const offsetsOf = (
  operation: Operation,
  text: string,
): [number, number] | null => {
  const range = operation["range"] as
    { startOffset?: unknown; endOffset?: unknown } | undefined;
  if (
    typeof range?.startOffset === "number" &&
    typeof range.endOffset === "number"
  ) {
    return [range.startOffset, range.endOffset];
  }
  if (typeof operation["offset"] === "number")
    return [operation["offset"], operation["offset"]];
  const find = operation["find"] ?? operation["quote"];
  const at =
    typeof find === "string" && find.length > 0 ? text.indexOf(find) : -1;
  return at === -1 ? null : [at, at + (find as string).length];
};

/** Record every attempted operation of an outcome in the coverage ledger. */
const recordOutcome = (pre: Pre, outcome: Outcome): void => {
  const applied: Set<Operation> = new Set(outcome.applied);
  for (const operation of outcome.attempted ?? outcome.applied) {
    if (typeof operation !== "object" || operation === null) continue;
    const range = operation["range"] as { blockId?: unknown } | undefined;
    const blockId = operation["blockId"] ?? range?.blockId;
    const block = pre.liveRows.find((row) => row.id === blockId);
    const features = new Set<Feature | "none">(
      pre.targets.features.get(String(blockId)) ?? [],
    );
    // A block with an astral character counts as a surrogate boundary only
    // where the operation's own offsets meet one.
    features.delete("surrogateBoundary");
    const offsets = block && offsetsOf(operation, block.text);
    if (
      block &&
      offsets &&
      touchesSurrogate(block.text, offsets[0], offsets[1])
    ) {
      features.add("surrogateBoundary");
    }
    if (features.size === 0) features.add("none");
    const story = storyKindOf(pre.targets, block);
    for (const feature of targetFeatureSignature(block, features, story)) {
      recordFeatureHit({
        operation: String(operation.type),
        feature,
        selection: operationSelection(operation),
      });
    }
    for (const feature of features) {
      recordHit(
        {
          op: String(operation.type),
          story,
          mode: LEDGER_MODES[pre.mode],
          feature,
          step: pre.step,
        },
        applied.has(operation),
      );
    }
  }
};

/**
 * Check an apply's outcome against what was asked. Throws with every
 * mismatch; returns the gaps the oracle could not model.
 */
export const assertRequestedOutcome = async (
  reviewer: Reviewer,
  pre: Pre,
  outcome: Outcome,
  context: string,
): Promise<string[]> => {
  recordOutcome(pre, outcome);
  if (outcome.applied.length === 0) {
    assert.equal(
      liveState(reviewer, pre.story),
      pre.live,
      `${context}: nothing applied, but the document changed`,
    );
    return [];
  }
  const model = modelOf(pre.rows, pre.liveRows, pre.joinSources);
  if (
    outcome.applied.some(
      (operation) =>
        operation.type === "insertAfterBlock" ||
        operation.type === "insertBeforeBlock" ||
        "styleId" in operation ||
        "numbering" in operation ||
        [
          "properties",
          "firstParagraphProperties",
          "secondParagraphProperties",
          "mergedParagraphProperties",
        ].some((key) => {
          const properties = operation[key];
          return (
            properties !== null &&
            typeof properties === "object" &&
            ("styleId" in properties || "numbering" in properties)
          );
        }),
    )
  ) {
    switch (pre.numberingSource.type) {
      case "package":
        model.numberingFacts = await numberingFactsOf(
          pre.numberingSource.bytes,
        );
        model.liveNumberingFacts = pre.numberingSource.live;
        break;
      case "live":
        model.numberingFacts = pre.numberingSource.facts;
        break;
      default: {
        const unhandled: never = pre.numberingSource;
        throw new Error(
          `Unhandled numbering source ${JSON.stringify(unhandled)}`,
        );
      }
    }
  }
  model.inTextBox = pre.targets.inTextBox;
  model.noteStoryByBlockId = new Map(
    [...pre.targets.noteStoryByBlockId].flatMap(([blockId, story]) =>
      story.type === "footnote" || story.type === "endnote"
        ? [[blockId, noteStoryKey(story)] as const]
        : [],
    ),
  );
  model.mode = pre.mode;
  model.story = pre.resolvedStory;
  const removedBlocksBy = new Map<string, { id: string; type: string }>();
  /** Paragraph property edits of blocks a deletion empties and keeps. */
  const keptBlockFormats: {
    id: string;
    /** Absent when the oracle cannot model the edit (an oracle gap). */
    requested: ParagraphPropertiesRequest | undefined;
  }[] = [];
  // The note references each live block holds, followed through merges.
  const noteReferences = new Map<string, readonly FolioDocumentStoryHandle[]>(
    pre.targets.noteReferences,
  );
  const replayOperations = sequentialGroups({
    applied: outcome.applied,
    preRows: pre.rows,
    liveRows: pre.liveRows,
    mode: pre.mode,
  }).flat();
  for (const operation of replayOperations) {
    const index = outcome.applied.findIndex(({ id }) => id === operation.id);
    const range = operation["range"];
    const targetBlockId =
      typeof range === "object" && range !== null && "blockId" in range
        ? range.blockId
        : operation["blockId"];
    const targetsBlockIdentity = ![
      "insertAfterBlock",
      "insertBeforeBlock",
      "insertSignatureTable",
      "insertTable",
      "deleteTable",
      "insertTableRow",
      "deleteTableRow",
      "insertTableColumn",
      "deleteTableColumn",
      "mergeTableCells",
      "splitTableCell",
    ].includes(operation.type);
    const removedBy =
      targetsBlockIdentity && typeof targetBlockId === "string"
        ? removedBlocksBy.get(targetBlockId)
        : undefined;
    const continuesMergeChain =
      operation.type === "mergeBlockWithNext" &&
      removedBy?.type === "mergeBlockWithNext";
    // The shared order runs a property edit of a paragraph before a deletion
    // that takes the paragraph out (`orderAddressingBeforeRemoval`). One
    // replayed after the deletion formats the paragraph that deletion keeps:
    // the last of its story or table cell before the batch, which it only
    // empties. The result must still hold that block, emptied and last, with
    // the requested properties (`deletionKeptBlock`, checked once it is read).
    const formatsKeptBlock =
      operation.type === "setBlockParagraphProperties" &&
      removedBy?.type === "deleteBlock" &&
      typeof targetBlockId === "string" &&
      endsContainer(pre.rows, targetBlockId);
    if (removedBy !== undefined && !continuesMergeChain && !formatsKeptBlock) {
      throw new Error(
        `applied an operation on block ${targetBlockId}, which an earlier one of the batch removed`,
      );
    }
    if (formatsKeptBlock) {
      // The model row stays removed, so the text comparison skips it; the
      // edit's expectation is checked on the kept block itself.
      const row = model.rows.find(
        (candidate) => candidate.pre?.id === targetBlockId,
      );
      let requested: ParagraphPropertiesRequest | undefined;
      try {
        if (row === undefined) {
          throw new Unmodelled(
            `block ${targetBlockId} is not in the resolved pre-state`,
          );
        }
        requested = setParagraphProperties(model, row, operation);
      } catch (error) {
        if (!(error instanceof Unmodelled)) throw error;
        model.unmodelled.push(`${operation.type}: ${error.message}`);
      }
      keptBlockFormats.push({ id: targetBlockId, requested });
    } else {
      expectOperation(model, operation);
    }

    if (pre.mode !== "direct" && operation.type === "mergeBlockWithNext") {
      const anchorIndex = pre.rows.findIndex(
        ({ id }) => id === operation["blockId"],
      );
      const carrier = anchorIndex < 0 ? undefined : pre.rows[anchorIndex + 1];
      const carrierRow = model.rows.find((row) => row.pre?.id === carrier?.id);
      if (carrierRow !== undefined && isContainerCarrier(model, carrierRow)) {
        removedBlocksBy.delete(carrier?.id ?? "");
      }
    }

    const blockId = operation["blockId"];
    if (typeof blockId !== "string") continue;
    const removal = (() => {
      switch (operation.type) {
        case "deleteBlock":
        case "mergeBlockWithNext":
        case "deleteTable":
        case "deleteTableRow":
        case "deleteTableColumn":
          const getRemovedBlockIds =
            pre.mode === "direct"
              ? getFolioDocumentOperationRemovedBlockIds
              : getTrackedFolioDocumentOperationRemovedBlockIds;
          return getRemovedBlockIds(
            pre.rows.map(({ id, table }) => ({
              id,
              ...(table !== undefined && {
                table: { ...table, paragraphIndex: 0 },
              }),
            })),
            {
              id: `oracle-${String(index)}`,
              type: operation.type,
              blockId,
            },
          );
        default:
          return [];
      }
    })();
    const removedIds = [...removal];
    for (const story of retireNoteReferences(
      noteReferences,
      operation.type,
      blockId,
      removedIds,
    )) {
      model.removedNoteStories.add(story);
    }
    if (removedIds.length === 0 && !KEEPS_NOTE_REFERENCES.has(operation.type)) {
      // A split, replacement or other edit of a block may move or delete the
      // note references it holds; the oracle does not follow them.
      for (const story of noteReferences.get(blockId) ?? []) {
        if (story.type === "footnote" || story.type === "endnote") {
          model.uncertainNoteStories.add(noteStoryKey(story));
        }
      }
    }
    for (const removedBlockId of removal) {
      if (!removedBlocksBy.has(removedBlockId)) {
        removedBlocksBy.set(removedBlockId, {
          id: String(operation["id"] ?? index),
          type: operation.type,
        });
      }
    }
  }
  // An operation the oracle cannot model changes the document in a way it
  // cannot predict; the rest of the batch is not compared either.
  const predictable = model.unmodelled.length === 0;

  let rows: Row[];
  let comments: Comment[];
  let links: LinkSnapshot;
  let bytes: Uint8Array | null = null;
  if (pre.mode === "suggested") {
    rows = rowsOf(reviewer, pre.story);
    comments = commentsOf(reviewer);
    links = captureLinks(reviewer);
  } else {
    const saved = await save(reviewer);
    const accepted = await resolvedState(saved, "accept", pre.story);
    ({ rows, comments, bytes, links } = accepted);
    if (pre.rejected) {
      assert.deepEqual(
        project((await resolvedState(saved, "reject", pre.story)).rows),
        project(pre.rejected),
        `${context}: rejecting every change does not give the document before back`,
      );
    }
  }
  const keptBlocks = keptBlockFormats.map(({ id, requested }) => {
    const row = rows.find((candidate) => candidate.id === id);
    if (
      row === undefined ||
      !deletionKeptBlock({ before: pre.rows, after: rows, id })
    ) {
      throw new Error(
        `applied an operation on block ${id}, which an earlier one of the batch removed`,
      );
    }
    return { row, requested };
  });
  const problems = [
    ...(predictable
      ? keptBlocks.flatMap(({ row, requested }) =>
          requested === undefined ? [] : keptBlockProblems(requested, row),
        )
      : []),
    ...(predictable ? compareWithModel(model, rows) : []),
    ...(model.tableGaps.length === 0
      ? compareTableGeometry(model.tables, rows)
      : []),
    ...comparePreservedLinks({
      before: pre.links,
      after: links,
      afterRows: rows,
    }).problems,
    ...compareComments(
      model,
      pre.comments,
      comments,
      rows,
      pre.liveComments,
      pre.mode,
    ),
    ...(bytes ? await compareStyles(model, bytes) : []),
  ];
  if (
    process.env["FOLIO_ORACLE_GAPS"] &&
    (!predictable || model.tableGaps.length > 0)
  ) {
    console.log(
      `oracle gap: ${context}: ${[...model.unmodelled, ...model.tableGaps].join("; ")}`,
    );
  }
  if (problems.length > 0) {
    throw new Error(
      `${context}: the result is not what was asked (${outcome.applied
        .map((operation) => operation.type)
        .join(", ")}):\n  ${problems.join("\n  ")}`,
    );
  }
  return [...model.unmodelled, ...model.tableGaps];
};

/** What resolving every change must leave, and whether block boundaries are known. */
export type Resolution = { rows: Row[]; boundaries: boolean };

/** Apply `operations` as one core batch in `mode`, and check the result is what they asked. */
export const applyChecked = async (
  reviewer: Reviewer,
  operations: readonly Operation[],
  mode: Mode,
  context: string,
): Promise<{ applied: string[]; issues: string[] }> => {
  const pre = await capture(reviewer, mode);
  const batch = coreBatch(operations, mode);
  const result = reviewer.applyDocumentOperations(batch as never, {
    undefinedReferences: "refuse",
  });
  const applied = new Set(result.applied.map(({ id }) => id));
  const issues = result.issues.map(
    (issue) => `${issue.operationId}: ${issue.code}`,
  );
  await assertRequestedOutcome(
    reviewer,
    pre,
    {
      applied: batch.operations.filter((operation) =>
        applied.has(operation.id),
      ),
    },
    `${context} ${JSON.stringify(operations)} (refused: ${issues.join(", ") || "none"})`,
  );
  return { applied: [...applied], issues };
};

/**
 * What bulk resolution must leave: the accepted (or rejected) view of the
 * saved package. The package excludes pending suggestions, which remain
 * staged in memory. A paragraph-mark deletion can join blocks on acceptance,
 * so the suggested live view may not predict their boundaries.
 */
export const captureResolution = async (
  reviewer: Reviewer,
  _mode: Mode,
  resolution: "accept" | "reject",
): Promise<Resolution> => ({
  rows: (await resolvedState(await save(reviewer), resolution)).rows,
  boundaries: true,
});

/** After `acceptAll()` / `rejectAll()`: the saved package reads `expected`. */
export const assertResolvedTo = async (
  reviewer: Reviewer,
  expected: Resolution,
  context: string,
): Promise<void> => {
  const rows = rowsOf(await openReviewer(await save(reviewer)));
  const message = `${context}: resolving every change does not leave what the reviewer showed`;
  if (expected.boundaries) {
    assert.deepEqual(fullProject(rows), fullProject(expected.rows), message);
  } else {
    assert.deepEqual(
      authoredInlineProject(rows),
      authoredInlineProject(expected.rows),
      message,
    );
    if (
      JSON.stringify(visible(rows).map((row) => row.text)) ===
      JSON.stringify(visible(expected.rows).map((row) => row.text))
    ) {
      assert.deepEqual(
        suggestedProject(rows),
        suggestedProject(expected.rows),
        message,
      );
    }
  }
};
