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

import { sourceDocumentOf } from "@stll/folio-core/ai-edits/snapshot";
import type { FolioContentStatedNumbering } from "@stll/folio-core/compare/content-types";
import { expectParagraphAttrs } from "@stll/folio-core/prosemirror/attrs";
import {
  FOLIO_DOCUMENT_OPERATION_TYPES,
  type FolioDocumentOperationType,
  type FolioDocumentStoryHandle,
  getFolioDocumentOperationRemovedBlockIds,
  getTrackedFolioDocumentOperationRemovedBlockIds,
  inspectDocumentStylesFromDocx,
  parseDocx,
} from "@stll/folio-core/server";

import { blanksGoingWhole, sequentialGroups } from "./batch-replay-order.ts";
import { recordFeatureHit, recordHit, type StepKind } from "./coverage.ts";
import { openReviewer, scenarioEngine, toArrayBuffer } from "./documents.ts";
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
  captureSourceOwners,
  compareTableGeometry,
  modelFromRows,
  removeTextBoxTables,
  sourceOperationFootprint,
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

type LiveNumberingLevel = NonNullable<
  ReturnType<Reviewer["readNumberingDefinitions"]>
>[number];

/**
 * The marker levels of a story's lists that only its paragraphs
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

/** Capture the authored story, including proposals omitted from the saved model. */
const liveNumberingFacts = (
  reviewer: Reviewer,
  story: Story,
): NumberingFacts => {
  const facts = numberingFactsFromDocument(reviewer.toDocument());
  const snapshot = reviewer.snapshotStory(story);
  if (!snapshot) return facts;
  const { anchors, blocks } = snapshot;
  // A story reads the lists it started itself; no other story defines them.
  const definitions = reviewer.readNumberingDefinitions(story);
  assert.ok(definitions, "A snapshotted story must have numbering to read");
  addLiveListMarkerLevels(facts, definitions);
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
        case "commentRangeStart":
        case "commentRangeEnd":
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
  captureSourceOwners({ document: reviewer.toDocument(), rows });
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

type NoteLifetime =
  | { type: "not-observed" }
  | {
      type: "note";
      story: NoteStory;
      accepted: StoryPresence;
      rejected: StoryPresence;
      commentIds: ReadonlySet<number>;
    };

type ResolvedState = {
  rows: Row[];
  story: StoryPresence;
  bytes: Uint8Array;
  comments: Comment[];
  commentThreads: CommentThread[];
  sourceCommentParents: CommentParent[];
  commentParents: CommentParent[];
  links: LinkSnapshot;
};

type ResolutionScope =
  { type: "document"; story: Story } | { type: "story"; story: Story };

/** Instructions in a secondary story cannot resolve changes to its owner. */
export const intentResolutionScope = (story: Story): ResolutionScope =>
  story.type === "main"
    ? { type: "document", story }
    : { type: "story", story };

type ResolvedStateOptions = {
  bytes: Uint8Array;
  resolution: "accept" | "reject";
  scope: ResolutionScope;
};

/**
 * Resolve the explicitly named scope, then save and reopen its read story.
 * A story projection keeps pending changes to references in other stories.
 */
export const resolvedState = async ({
  bytes,
  resolution,
  scope,
}: ResolvedStateOptions): Promise<ResolvedState> => {
  const reviewer = await openReviewer(bytes);
  const { story } = scope;
  const observesNote = story.type === "footnote" || story.type === "endnote";
  const sourceCommentParents = observesNote ? commentParentsOf(reviewer) : [];
  switch (scope.type) {
    case "document":
      if (resolution === "accept") reviewer.acceptAll();
      else reviewer.rejectAll();
      break;
    case "story": {
      assert.equal(
        reviewer.resolveReviewedStory({
          story,
          view: resolution === "accept" ? "final" : "original",
        }),
        true,
        `The ${resolution} projection did not resolve ${JSON.stringify(story)}`,
      );
      const reviewed = reviewer.readReviewedStory({
        story,
        view: "current-markup",
      });
      assert.ok(reviewed, "A resolved intent story must remain readable");
      assert.equal(
        reviewed.changes.length,
        0,
        "A story projection left unresolved changes",
      );
      break;
    }
    default: {
      const unhandled: never = scope;
      assert.fail(`Unknown resolution scope ${JSON.stringify(unhandled)}`);
    }
  }
  const saved = await save(reviewer);
  const reopened = await openReviewer(saved);
  return {
    rows: rowsOf(reopened, story),
    story: reopened.readStory(story) === null ? "removed" : "present",
    bytes: saved,
    comments: commentsOf(reopened),
    commentThreads: observesNote ? commentThreadsOf(reopened) : [],
    sourceCommentParents,
    commentParents: observesNote ? commentParentsOf(reopened) : [],
    links: captureLinks(reopened),
  };
};

type AssertNoteLifetimeOptions = {
  before: NoteLifetime;
  bytes: Uint8Array;
  accepted: ResolvedState;
  rejected: ResolvedState | undefined;
};

/** Note edits preserve reference fate, and every surviving global note keeps its content. */
const assertNoteLifetime = async ({
  before,
  bytes,
  accepted,
  rejected,
}: AssertNoteLifetimeOptions): Promise<void> => {
  if (before.type === "not-observed") return;
  const { story } = before;
  const localRejected =
    rejected ??
    (await resolvedState({
      bytes,
      resolution: "reject",
      scope: { type: "story", story },
    }));
  const decisions = [
    { resolution: "accept", expected: before.accepted, local: accepted },
    { resolution: "reject", expected: before.rejected, local: localRejected },
  ] as const;
  const noteCommentIds = new Set(before.commentIds);
  for (const state of [accepted, localRejected]) {
    const ids = new Set(state.rows.map(({ id }) => id));
    for (const comment of state.comments)
      if (comment.blockId !== null && ids.has(comment.blockId))
        noteCommentIds.add(comment.id);
  }
  closeCommentReplies(noteCommentIds, accepted.sourceCommentParents);
  closeCommentReplies(noteCommentIds, localRejected.sourceCommentParents);
  for (const { resolution, expected, local } of decisions) {
    const global = await resolvedState({
      bytes,
      resolution,
      scope: { type: "document", story },
    });
    assert.equal(
      global.story,
      expected,
      `A note-only edit changed its globally ${resolution}ed reference fate`,
    );
    const localIds = new Set(local.rows.map(({ id }) => id));
    const localComments = local.commentThreads.filter(
      ({ blockId }) => blockId !== null && localIds.has(blockId),
    );
    const globalIds = new Set(global.rows.map(({ id }) => id));
    const globalComments = global.commentThreads.filter(
      ({ id, blockId }) =>
        noteCommentIds.has(id) || (blockId !== null && globalIds.has(blockId)),
    );
    if (global.story === "removed") {
      assert.deepEqual(
        global.rows,
        [],
        "A globally removed note retained blocks",
      );
      assert.deepEqual(
        globalComments,
        [],
        "A globally removed note retained comments",
      );
      assert.deepEqual(
        global.commentParents.filter(({ id }) => noteCommentIds.has(id)),
        [],
        "A globally removed note retained comment definitions or replies",
      );
      continue;
    }
    // Package-minted paragraph ids need not survive the two save projections.
    const content = (rows: readonly Row[]) =>
      rows.map(({ id: _id, ...row }) => row);
    assert.deepEqual(
      content(global.rows),
      content(local.rows),
      `Global ${resolution} changed surviving note content`,
    );
    const comments = (
      state: ResolvedState,
      entries: readonly CommentThread[],
    ) => {
      const positions = new Map(state.rows.map(({ id }, index) => [id, index]));
      return entries.map(({ blockId, ...entry }) => {
        const paragraph = blockId === null ? null : positions.get(blockId);
        assert.notEqual(
          paragraph,
          undefined,
          "A note comment has no containing paragraph",
        );
        return { ...entry, paragraph };
      });
    };
    assert.deepEqual(
      comments(global, globalComments),
      comments(local, localComments),
      `Global ${resolution} changed surviving note comments`,
    );
  }
};

type CommentParent = { id: number; parentId: number | null };

const commentParentsOf = (reviewer: Reviewer): CommentParent[] =>
  (reviewer.toDocument().package.document.comments ?? []).map(
    ({ id, parentId }) => ({
      id,
      parentId: parentId ?? null,
    }),
  );

/** Include replies even when a broken projection leaves them without a visible root. */
const closeCommentReplies = (
  ids: Set<number>,
  parents: readonly CommentParent[],
): void => {
  const children = new Map<number, number[]>();
  for (const { id, parentId } of parents) {
    if (parentId === null) continue;
    const siblings = children.get(parentId) ?? [];
    siblings.push(id);
    children.set(parentId, siblings);
  }
  const pending = [...ids];
  for (let index = 0; index < pending.length; index += 1) {
    const parent = pending[index];
    if (parent === undefined) assert.fail("A queued comment parent must exist");
    for (const child of children.get(parent) ?? []) {
      if (ids.has(child)) continue;
      ids.add(child);
      pending.push(child);
    }
  }
};

type CommentThread = Comment & { replies: { id: number; text: string }[] };
const commentThreadsOf = (reviewer: Reviewer): CommentThread[] =>
  reviewer.getComments().map((comment) => ({
    id: comment.id,
    text: comment.text,
    anchor: comment.anchoredText ?? "",
    blockId: comment.blockId ?? null,
    replies: comment.replies.map(({ id, text }) => ({ id, text })),
  }));

type Comment = {
  id: number;
  text: string;
  anchor: string;
  blockId: string | null;
};
type ComparedComment = Comment & { finalAnchor?: string };

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
  /** Suggested mode: the properties a paragraph shows while its join is pending. */
  ownWhileJoinPending?: { fields: Fields; listLineage: ListLineage };
  removed: boolean;
  /**
   * The block ends in a paragraph mark pending deletion: a reader lists it
   * and the block after it apart, accepting joins them. Text operations use
   * verified slices of the joined text; structural operations stay unmodelled.
   */
  pendingJoin?: boolean;
  /**
   * A whole paragraph pending deletion, which a reader lists as a blank
   * block: accepting removes it, mark and all, unless an operation wrote
   * into it.
   */
  pendingDeletion?: boolean;
  /**
   * A paragraph pending deletion that holds nothing at all: its join runs
   * into the paragraph after it, so a table inserted after it goes after
   * that paragraph when it keeps words.
   */
  emptyJoinSource?: boolean;
  /**
   * A block the reviewer still lists but the resolved pre-state does not
   * hold: nothing read from that pre-state describes it.
   */
  outsidePreState?: boolean;
  /**
   * Such a block whose words accepting drops with no pending join to place
   * them (a table row pending deletion, say): removed before the batch ran.
   */
  wordsAcceptingDrops?: true;
  /**
   * A block inserted before or after the table its anchor sits in. Rows and
   * cells the batch adds to that table go inside it, so they stay between
   * the table's own blocks and this one, whatever the batch order.
   */
  besideTable?: boolean;
  /** A block of a table row the batch deletes, rather than one emptied alone. */
  rowDeleted?: boolean;
  /** A cell paragraph of a table the batch inserts: no join runs into it. */
  inInsertedTable?: boolean;
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
  /** The engine the result comes from. */
  engine: ReturnType<typeof scenarioEngine>;
  rows: ModelRow[];
  /** Live source paragraphs whose pending marks run into an accepted row. */
  pendingJoins: Map<
    ModelRow,
    {
      sources: readonly Row[];
      destination: Row;
      /** Live sources still proved to reach the destination during this batch. */
      deletionSources: readonly { row: Row; from: number }[] | undefined;
      /** Source coordinates remain known when an applied effect defeats destination prediction. */
      destinationPrediction:
        { type: "exact" } | { type: "unknown"; operationType: string };
    }
  >;
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
  /** Live rows retain the capture's source ancestry facts for applied-effect footprints. */
  sourceRows: readonly Row[];
  /** Live blocks explicitly deleted by applied operations, including pending joins. */
  deletedBlockIds: Set<string>;
  /** Source text consumed by a canonical tracked deletion of its join destination. */
  deletedJoinRuns: Map<
    string,
    { destination: string; from: number; text: string }
  >;
  /** Live blocks whose whole text an applied replaceBlock replaced, with the new text. */
  replacedBlocks: Map<string, string>;
  /** Live paragraphs whose every word is a pending deletion (`JoinSources`). */
  deletedContentIds: ReadonlySet<string>;
  /** Paragraphs of table rows pending insertion (`JoinSources`). */
  insertedRowIds: ReadonlySet<string>;
  /** Paragraphs a pending insertion added after another (`JoinSources`). */
  addedAfterIds: ReadonlySet<string>;
  /** Live paragraphs whose mark accepting removes (`JoinSources`). */
  joinSourceIds: ReadonlySet<string>;
  /** Live paragraphs whose mark is pending insertion (`JoinSources`). */
  markInsertedIds: ReadonlySet<string>;
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
    /** Whether the batch leaves this comment's survival unpredicted. */
    ambiguous?: () => boolean;
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
 * drops words the model does not place. `emptyIds` are those holding nothing
 * at all, not even deleted words: a pending join starts from them, while a
 * paragraph with only deleted words goes whole. `assumed`: every one of them
 * (for hand-built rows with no package to read). `live`: the suggested model,
 * which reads the live rows themselves, so no block is missing from them and
 * none joins another; `markDeletedIds` are the paragraphs whose mark is a
 * pending deletion, `markInsertedIds` those absent from the original stream,
 * and `goWhole` the blank ones accepting removes whole.
 */
export type JoinSources =
  | {
      type: "marked";
      ids: ReadonlySet<string>;
      emptyIds: ReadonlySet<string>;
      /**
       * Paragraphs whose every word is a pending deletion, whatever their
       * mark: a merge into one that ends its container has nothing to put a
       * separator in front of.
       */
      deletedContentIds?: ReadonlySet<string>;
      /**
       * Paragraphs of table rows pending insertion: a join through a deleted
       * paragraph does not run into a table under review.
       */
      insertedRowIds?: ReadonlySet<string>;
      /**
       * Paragraphs a pending insertion by the capturing author added after
       * the paragraph before them, taking over its mark (that mark is
       * inserted with `survivor: "first"`), and holding only that author's
       * inserted words: that author removing one under tracking retracts
       * the insertion, so the paragraph before keeps its own mark and
       * identity. Another author's deletion of it is an ordinary one.
       */
      addedAfterIds?: ReadonlySet<string>;
    }
  | {
      type: "live";
      markDeletedIds: ReadonlySet<string>;
      markInsertedIds: ReadonlySet<string>;
      goWhole: ReadonlySet<string>;
    }
  | { type: "assumed" };

const ASSUMED_JOIN_SOURCES: JoinSources = { type: "assumed" };

/** The live paragraphs whose mark accepting removes. */
const joinSourceIdsOf = (sources: JoinSources): ReadonlySet<string> => {
  switch (sources.type) {
    case "marked":
      return sources.ids;
    case "live":
      return sources.markDeletedIds;
    case "assumed":
      return new Set();
    default: {
      const unhandled: never = sources;
      throw new Error(`Unhandled join sources ${JSON.stringify(unhandled)}`);
    }
  }
};

const joinsOnAccept = (sources: JoinSources, id: string): boolean => {
  switch (sources.type) {
    case "marked":
      return sources.ids.has(id);
    case "live":
      return false;
    case "assumed":
      return true;
    default: {
      const unhandled: never = sources;
      throw new Error(`Unhandled join sources ${JSON.stringify(unhandled)}`);
    }
  }
};

/** Paragraph content that holds no words: markers a deletion leaves standing. */
const ZERO_WIDTH_INLINES: ReadonlySet<string> = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

/** Paragraph marks accepting removes: a pending deletion, or a move's source. */
const JOINING_MARK_KINDS: ReadonlySet<string> = new Set(["del", "moveFrom"]);

/** The paraIds of the paragraphs of `bytes` whose mark accepting removes. */
/** Whether `inline` is that author's pending insertion, or a zero-width anchor. */
const insertedBy = (inline: unknown, author: string): boolean => {
  if (typeof inline !== "object" || inline === null) return false;
  const type: unknown = Reflect.get(inline, "type");
  if (typeof type === "string" && ZERO_WIDTH_INLINES.has(type)) return true;
  if (type !== "insertion") return false;
  const info: unknown = Reflect.get(inline, "info");
  return (
    typeof info === "object" &&
    info !== null &&
    Reflect.get(info, "author") === author
  );
};

const markedJoinSourcesOf = async (
  bytes: Uint8Array,
  author?: string,
): Promise<JoinSources> => {
  const ids = new Set<string>();
  const emptyIds = new Set<string>();
  const deletedContentIds = new Set<string>();
  const insertedRowIds = new Set<string>();
  const addedAfterIds = new Set<string>();
  const parsed = await parseDocx(toArrayBuffer(bytes), { preloadFonts: false });
  const visitBlocks = (blocks: readonly unknown[], inserted: boolean): void => {
    let previousMark: unknown;
    for (const block of blocks) {
      if (typeof block !== "object" || block === null) continue;
      const type: unknown = Reflect.get(block, "type");
      if (type === "paragraph") {
        const id: unknown = Reflect.get(block, "paraId");
        if (inserted && typeof id === "string") insertedRowIds.add(id);
        const content: unknown = Reflect.get(block, "content");
        if (
          author !== undefined &&
          typeof id === "string" &&
          Array.isArray(content) &&
          content.length > 0 &&
          content.every((inline) => insertedBy(inline, author)) &&
          typeof previousMark === "object" &&
          previousMark !== null &&
          Reflect.get(previousMark, "kind") === "ins" &&
          Reflect.get(previousMark, "survivor") === "first"
        )
          addedAfterIds.add(id);
        previousMark = Reflect.get(block, "pPrMark");
        continue;
      }
      previousMark = undefined;
      if (type !== "table") continue;
      const rows: unknown = Reflect.get(block, "rows");
      for (const row of Array.isArray(rows) ? rows : []) {
        const change: unknown = Reflect.get(row, "structuralChange");
        const rowInserted =
          inserted ||
          (typeof change === "object" &&
            change !== null &&
            Reflect.get(change, "type") === "tableRowInsertion");
        const cells: unknown = Reflect.get(row, "cells");
        for (const cell of Array.isArray(cells) ? cells : []) {
          const content: unknown = Reflect.get(cell, "content");
          visitBlocks(Array.isArray(content) ? content : [], rowInserted);
        }
      }
    }
  };
  visitBlocks(parsed.package.document.content, false);
  visitParsedParagraphs(parsed, (paragraph) => {
    if (
      paragraph.paraId !== undefined &&
      paragraph.content.some(({ type }) => type === "deletion") &&
      paragraph.content.every(
        ({ type }) => type === "deletion" || ZERO_WIDTH_INLINES.has(type),
      )
    )
      deletedContentIds.add(paragraph.paraId);
    if (
      paragraph.paraId === undefined ||
      paragraph.pPrMark === undefined ||
      !JOINING_MARK_KINDS.has(paragraph.pPrMark.kind)
    )
      return;
    ids.add(paragraph.paraId);
    if (paragraph.content.length === 0) emptyIds.add(paragraph.paraId);
  });
  return {
    type: "marked",
    ids,
    emptyIds,
    deletedContentIds,
    insertedRowIds,
    addedAfterIds,
  };
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
  inTextBox: ReadonlySet<string> = new Set(),
): Model => {
  const model: Model = {
    engine: scenarioEngine(),
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
    sourceRows: live,
    deletedBlockIds: new Set(),
    deletedJoinRuns: new Map(),
    replacedBlocks: new Map(),
    deletedContentIds:
      joinSources.type === "marked"
        ? (joinSources.deletedContentIds ?? new Set())
        : new Set(),
    insertedRowIds:
      joinSources.type === "marked"
        ? (joinSources.insertedRowIds ?? new Set())
        : new Set(),
    addedAfterIds:
      joinSources.type === "marked"
        ? (joinSources.addedAfterIds ?? new Set())
        : new Set(),
    joinSourceIds: joinSourceIdsOf(joinSources),
    markInsertedIds:
      joinSources.type === "live" ? joinSources.markInsertedIds : new Set(),
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
  for (const [liveIndex, row] of live.entries()) {
    // A text box's paragraphs are listed right after the paragraph it is
    // drawn in, in a story of their own: words joining on run past them into
    // the next paragraph of the joining one's story.
    if (joining && inTextBox.has(row.id)) continue;
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
          const sources = joiningSources;
          model.pendingJoins.set(survivor, {
            sources,
            destination: row,
            deletionSources:
              joinSources.type === "marked" &&
              !joinSources.ids.has(row.id) &&
              sources.every(
                (source, index) =>
                  live.at(liveIndex - sources.length + index) === source &&
                  sameParagraphContainer(source, row),
              ) &&
              survivor.pre.text ===
                [...sources, row].map(({ text }) => text).join("")
                ? sources.map((source) => ({ row: source, from: 0 }))
                : undefined,
            destinationPrediction: { type: "exact" },
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
      ghost.wordsAcceptingDrops = true;
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
    if (joinSources.type === "marked" && joinSources.emptyIds.has(row.id))
      ghost.emptyJoinSource = true;
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
  if (joinSources.type === "live")
    for (const row of model.rows)
      if (row.pre !== undefined && joinSources.goWhole.has(row.pre.id))
        row.pendingDeletion = true;
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
  // Accepting already drops it, before the batch: an edit there is not one
  // the batch makes on a block it removed.
  if (row.wordsAcceptingDrops && !adjacent)
    throw new Unmodelled(
      `block ${String(blockId)} has words accepting drops with no pending paragraph-mark deletion`,
    );
  // An insertion only places a block beside its anchor, which may go.
  if (row.removed && !adjacent) {
    throw new Error(
      `applied an operation on block ${String(blockId)}, which an earlier one of the batch removed`,
    );
  }
  return row;
};

/** A live paragraph's proven slice of the accepted pre-state. */
const paragraphTextProjection = (model: Model, blockId: unknown) => {
  const original = model.rows.find(
    (candidate) => candidate.pre?.id === blockId,
  );
  if (original?.pendingJoin !== true)
    return original === undefined
      ? ({ type: "unresolved" } as const)
      : ({
          type: "ordinary",
          row: original,
          text: original.pre?.text ?? original.text,
          offset: 0,
        } as const);
  const candidates = [...model.pendingJoins].flatMap(([row, ownership]) => {
    const pieces = [...ownership.sources, ownership.destination];
    if (!pieces.some(({ id }) => id === blockId)) return [];
    const ids = new Set(pieces.map(({ id }) => id));
    if (
      ids.size !== pieces.length ||
      row.pre?.text !== pieces.map(({ text }) => text).join("") ||
      !pieces.every((piece) =>
        sameParagraphContainer(piece, ownership.destination),
      )
    )
      return [];
    let offset = 0;
    for (const piece of pieces) {
      if (piece.id === blockId)
        return [
          {
            row,
            text: piece.text,
            offset,
            destinationPrediction: ownership.destinationPrediction,
          },
        ];
      offset += piece.text.length;
    }
    return [];
  });
  const found = candidates.at(0);
  return candidates.length === 1 && found !== undefined
    ? ({
        type: "pending-join",
        row: found.row,
        text: found.text,
        offset: found.offset,
        destinationPrediction: found.destinationPrediction,
      } as const)
    : ({ type: "unresolved" } as const);
};

/** Text operations keep live coordinates; structural targets retain their existing guard. */
const textTarget = (model: Model, blockId: unknown) => {
  const view = paragraphTextProjection(model, blockId);
  switch (view.type) {
    case "ordinary":
      target(model, blockId);
      return view;
    case "pending-join":
      assert.ok(
        !view.row.removed &&
          !(typeof blockId === "string" && model.deletedBlockIds.has(blockId)),
        `applied an operation on block ${String(blockId)}, which an earlier one of the batch removed`,
      );
      return view;
    case "unresolved": {
      const row = model.rows.find((candidate) => candidate.pre?.id === blockId);
      if (row?.pendingJoin !== true) target(model, blockId);
      throw new Unmodelled(
        `block ${String(blockId)} has a pending join with no exact live-to-accepted text projection`,
      );
    }
    default: {
      const unhandled: never = view;
      return assert.fail(
        `Unknown text projection ${JSON.stringify(unhandled)}`,
      );
    }
  }
};

type ProjectedTextRangeOptions = {
  model: Model;
  blockId: unknown;
  start: number;
  end: number;
};
const projectedTextRange = ({
  model,
  blockId,
  start,
  end,
}: ProjectedTextRangeOptions) => {
  const view = textTarget(model, blockId);
  assert.ok(
    start >= 0 && start <= end && end <= view.text.length,
    "An applied range must fit its live paragraph",
  );
  return { row: view.row, start: view.offset + start, end: view.offset + end };
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

/**
 * Whether `later`, a block after `row`, is in `row`'s story or table cell:
 * a paragraph there, or a cell of a table there. A story paragraph before a
 * table does not end the story; the table follows it there.
 */
const followsInContainer = (row: Row, later: Row): boolean =>
  sameParagraphContainer(row, later) ||
  (row.table === undefined && later.table !== undefined);

/** Whether block `id` is in `rows` and the last of its story or table cell there. */
const endsContainer = (rows: readonly Row[], id: string): boolean => {
  const index = rows.findIndex((row) => row.id === id);
  const row = rows[index];
  return (
    row !== undefined &&
    rows.slice(index + 1).every((later) => !followsInContainer(row, later))
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
  return model.rows.slice(index + 1).every(
    (candidate) =>
      candidate.pre === undefined ||
      // A live-only row accepting removes (a join's source, which the
      // model lists after every row) is not in the accepted container.
      (candidate.outsidePreState === true && candidate.removed) ||
      !followsInContainer(pre, candidate.pre),
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
  ambiguous?: () => boolean,
): void => {
  const comment = operation["comment"] as { text?: string } | undefined;
  if (typeof comment?.text === "string") {
    model.comments.push({
      text: comment.text,
      ...(anchor === undefined ? {} : { anchor }),
      ...(removed === undefined ? {} : { removed }),
      ...(ambiguous === undefined ? {} : { ambiguous }),
    });
  }
};

/** Words and marks, the finest pieces a word diff keeps unchanged. */
const diffPieces = (text: string): Set<string> =>
  new Set(text.match(/[\p{L}\p{N}_]+|[^\s\p{L}\p{N}_]/gu) ?? []);

const WORD_CHARACTER = /[\p{L}\p{N}\p{M}\p{Pc}]/u;
const ONE_WORD =
  /^[^\p{L}\p{N}\p{M}\p{Pc}]*[\p{L}\p{N}\p{M}\p{Pc}]*[^\p{L}\p{N}\p{M}\p{Pc}]*$/u;

const sharedLength = (left: string[], right: string[]): number => {
  let length = 0;
  while (
    length < Math.min(left.length, right.length) &&
    left[length] === right[length]
  )
    length++;
  return length;
};

/** Whether `characters[0, index)` ends on a word boundary. */
const endsWord = (characters: readonly string[], index: number): boolean =>
  index <= 0 ||
  index >= characters.length ||
  !WORD_CHARACTER.test(characters[index - 1] ?? "") ||
  !WORD_CHARACTER.test(characters[index] ?? "");

/** The longest shared length up to `length` ending on a word boundary of both. */
const wordAligned = (
  left: readonly string[],
  right: readonly string[],
  length: number,
): number => {
  for (let candidate = length; candidate > 0; candidate--)
    if (endsWord(left, candidate) && endsWord(right, candidate))
      return candidate;
  return 0;
};

type KeptEdges = { prefix: number; suffix: number };

/**
 * How many characters a direct replacement of `removed` by `inserted` that
 * share no word or mark keeps at each end, in UTF-16 units: the smallest
 * change keeps shared whole words first and, once what is left is one word
 * on both sides, shared letters ("defined" to "updated" keeps "ed").
 */
const smallestChangeKeeps = (removed: string, inserted: string): KeptEdges => {
  let old = [...removed];
  let added = [...inserted];
  const kept = { prefix: 0, suffix: 0 };
  for (const granularity of ["word", "letter"] as const) {
    const oneWord =
      ONE_WORD.test(old.join("")) && ONE_WORD.test(added.join(""));
    if (granularity === "letter" && !oneWord) break;
    const shared = sharedLength(old, added);
    const prefix = oneWord ? shared : wordAligned(old, added, shared);
    kept.prefix += old.slice(0, prefix).join("").length;
    old = old.slice(prefix);
    added = added.slice(prefix);
    const reversed = sharedLength(old.toReversed(), added.toReversed());
    const suffix = oneWord
      ? reversed
      : wordAligned(old.toReversed(), added.toReversed(), reversed);
    kept.suffix += old.slice(old.length - suffix).join("").length;
    old = old.slice(0, old.length - suffix);
    added = added.slice(0, added.length - suffix);
    if (oneWord) break;
  }
  return kept;
};

/**
 * What becomes of a comment over `text[start, end)` once the edits apply.
 * A comment goes with its words: one every character of which the edits
 * remove retires, and one with a character left keeps a range over it. The
 * words that replace a range go past the end of a comment the range
 * reaches. A direct edit, and a tracked one meeting a comment once accepted,
 * writes the smallest change, which keeps the characters the old and new
 * text share at its ends (`smallestChangeKeeps`), and the comment on them. A
 * replacement sharing a word or mark with the covered text may keep that
 * piece through a word diff, and the comment on it: neither is predicted.
 * Neither is it on an engine that hands the comment to the new words.
 */
const coveredRangeFate = (
  engine: Model["engine"],
  edits: readonly TextEdit[],
  text: string,
  start: number,
  end: number,
): "removed" | "kept" | "ambiguous" => {
  if (start >= end) return "kept";
  const covering = new Set<TextEdit>();
  for (let offset = start; offset < end; offset++) {
    const edit = edits.find(
      (candidate) => candidate.start <= offset && offset < candidate.end,
    );
    if (edit === undefined) return "kept";
    covering.add(edit);
  }
  let fate: "removed" | "kept" = "removed";
  for (const edit of covering) {
    if (edit.replace === "") continue;
    if (engine !== "canonical") return "ambiguous";
    const kept = diffPieces(edit.replace);
    const covered = text.slice(
      Math.max(start, edit.start),
      Math.min(end, edit.end),
    );
    if ([...diffPieces(covered)].some((piece) => kept.has(piece)))
      return "ambiguous";
    const { prefix, suffix } = smallestChangeKeeps(
      text.slice(edit.start, edit.end),
      edit.replace,
    );
    if (
      (prefix > 0 && start < edit.start + prefix) ||
      (suffix > 0 && edit.end - suffix < end)
    )
      fate = "kept";
  }
  return fate;
};

const coveredRangeWasRemoved = (
  model: Model,
  row: ModelRow,
  start: number,
  end: number,
): boolean =>
  coveredRangeFate(
    model.engine,
    row.edits,
    row.pre?.text ?? row.text,
    start,
    end,
  ) === "removed";

/** Whether the edits leave a comment over `text[start, end)` unpredicted. */
const coveredRangeIsAmbiguous = (
  model: Model,
  row: ModelRow,
  start: number,
  end: number,
): boolean =>
  coveredRangeFate(
    model.engine,
    row.edits,
    row.pre?.text ?? row.text,
    start,
    end,
  ) === "ambiguous";

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

type Container = TableLocation | undefined;

const sameContainer = (left: Container, right: Container): boolean =>
  left === undefined
    ? right === undefined
    : right !== undefined &&
      left.tableIndex === right.tableIndex &&
      left.rowIndex === right.rowIndex &&
      left.cellIndex === right.cellIndex;

/** The block after `row` that the resolved pre-state holds, in its container. */
const nextInContainer = (model: Model, row: ModelRow): ModelRow | undefined => {
  const next = model.rows
    .slice(model.rows.indexOf(row) + 1)
    .find((candidate) => candidate.outsidePreState !== true);
  return next !== undefined && sameContainer(row.pre?.table, next.pre?.table)
    ? next
    : undefined;
};

/**
 * The paragraph a table inserted after `row` follows: the one a pending join
 * from an empty `row` leaves, when that paragraph keeps words of its own;
 * `row` itself otherwise. Other blocks the reader lists but accepting removes go
 * with the join on the way there.
 */
const pastEmptyJoin = (model: Model, row: ModelRow): ModelRow => {
  if (row.emptyJoinSource !== true) return row;
  const next = nextInContainer(model, row);
  return next !== undefined && next.text.length > 0 ? next : row;
};

/**
 * The paragraph a paragraph inserted after `row` follows: the one the
 * pending join from `row` leaves, when `row` keeps no words once accepted and
 * reads as one paragraph with it (it is empty, or goes whole with words
 * joining on through it) and that one keeps words; `row` itself otherwise.
 */
const pastWordlessJoin = (model: Model, row: ModelRow): ModelRow => {
  if (
    row.emptyJoinSource !== true &&
    !(row.pendingDeletion === true && row.pendingJoin === true)
  )
    return row;
  const next = nextInContainer(model, row);
  return next !== undefined && next.text.length > 0 ? next : row;
};

/**
 * The paragraph a paragraph inserted before `row` copies: the one the
 * pending join from `row` leaves, when `row` holds nothing at all and its
 * break is pending deletion. Put before that emptied paragraph, it reads as
 * the joined paragraph, which has that paragraph's properties (the engine's
 * `besidePendingJoin`, canonicalInsertBesidePendingJoin.test.ts), so it
 * takes no numbering from the emptied list item. Before a paragraph still
 * holding its deleted words, it copies that paragraph; `row` itself.
 */
const pastEmptiedJoinBefore = (model: Model, row: ModelRow): ModelRow => {
  if (row.emptyJoinSource !== true) return row;
  return nextInContainer(model, row) ?? row;
};

/**
 * The paragraph a paragraph inserted `position` of `named` copies: put
 * after the paragraph a join leaves, that paragraph; put before an emptied
 * paragraph a pending join runs on, the paragraph the join leaves. A
 * suggestion is read against the live
 * rows, which hold no join.
 */
const copiedParagraph = (
  model: Model,
  named: ModelRow,
  position: "before" | "after",
): ModelRow => {
  if (model.mode === "suggested") return named;
  return position === "after"
    ? pastWordlessJoin(model, named)
    : pastEmptiedJoinBefore(model, named);
};

/**
 * Where a block inserted next to `row` goes: next to the row itself, or, for
 * a row inside a table, next to the outermost table (the documented rule:
 * an insertion anchored in a cell lands beside the table, not in the cell).
 */
const insertionAnchor = (
  model: Model,
  row: ModelRow,
  position: "before" | "after",
  block: "paragraph" | "table" = "table",
): ModelRow => {
  // A paragraph put after one that keeps no words once accepted, whose break
  // a pending join removes, follows the paragraph the join leaves, with that
  // one's properties, as the same edits applied directly put it (the
  // engine's `insertionAfterPendingBreak`); a suggested one, read against
  // live rows that hold no join, goes right after the paragraph it names
  // (`insertBlock` leaves one beside a pending join unmodelled). A table follows the paragraph a
  // join from an empty paragraph leaves.
  if (position === "after") {
    let joined = pastWordlessJoin(model, row);
    if (block === "table") joined = pastEmptyJoin(model, row);
    else if (model.mode === "suggested") joined = row;
    row = pastTextBoxes(model, joined);
  }
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
    // Direct numbering outlives the restyle, and so does an outline level the
    // paragraph states itself, but a heading style still makes the block a
    // heading. A heading only its old style made is a list item or a plain
    // paragraph once restyled, as its direct numbering shows.
    if (result.fields.kind !== "heading") {
      const headingFromOldStyle =
        live.kind === "heading" && live.directOutlineLevel?.kind !== "heading";
      const directKind =
        live.listLevel === undefined ? "paragraph" : "listItem";
      result.fields.kind = headingFromOldStyle ? directKind : live.kind;
    }
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

/**
 * Whether a suggestion put before `named` lands before a pending join rather
 * than beside it: `named` or the paragraph a reader lists before it, past
 * blank ones, has its mark pending deletion. It lands where the tracked
 * insertion does, before the join, with the joined paragraph's properties,
 * which the live rows a suggestion is read against do not say.
 */
const suggestedBeforePendingJoin = (model: Model, named: ModelRow): boolean => {
  if (model.mode !== "suggested" || named.pre === undefined) return false;
  if (model.joinSourceIds.has(named.pre.id)) return true;
  const previous = model.rows
    .slice(0, model.rows.indexOf(named))
    .findLast((row) => row.text.length > 0);
  return (
    previous?.pre !== undefined && model.joinSourceIds.has(previous.pre.id)
  );
};

const insertBlock =
  (position: "before" | "after"): Expect =>
  (model, operation) => {
    const named = target(model, operation["blockId"], { adjacent: true });
    if (position === "before" && suggestedBeforePendingJoin(model, named))
      throw new Unmodelled(
        `a suggestion before block ${String(operation["blockId"])} lands before the pending join it is part of`,
      );
    // Put after a paragraph whose mark is pending deletion, a suggestion
    // lands where the tracked insertion does (`insertionAfterPendingBreak`),
    // which the live rows a suggestion is read against do not say.
    if (
      position === "after" &&
      model.mode === "suggested" &&
      named.pre !== undefined &&
      model.joinSourceIds.has(named.pre.id)
    )
      throw new Unmodelled(
        `a suggestion after block ${String(operation["blockId"])} lands where its pending join puts it`,
      );
    const anchor = insertionAnchor(model, named, position, "paragraph");
    const copied = copiedParagraph(model, named, position);
    const inherited =
      operation["inheritFormatting"] === false
        ? undefined
        : liveParagraph(copied);
    const { fields: requested, checks } = paragraphRequestForLiveRow({
      model,
      request: operation,
      row: copied,
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
          source: copied,
          inheritsFormatting: operation["inheritFormatting"] !== false,
        });
      }
      if (holder) row.inTextBoxOf = holder;
      return row;
    });
    const first = rows.at(0);
    const shown =
      inherited === undefined
        ? undefined
        : inheritedListFields(model, inherited);
    if (
      first !== undefined &&
      shown?.kind === "listItem" &&
      "numbering" in operation &&
      first.fields.listLevel === undefined &&
      takesFinalMarkWhileSuggested(model, named, position)
    ) {
      // The paragraph takes the anchor's final mark, its stated numbering a
      // pending change from the anchor's. A pending change that shows no
      // marker of its own still shows the one it takes away, as for any
      // pending numbering removal.
      first.fields = {
        ...first.fields,
        ...(first.fields.kind === "heading" ? {} : { kind: shown.kind }),
        listLevel: shown.listLevel,
      };
      first.listLineage = currentLineage(model, copied);
    }
    pushBesideBlocks(model, anchor, position, rows);
    addComment(model, operation);
  };

/**
 * Whether a suggested paragraph put `position` of `named` takes its mark: on
 * the canonical engine, one put after the last paragraph of its story or
 * table cell carries that container's final mark from then on, with its own
 * properties pending over the anchor's. The legacy engine keeps the anchor's
 * mark and adds a paragraph with its own properties.
 */
const takesFinalMarkWhileSuggested = (
  model: Model,
  named: ModelRow,
  position: "before" | "after",
): boolean => {
  if (
    model.engine !== "canonical" ||
    model.mode !== "suggested" ||
    position !== "after" ||
    named.pre === undefined ||
    named.after.length > 0
  )
    return false;
  const rows = model.rows.flatMap(({ pre }) =>
    pre === undefined ? [] : [pre],
  );
  return endsContainer(rows, named.pre.id);
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
 * Whether accepting removes `row` whole: a block pending deletion, mark and
 * all, that no operation wrote into and no pending join runs through.
 */
const goesWhole = (row: ModelRow): boolean =>
  row.pendingDeletion === true &&
  row.pendingJoin !== true &&
  row.edits.length === 0;

/**
 * The blocks inserted between `row` and its neighbour on `position`'s side,
 * and the row whose gap that is. An insertion after one block and one before
 * the next land in one gap, in batch order, past the blocks between them that
 * accepting removes whole. A block a pending join runs through stops that.
 *
 * This is the engine's rule, read on the model rather than imported, since
 * the scenarios run against published packages: a block put after a block,
 * paragraph or table, lands right after it, one put before a block at the
 * start of the gap (`insertionGap`), and each lands behind the blocks the
 * batch put into that gap before it (`InsertionGapOrder`), which `gap.push`
 * says here.
 */
const gapOf = (
  model: Model,
  row: ModelRow,
  position: "before" | "after",
): { gap: ModelRow[]; bordering: ModelRow } => {
  if (position === "after") return { gap: row.after, bordering: row };
  let at = row;
  for (let index = model.rows.indexOf(row) - 1; ; index--) {
    const previous = model.rows[index];
    if (!previous || previous.pendingJoin)
      return { gap: at.before, bordering: at };
    if (!goesWhole(previous))
      return { gap: previous.after, bordering: previous };
    at = previous;
  }
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
  const { gap, bordering } = gapOf(model, anchor, position);
  const nearTable = bordering.pre?.table !== undefined;
  for (const row of rows) if (nearTable) row.besideTable = true;
  gap.push(...rows);
  // After a text box's last paragraph, a block put in the box and one put in
  // the body after the paragraph the box is drawn in share a gap, but not a
  // place: a reader lists the box's paragraphs before the body's next block,
  // so the box's blocks come first, whatever the batch order.
  const boxed = gap.filter((row) => row.inTextBoxOf !== undefined);
  if (boxed.length === 0 || boxed.length === gap.length) return;
  gap.splice(
    0,
    gap.length,
    ...boxed,
    ...gap.filter((row) => row.inTextBoxOf === undefined),
  );
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
    const view = textTarget(model, operation["blockId"]);
    const find = String(operation["find"]);
    const start = view.text.indexOf(find);
    assert.notEqual(
      start,
      -1,
      `applied, but "${find}" is not in "${view.text}"`,
    );
    view.row.edits.push({
      start: view.offset + start,
      end: view.offset + start + find.length,
      replace: String(operation["replace"]),
    });
    addComment(model, operation);
  },
  replaceRange: (model, operation) => {
    const { blockId, start, end } = rangeOf(operation);
    const range = projectedTextRange({ model, blockId, start, end });
    range.row.edits.push({
      start: range.start,
      end: range.end,
      replace: String(operation["replace"]),
    });
    addComment(model, operation);
  },
  formatRange: (model, operation) => {
    const { blockId, start, end } = rangeOf(operation);
    const {
      row,
      start: mappedStart,
      end: mappedEnd,
    } = projectedTextRange({ model, blockId, start, end });
    const formatting = operation["formatting"] as Record<string, unknown>;
    for (const property of ["bold", "italic", "underline"] as const) {
      if (formatting[property] === true)
        row.formats.push({ start: mappedStart, end: mappedEnd, property });
    }
  },
  commentOnRange: (model, operation) => {
    const range = rangeOf(operation);
    const { row, start, end } = projectedTextRange({ model, ...range });
    addComment(
      model,
      operation,
      () =>
        untouchedText(row) && !row.removed
          ? row.text.slice(start, end)
          : undefined,
      () => coveredRangeWasRemoved(model, row, start, end),
      () => coveredRangeIsAmbiguous(model, row, start, end),
    );
  },
  commentOnBlock: (model, operation) => {
    const view = textTarget(model, operation["blockId"]);
    const row = view.row;
    const quote = operation["quote"];
    const anchored = typeof quote === "string" ? quote : undefined;
    const localStart = anchored === undefined ? 0 : view.text.indexOf(anchored);
    const start = view.offset + localStart;
    const end = start + (anchored?.length ?? view.text.length);
    addComment(
      model,
      operation,
      () =>
        untouchedText(row) && !row.removed
          ? (anchored ?? view.text)
          : undefined,
      () => localStart >= 0 && coveredRangeWasRemoved(model, row, start, end),
      () => localStart >= 0 && coveredRangeIsAmbiguous(model, row, start, end),
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
    // The paragraphs of a text box drawn in it are listed right after it, in
    // a story of their own: the join meets the paragraph after them.
    const next = model.rows[model.rows.indexOf(pastTextBoxes(model, row)) + 1];
    if (row.pendingDeletion || next?.pendingDeletion) {
      // Whether the join keeps that paragraph's pending deletion is not modelled yet.
      throw new Unmodelled("a join with a paragraph pending deletion");
    }
    // A suggested join stays pending: the paragraph shows its own properties
    // until the join is accepted.
    if (model.mode === "suggested")
      row.ownWhileJoinPending ??= {
        fields: { ...row.fields },
        listLineage: row.listLineage,
      };
    row.mergeSeparator =
      typeof operation["separator"] === "string" ? operation["separator"] : "";
    if (
      model.mode !== "direct" &&
      next?.removed &&
      isContainerCarrier(model, next) &&
      // A block the batch puts after it ends the container instead: the
      // paragraph goes, and the join runs across it into that block.
      next.after.length === 0
    ) {
      // The batch runs the deletion first, so the join meets a paragraph with
      // no text left and has nothing to separate, as when no block survives.
      next.removed = false;
      next.edits.push({ start: 0, end: next.text.length, replace: "" });
      row.mergeSeparator = "";
    } else if (
      next?.pre !== undefined &&
      !next.removed &&
      next.text === "" &&
      next.edits.length === 0 &&
      next.before.length === 0 &&
      next.after.length === 0 &&
      model.deletedContentIds.has(next.pre.id) &&
      isContainerCarrier(model, next)
    ) {
      // Its words are pending deletions from before the batch, and it ends
      // its container: the join meets nothing that stays, so no separator.
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
        row.inInsertedTable = true;
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

type ExpireJoinedDestinationPredictionOptions = {
  model: Model;
  operationType: string;
  footprint: ReturnType<typeof sourceOperationFootprint> | undefined;
};
/** An unmodelled applied effect cannot prove its joined paragraph's final owner or quote. */
const expireJoinedDestinationPrediction = ({
  model,
  operationType,
  footprint,
}: ExpireJoinedDestinationPredictionOptions) => {
  if (model.pendingJoins.size === 0) return;
  assert.ok(
    footprint,
    "A joined destination effect must retain its pre-operation footprint",
  );
  if (footprint.type === "unsupported")
    return assert.fail(
      `Cannot prove an applied ${operationType} footprint: ${footprint.reason}`,
    );
  for (const ownership of model.pendingJoins.values())
    if (
      [...ownership.sources, ownership.destination].some(({ id }) =>
        footprint.ids.has(id),
      )
    )
      ownership.destinationPrediction = { type: "unknown", operationType };
};

type RecordJoinedDeletionOptions = {
  model: Model;
  operation: Operation;
  footprint: ReturnType<typeof sourceOperationFootprint> | undefined;
};

/** Follow the pieces that still join a tracked deletion's named paragraph. */
const recordJoinedDeletion = ({
  model,
  operation,
  footprint,
}: RecordJoinedDeletionOptions) => {
  if (model.engine !== "canonical" || model.mode !== "tracked-changes") return;
  // Formatting a paragraph leaves the live text and its pending join intact.
  if (operation.type === "setBlockParagraphProperties") return;
  const blockId = operation["blockId"];
  for (const join of model.pendingJoins.values()) {
    const sources = join.deletionSources;
    if (sources === undefined) continue;
    if (operation.type === "deleteBlock" && blockId === join.destination.id) {
      for (const [index, { row: source, from }] of sources.entries()) {
        model.deletedJoinRuns.set(source.id, {
          destination: join.destination.id,
          from,
          text:
            source.text +
            sources
              .slice(index + 1)
              .map(({ row }) => row.text)
              .join("") +
            join.destination.text,
        });
      }
      join.deletionSources = undefined;
      continue;
    }
    if (operation.type === "splitBlock") {
      const index = sources.findIndex(({ row }) => row.id === blockId);
      const source = index < 0 ? undefined : sources.at(index);
      const offset = operation["offset"];
      if (
        source !== undefined &&
        source.from === 0 &&
        typeof offset === "number" &&
        offset > 0 &&
        offset < source.row.text.length &&
        (operation["separator"] === undefined || operation["separator"] === "")
      ) {
        // The leading half ends at the inserted break. Any earlier sources
        // now join that half, while the trailing half keeps the pending join.
        join.deletionSources = [
          { row: source.row, from: offset },
          ...sources.slice(index + 1),
        ];
        continue;
      }
    }
    if (
      footprint?.type !== "known" ||
      footprint.ids.has(join.destination.id) ||
      sources.some(({ row }) => footprint.ids.has(row.id))
    )
      join.deletionSources = undefined;
  }
};

/** Apply one applied operation's expectation to `model`. */
export const expectOperation = (model: Model, operation: Operation): void => {
  const footprint =
    model.pendingJoins.size === 0
      ? undefined
      : sourceOperationFootprint({
          rows: model.sourceRows,
          operation,
          requiredParagraphIds: [...model.pendingJoins.values()].flatMap(
            ({ sources, destination }) =>
              [...sources, destination].map(({ id }) => id),
          ),
        });
  recordJoinedDeletion({ model, operation, footprint });
  if (
    operation.type === "deleteBlock" &&
    typeof operation["blockId"] === "string"
  ) {
    model.deletedBlockIds.add(operation["blockId"]);
  }
  if (
    operation.type === "replaceBlock" &&
    typeof operation["blockId"] === "string"
  ) {
    model.replacedBlocks.set(operation["blockId"], String(operation["text"]));
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
    expireJoinedDestinationPrediction({
      model,
      operationType: operation.type,
      footprint,
    });
    model.unmodelled.push(`${operation.type}: no block-text expectation yet`);
    return;
  }
  try {
    expectation(model, operation);
  } catch (error) {
    if (error instanceof Unmodelled) {
      expireJoinedDestinationPrediction({
        model,
        operationType: operation.type,
        footprint,
      });
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
  let joinNext:
    | {
        separator: string;
        container: Container;
        host: Expected;
        hostInTextBox: boolean;
      }
    | undefined;
  /** Suggested mode: the properties accepting a pending join gives the paragraph. */
  const joinedProperties = new Map<
    Expected,
    { fields: Fields; listLineage: ListLineage }
  >();
  const boxed = (row: ModelRow): boolean =>
    row.inTextBoxOf !== undefined ||
    (row.pre !== undefined && model.inTextBox?.has(row.pre.id) === true);
  // A join that meets a table ends before it: a paragraph cannot run into
  // a cell, so the words keep their own paragraph, separator and all. The
  // paragraphs of a text box drawn in the joined paragraph stay listed
  // after it; the join runs past them.
  const push = (entry: Expected, row: ModelRow) => {
    const container = row.pre?.table;
    const inTable = row.inInsertedTable === true;
    const inTextBox = boxed(row);
    if (joinNext !== undefined && inTextBox && !joinNext.hostInTextBox) {
      out.push(entry);
      return;
    }
    if (joinNext !== undefined) {
      const { host } = joinNext;
      if (!inTable && sameContainer(joinNext.container, container)) {
        if (model.mode === "suggested") {
          // A suggested join is pending: the live view lists both paragraphs,
          // the second already in the properties the joined one gets.
          host.text += joinNext.separator;
          const head = joinedProperties.get(host) ?? host;
          // A pending numbering removal can show the original marker only
          // when the trailing paragraph's mark exists in the original view.
          const keepsMarker =
            head.listLineage.type === "none" &&
            entry.listLineage.type === "reference" &&
            row.pre !== undefined &&
            !model.markInsertedIds.has(row.pre.id);
          const fields: Fields = keepsMarker
            ? {
                ...head.fields,
                kind:
                  head.fields.kind === "paragraph"
                    ? entry.fields.kind
                    : head.fields.kind,
                listLevel: entry.fields.listLevel,
              }
            : head.fields;
          const joined = {
            ...entry,
            fields,
            listLineage: keepsMarker ? entry.listLineage : head.listLineage,
          };
          joinedProperties.set(joined, head);
          out.push(joined);
          joinNext = undefined;
          return;
        }
        host.text = `${host.text}${joinNext.separator}${entry.text}`;
        // A join keeps the first paragraph's properties; formatting offsets no longer hold.
        host.formats = [];
        joinNext = undefined;
        return;
      }
      // Accepted, on the canonical engine, a join whose next paragraph goes
      // runs its words into the first cell of the table that follows, whose
      // paragraph keeps its own mark and properties. The legacy engine keeps
      // them in their own paragraph.
      if (
        !inTable &&
        model.mode !== "suggested" &&
        model.engine === "canonical" &&
        !(entry.pre && model.insertedRowIds.has(entry.pre.id)) &&
        joinNext.container === undefined &&
        container?.rowIndex === 0 &&
        container.cellIndex === 0 &&
        out.at(-1) === host
      ) {
        out.pop();
        out.push({
          text: `${host.text}${joinNext.separator}${entry.text}`,
          fields: entry.fields,
          checks: entry.checks,
          formats: [],
          listLineage: entry.listLineage,
        });
        joinNext = undefined;
        return;
      }
      host.text += joinNext.separator;
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
      const shown =
        (model.mode === "suggested" ? row.ownWhileJoinPending : undefined) ??
        row;
      segments.forEach((segment, index) =>
        push(
          {
            text: segment.text,
            fields:
              index === 0
                ? shown.fields
                : (segment.split?.second ?? shown.fields),
            checks: row.checks,
            formats: whole ? row.formats : [],
            listLineage: shown.listLineage,
            ...(whole && row.pre ? { pre: row.pre } : {}),
          },
          row,
        ),
      );
      const host = out.at(-1);
      if (row.mergeSeparator !== undefined && host !== undefined) {
        joinNext = {
          separator: row.mergeSeparator,
          container: row.pre?.table,
          host,
          hostInTextBox: boxed(row),
        };
        // The joined paragraph takes what the accepted model gives its first
        // part; a part a pending join already reached keeps that one.
        const lastSegment = segments.at(-1);
        if (model.mode === "suggested" && !joinedProperties.has(host))
          joinedProperties.set(host, {
            fields:
              segments.length === 1
                ? row.fields
                : (lastSegment?.split?.second ?? row.fields),
            listLineage: row.listLineage,
          });
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

type QuotePlacementsOptions = {
  text: string;
  quote: string;
  first: number;
  limit: number;
};
const quotePlacements = ({
  text,
  quote,
  first,
  limit,
}: QuotePlacementsOptions): number[] => {
  const starts: number[] = [];
  for (
    let start = first;
    start <= text.length - quote.length && start < limit;
    start++
  )
    if (text.startsWith(quote, start)) starts.push(start);
  return starts;
};

/** Prove the accepted quote lies wholly in the consumed part of a live join. */
const joinedDeletionTakesComment = (
  model: Model,
  accepted: Comment,
  live: Comment | undefined,
): boolean => {
  if (live === undefined || live.blockId === null || accepted.anchor === "")
    return false;
  const run = model.deletedJoinRuns.get(live.blockId);
  if (run === undefined || run.destination !== accepted.blockId) return false;
  const starts = quotePlacements({
    text: run.text,
    quote: accepted.anchor,
    first: 0,
    limit: model.liveTextById.get(live.blockId)?.length ?? 0,
  });
  return starts.length > 0 && starts.every((start) => start >= run.from);
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
  const view = paragraphTextProjection(model, entry.blockId);
  const mapped =
    view.type === "pending-join" && view.destinationPrediction.type === "exact"
      ? view
      : undefined;
  const located = mapped?.row ?? row;
  // A pending join's source paragraph is no modelled row of its own, so a
  // replaceBlock of it is not in any row's edits: every word of it goes on
  // acceptance, and with them a comment anchored in them.
  const replacement = model.replacedBlocks.get(entry.blockId);
  const joinSource = [...model.pendingJoins.values()].some(({ sources }) =>
    sources.some(({ id }) => id === entry.blockId),
  );
  if (joinSource && replacement !== undefined && entry.anchor !== "") {
    const text = model.liveTextById.get(entry.blockId) ?? "";
    const whole = [{ start: 0, end: text.length, replace: replacement }];
    const fates = quotePlacements({
      text,
      quote: entry.anchor,
      first: 0,
      limit: text.length,
    }).map((start) =>
      coveredRangeFate(
        model.engine,
        whole,
        text,
        start,
        start + entry.anchor.length,
      ),
    );
    if (fates.length > 0 && fates.every((fate) => fate === "removed"))
      return "removed";
    if (fates.some((fate) => fate !== "kept")) return "ambiguous";
  }
  // A later effect can leave a join's accepted owner unproven, but not bring
  // back words an edit of the joined paragraph itself already took: the edits
  // within its slice of the joined text still decide whether its comment goes.
  if (
    view.type === "pending-join" &&
    mapped === undefined &&
    view.row.pre !== undefined &&
    entry.anchor !== ""
  ) {
    const sliceEnd = view.offset + view.text.length;
    const within = view.row.edits.filter(
      (edit) => edit.start >= view.offset && edit.end <= sliceEnd,
    );
    const sourceText = view.row.pre.text;
    const fates = quotePlacements({
      text: sourceText,
      quote: entry.anchor,
      first: view.offset,
      limit: sliceEnd,
    })
      .filter((start) => start + entry.anchor.length <= sliceEnd)
      .map((start) =>
        coveredRangeFate(
          model.engine,
          within,
          sourceText,
          start,
          start + entry.anchor.length,
        ),
      );
    if (fates.length > 0 && fates.every((fate) => fate === "removed"))
      return "removed";
  }
  // An unproven join keeps the existing preservation check; it does not excuse loss.
  if (!located?.pre || located.removed || entry.anchor === "") return "kept";
  const quote = mapped === undefined ? entry.anchor : acceptedAnchor;
  if (quote === "") return "kept";
  const first = mapped?.offset ?? 0;
  const limit =
    mapped === undefined
      ? located.pre.text.length
      : mapped.offset + mapped.text.length;
  const starts = quotePlacements({
    text: located.pre.text,
    quote,
    first,
    limit,
  });
  if (starts.length === 0) return "kept";
  // A replacement's words go past the end of a comment it reaches, so one
  // whose every word the edits take goes with them, direct or accepted.
  const edits = [
    ...located.edits,
    ...located.splits.map(({ offset, consumed }) => ({
      start: offset,
      end: offset + consumed,
      replace: "",
    })),
  ];
  const sourceText = located.pre.text;
  const fates = starts.map((start) =>
    coveredRangeFate(
      model.engine,
      edits,
      sourceText,
      start,
      start + quote.length,
    ),
  );
  if (fates.every((fate) => fate === "removed")) return "removed";
  if (fates.some((fate) => fate !== "kept")) return "ambiguous";
  return "kept";
};

const PLAIN_WORD = /^[\p{L}\p{N}\p{M}\p{Pc}]+$/u;

type WholeReplacementAnchorOptions = {
  source: string;
  replacement: string;
  mode: Mode;
};

/** Predict a whole quote for an unambiguous one-word rewrite. */
const wholeReplacementAnchor = ({
  source,
  replacement,
  mode,
}: WholeReplacementAnchorOptions): string | undefined => {
  if (mode === "suggested") return source;
  if (!PLAIN_WORD.test(source) || !PLAIN_WORD.test(replacement))
    return undefined;
  const { prefix, suffix } = smallestChangeKeeps(source, replacement);
  // An insertion exactly at either end stays outside the comment.
  if (prefix === source.length || suffix === source.length) return source;
  // Insertion before a kept suffix is inside the range. Insertion after the
  // last covered text goes outside, leaving only the kept prefix.
  if (suffix > 0) return replacement;
  return prefix > 0 ? source.slice(0, prefix) : undefined;
};

type ExpectedKeptAnchorOptions = { model: Model; entry: Comment; mode: Mode };

/** Preserve untouched joined quotes and predict complete rewrites where possible. */
const expectedKeptAnchor = ({
  model,
  entry,
  mode,
}: ExpectedKeptAnchorOptions): string | undefined => {
  const view = paragraphTextProjection(model, entry.blockId);
  if (
    view.type === "pending-join" &&
    view.destinationPrediction.type === "exact" &&
    view.row.pre &&
    !view.row.removed &&
    view.row.splits.length === 0 &&
    view.row.mergeSeparator === undefined
  ) {
    const starts = quotePlacements({
      text: view.row.pre.text,
      quote: entry.anchor,
      first: view.offset,
      limit: view.offset + view.text.length,
    });
    if (
      starts.length > 0 &&
      starts.every((start) =>
        view.row.edits.every(
          (edit) =>
            edit.end < start || edit.start > start + entry.anchor.length,
        ),
      )
    )
      return entry.anchor;
    const edit = view.row.edits.at(0);
    if (
      entry.anchor === view.text &&
      view.row.edits.length === 1 &&
      edit?.start === view.offset &&
      edit.end === view.offset + view.text.length &&
      edit.replace !== "" &&
      !edit.replace.includes("\n")
    )
      return wholeReplacementAnchor({
        source: entry.anchor,
        replacement: edit.replace,
        mode,
      });
  }
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
  return edit.replace.includes("\n")
    ? undefined
    : wholeReplacementAnchor({
        source: entry.anchor,
        replacement: edit.replace,
        mode,
      });
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

const hasInterleavedText = (markup: string, clean: string): boolean =>
  clean !== "" && markup !== clean && containsInOrder(markup, clean);

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

/**
 * The id the words of `row`, a proven pending join's accepted owner, read
 * under once accepted. Deleting the last paragraph of a story or cell keeps
 * its mark (`deletionKeptBlock`): the paragraph before it runs into it, so
 * when the batch removes every paragraph after `row` in its container, the
 * words read under the id of that last one. Anything else after `row` (a
 * block the batch adds, a table) keeps them under `row`'s own id.
 */
const acceptedOwnerId = (model: Model, row: ModelRow): string => {
  const pre =
    row.pre ?? assert.fail("A proven pending join must have an accepted owner");
  if (row.after.length > 0) return pre.id;
  let carrier: Row | undefined;
  for (const later of model.rows.slice(model.rows.indexOf(row) + 1)) {
    // A live-only row (a join's source, a paragraph pending deletion) is not
    // in the accepted document the ids are read from.
    if (later.outsidePreState === true) continue;
    // Its own author removing a paragraph a pending insertion added retracts
    // it: it never carries the words, and the paragraph before keeps its
    // own mark.
    if (
      model.mode === "tracked-changes" &&
      later.removed &&
      later.pre !== undefined &&
      model.addedAfterIds.has(later.pre.id)
    )
      continue;
    // A cell ends where the next one starts; in a story, a table stands
    // between the words and that story's last paragraph.
    if (
      later.pre !== undefined &&
      pre.table !== undefined &&
      !sameParagraphContainer(pre, later.pre)
    )
      break;
    if (
      later.pre === undefined ||
      !later.removed ||
      later.before.length > 0 ||
      later.after.length > 0 ||
      !sameParagraphContainer(pre, later.pre)
    )
      return pre.id;
    carrier = later.pre;
  }
  return carrier?.id ?? pre.id;
};

export const compareComments = (
  model: Model,
  before: readonly Comment[],
  after: readonly ComparedComment[],
  afterRows: readonly Row[],
  liveBefore: readonly Comment[] = before,
  mode: Mode = "direct",
): string[] => {
  const problems: string[] = [];
  const remaining = [...after];
  const kept: { entry: Comment; located: Comment }[] = [];
  const gone: { entry: Comment; located: Comment; mustRetire: boolean }[] = [];
  const storyBlockIds = new Set(
    model.rows.flatMap((row) => (row.pre === undefined ? [] : [row.pre.id])),
  );
  for (const entry of before) {
    const live = liveBefore.find(({ id }) => id === entry.id);
    // The accepted quote belongs to the accepted model. Use a live quote
    // only when a proven pending join maps that component to its accepted owner.
    // A split can instead leave the live range start in pending deleted text
    // while its entire accepted coverage belongs to the new paragraph.
    const source = paragraphTextProjection(model, live?.blockId);
    // A block an operation removed hands its words to the block it ran into,
    // so its live quote is the one that moved with them.
    const removedAt = model.rows.findIndex(
      (row) => row.pre?.id === live?.blockId && row.removed,
    );
    const removedSource =
      removedAt !== -1 && model.rows[removedAt + 1]?.pre?.id === entry.blockId;
    const located =
      live &&
      live.blockId !== entry.blockId &&
      ((source.type === "pending-join" &&
        source.row.pre?.id === entry.blockId) ||
        removedSource)
        ? live
        : entry;
    let disposition =
      mode === "suggested"
        ? "kept"
        : anchorDisposition(model, located, entry.anchor);
    const mustRetire =
      mode === "tracked-changes" &&
      joinedDeletionTakesComment(model, entry, live);
    if (mustRetire) disposition = "removed";
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
    if (disposition === "removed") gone.push({ entry, located, mustRetire });
    else if (disposition === "kept") kept.push({ entry, located });
  }
  for (const { entry: comment, located } of kept) {
    // Locate by the live component, but predict the accepted quote rather
    // than pending deleted words visible only in the live comment range.
    const view = paragraphTextProjection(model, located.blockId);
    const anchor = expectedKeptAnchor({
      model,
      entry:
        view.type === "pending-join" &&
        view.destinationPrediction.type === "exact"
          ? { ...comment, blockId: located.blockId }
          : comment,
      mode,
    });
    const owner =
      view.type === "pending-join" &&
      view.destinationPrediction.type === "exact"
        ? acceptedOwnerId(model, view.row)
        : undefined;
    const index = remaining.findIndex(
      ({ id, text, blockId, anchor: actualAnchor }) =>
        id === comment.id &&
        text === comment.text &&
        (owner === undefined || blockId === owner) &&
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
    if (comment.removed?.() || comment.ambiguous?.()) continue;
    const anchor = comment.anchor?.();
    const index = remaining.findIndex(
      (candidate) =>
        !beforeIds.has(candidate.id) &&
        candidate.text === comment.text &&
        (anchor === undefined ||
          candidate.anchor === anchor ||
          (candidate.finalAnchor === anchor &&
            hasInterleavedText(candidate.anchor, anchor))),
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
  for (const { entry, located, mustRetire } of gone) {
    const survivor = remaining.find(({ id }) => id === entry.id);
    const view = paragraphTextProjection(model, located.blockId);
    // A block id locates the first paragraph containing an anchor; it does not
    // own the whole range. Splits and joins can move the surviving part of a
    // range to another paragraph before this batch deletes the old locator.
    // In that case the comment is still anchored and must survive. A thread
    // left on the deleted locator, moved to another block the batch also
    // deleted, or left in no block at all, is the orphan this check is meant
    // to catch. A proven joined slice also identifies its accepted owner:
    // retaining the removed quote there is not a move to a surviving range.
    if (
      survivor &&
      (mustRetire ||
        survivor.blockId === located.blockId ||
        (view.type === "pending-join" &&
          view.destinationPrediction.type === "exact" &&
          survivor.blockId === view.row.pre?.id) ||
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

type CompareNewCommentQuotesOptions = {
  reviewer: Reviewer;
  model: Model;
  before: readonly Comment[];
  after: Comment[];
  story: Story;
};

/** Range requests use clean text, while live comments display tracked markup. */
const withFinalNewCommentQuotes = ({
  reviewer,
  model,
  before,
  after,
  story,
}: CompareNewCommentQuotesOptions): ComparedComment[] => {
  if (model.comments.length === 0) return after;
  const beforeIds = new Set(before.map(({ id }) => id));
  const added = after.filter(({ id }) => !beforeIds.has(id));
  const mayHoldDeletedText = model.comments.some((comment) => {
    const anchor = comment.anchor?.();
    if (anchor === undefined || comment.removed?.() || comment.ambiguous?.())
      return false;
    return added.some(
      (candidate) =>
        candidate.text === comment.text &&
        hasInterleavedText(candidate.anchor, anchor),
    );
  });
  if (!mayHoldDeletedText) return after;
  // A save drops staged suggestions. Resolve a live fork so the clean view
  // checks the same anchored comment, including when it crosses a suggestion.
  const final = reviewer.fork();
  if (!final.resolveReviewedStory({ story, view: "final" })) return after;
  const quotes = new Map(
    commentsOf(final).map(({ id, anchor }) => [id, anchor]),
  );
  return after.map((comment) => {
    if (beforeIds.has(comment.id)) return comment;
    const finalAnchor = quotes.get(comment.id);
    return finalAnchor === undefined ? comment : { ...comment, finalAnchor };
  });
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
  resolutionScope: ResolutionScope;
  noteLifetime: NoteLifetime;
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

/**
 * A live state as it reads once saved. A paragraph the package names no id for
 * has one folio minted from its text and ordinal, marked `idStability`. Saving
 * can write that id or a new one, so such a block stands down to its position.
 * An authored id is compared as it is, and a block that moved still differs.
 */
export const withoutMintedIds = (live: string, before: string): string => {
  type LiveBlock = { id?: string; idStability?: string };
  const blocksOf = (state: string) =>
    (JSON.parse(state) as { blocks: LiveBlock[] }).blocks;
  const minted = new Set(
    blocksOf(before).flatMap(({ idStability }, ordinal) =>
      idStability === "positional" ? [ordinal] : [],
    ),
  );
  const state = JSON.parse(live) as { blocks: LiveBlock[] };
  for (const [ordinal, block] of state.blocks.entries()) {
    delete block.idStability;
    if (minted.has(ordinal)) block.id = `minted-${String(ordinal)}`;
  }
  return JSON.stringify(state);
};

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

/**
 * The accepted rows under the ids their live rows hold. A paragraph the package
 * names no id for gets one minted from its text and place when read, so the
 * accepted view, read from a reopened file, can call a live row something else.
 * Between rows both views name, rows pair by equal text, then in order when as
 * many remain on each side. An authored id never changes, so nothing pairs there.
 */
export const withLiveMintedIds = (
  accepted: readonly Row[],
  live: readonly Row[],
): Row[] => {
  const liveIds = new Set(live.map(({ id }) => id));
  const acceptedIds = new Set(accepted.map(({ id }) => id));
  const renamed = new Map<Row, string>();
  const pair = (rows: readonly Row[], others: readonly Row[]) => {
    const left = new Set(others);
    const unpaired: Row[] = [];
    for (const row of rows) {
      const same = [...left].find(({ text }) => text === row.text);
      if (same) {
        renamed.set(row, same.id);
        left.delete(same);
      } else {
        unpaired.push(row);
      }
    }
    if (unpaired.length === left.size) {
      const remaining = [...left];
      for (const [index, row] of unpaired.entries()) {
        const other = remaining[index];
        if (other) renamed.set(row, other.id);
      }
    }
  };
  let rows: Row[] = [];
  let others: Row[] = [];
  let at = 0;
  for (const row of accepted) {
    if (!liveIds.has(row.id)) {
      rows.push(row);
      continue;
    }
    while (at < live.length && live[at]?.id !== row.id) {
      const other = live[at];
      if (other && !acceptedIds.has(other.id)) others.push(other);
      at += 1;
    }
    at += 1;
    pair(rows, others);
    rows = [];
    others = [];
  }
  for (; at < live.length; at += 1) {
    const other = live[at];
    if (other && !acceptedIds.has(other.id)) others.push(other);
  }
  pair(rows, others);
  const named: Row[] = [];
  for (const row of accepted) {
    const id = renamed.get(row);
    named.push(id === undefined ? row : Object.assign({}, row, { id }));
  }
  return named;
};

/** Capture what the oracle needs before an apply. */
/** The reviewer reads a capture repeats, each answered once while nothing edits it. */
const READS = new Set<PropertyKey>([
  "toDocument",
  "getContent",
  "getComments",
  "snapshot",
  "snapshotStory",
  "listStories",
  "readNumberingDefinitions",
  "exportPendingSuggestions",
  "toBuffer",
]);

/**
 * The reviewer as one capture reads it. A capture only reads, and each read
 * converts the whole document again; on a long list-heavy file that made one
 * capture take seconds. Results are kept per method and arguments for the
 * capture's own lifetime, so the edit that follows reads afresh.
 */
const readOnce = (reviewer: Reviewer): Reviewer => {
  const cache = new Map<string, unknown>();
  return new Proxy(reviewer, {
    get(read, key, receiver) {
      const value: unknown = Reflect.get(read, key, receiver);
      if (!READS.has(key) || typeof value !== "function") return value;
      return (...args: unknown[]) => {
        const cacheKey = `${String(key)}:${JSON.stringify(args)}`;
        if (!cache.has(cacheKey))
          cache.set(cacheKey, Reflect.apply(value, read, args));
        return cache.get(cacheKey);
      };
    },
  });
};

/**
 * The ids of the comments anchored in the table that holds `blockId`, which
 * removing that table retires. Each engine allocates comment ids its own way,
 * so a pinned flow names the comments by where they are, not by number.
 */
export const commentIdsInTableOf = (
  reviewer: Reviewer,
  blockId: string,
): number[] => {
  const blocks = reviewer.getContent();
  const table = blocks.find(({ id }) => id === blockId)?.table;
  assert.ok(table, `block ${blockId} is not in a table`);
  const cells = new Set(
    blocks
      .filter((block) => block.table?.tableIndex === table.tableIndex)
      .map(({ id }) => id),
  );
  return reviewer
    .getComments()
    .filter(
      ({ blockId: anchor }) => typeof anchor === "string" && cells.has(anchor),
    )
    .map(({ id }) => id)
    .sort((left, right) => left - right);
};

/** The live paragraphs holding words whose deletion is pending. */
export const wordsDeletedIds = (reviewer: Reviewer): Set<string> =>
  new Set(
    reviewer
      .getChanges()
      .flatMap(({ type, blockId }) =>
        type === "deletion" && blockId !== null ? [blockId] : [],
      ),
  );

/** The live paragraphs whose own mark is a pending deletion. */
export const markDeletedIds = (reviewer: Reviewer): Set<string> =>
  new Set(
    reviewer
      .getChanges()
      .flatMap(({ type, blockId }) =>
        type === "paragraphMarkDeleted" && blockId !== null ? [blockId] : [],
      ),
  );

/** The inserted marks in this story, including those of staged suggestions. */
export const markInsertedIds = (
  reviewer: Reviewer,
  story: Story,
): Set<string> => {
  const ids = new Set<string>();
  const snapshot = reviewer.snapshotStory(story);
  if (snapshot === null) return ids;
  sourceDocumentOf(snapshot).descendants((node) => {
    if (node.type.name !== "paragraph") return;
    const { paraId, pPrMark } = expectParagraphAttrs(node);
    if (typeof paraId === "string" && pPrMark?.kind === "ins") ids.add(paraId);
  });
  return ids;
};

export const capture = async (
  editable: Reviewer,
  mode: Mode,
  { story = MAIN, step = "fresh" }: CaptureOptions = {},
): Promise<Pre> => {
  const reviewer = readOnce(editable);
  const live = liveState(reviewer, story);
  const liveComments = commentsOf(reviewer);
  const resolutionScope = intentResolutionScope(story);
  const context = { story, step, targets: featureIndex(reviewer, story) };
  if (mode === "suggested") {
    const rows = rowsOf(reviewer, story);
    return {
      mode,
      resolutionScope,
      noteLifetime: { type: "not-observed" },
      live,
      rows,
      resolvedStory: "present",
      liveRows: rows,
      // The marks pending deletion still say where a block put before a
      // paragraph lands.
      joinSources: {
        type: "live",
        markDeletedIds: markDeletedIds(reviewer),
        markInsertedIds: markInsertedIds(reviewer, story),
        goWhole: blanksGoingWhole(rows, markDeletedIds(reviewer)),
      },
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
  const accepted = await resolvedState({
    bytes,
    resolution: "accept",
    scope: resolutionScope,
  });
  const liveRows = rowsOf(reviewer, story);
  const liveIds = new Set(liveRows.map(({ id }) => id));
  const noteCommentIds = new Set(
    liveComments.flatMap(({ id, blockId }) =>
      blockId !== null && liveIds.has(blockId) ? [id] : [],
    ),
  );
  if (story.type === "footnote" || story.type === "endnote")
    closeCommentReplies(noteCommentIds, commentParentsOf(reviewer));
  const noteLifetime: NoteLifetime =
    story.type === "footnote" || story.type === "endnote"
      ? {
          type: "note",
          story,
          accepted: (
            await resolvedState({
              bytes,
              resolution: "accept",
              scope: { type: "document", story },
            })
          ).story,
          rejected: (
            await resolvedState({
              bytes,
              resolution: "reject",
              scope: { type: "document", story },
            })
          ).story,
          commentIds: noteCommentIds,
        }
      : { type: "not-observed" };
  return {
    mode,
    resolutionScope,
    noteLifetime,
    live,
    rows: withLiveMintedIds(accepted.rows, liveRows),
    resolvedStory: accepted.story,
    liveRows,
    joinSources: await markedJoinSourcesOf(bytes, editable.author),
    numberingSource: {
      type: "package",
      bytes: accepted.bytes,
      live: liveNumberingFacts(reviewer, story),
    },
    comments: accepted.comments,
    links: accepted.links,
    liveComments,
    ...(mode === "tracked-changes"
      ? {
          rejected: (
            await resolvedState({
              bytes,
              resolution: "reject",
              scope: resolutionScope,
            })
          ).rows,
        }
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
      withoutMintedIds(liveState(reviewer, pre.story), pre.live),
      withoutMintedIds(pre.live, pre.live),
      `${context}: nothing applied, but the document changed`,
    );
    return [];
  }
  const model = modelOf(
    pre.rows,
    pre.liveRows,
    pre.joinSources,
    pre.targets.inTextBox,
  );
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
  model.tables.inTextBox = pre.targets.inTextBox;
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
    ...(pre.joinSources.type === "live" && {
      goWhole: pre.joinSources.goWhole,
    }),
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
        case "deleteTableColumn": {
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
        }
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
  let comments: ComparedComment[];
  let links: LinkSnapshot;
  let bytes: Uint8Array | null = null;
  if (pre.mode === "suggested") {
    rows = rowsOf(reviewer, pre.story);
    comments = withFinalNewCommentQuotes({
      reviewer,
      model,
      before: pre.comments,
      after: commentsOf(reviewer),
      story: pre.story,
    });
    links = captureLinks(reviewer);
  } else {
    const saved = await save(reviewer);
    const accepted = await resolvedState({
      bytes: saved,
      resolution: "accept",
      scope: pre.resolutionScope,
    });
    ({ rows, comments, bytes, links } = accepted);
    let rejected: ResolvedState | undefined;
    if (pre.rejected) {
      rejected = await resolvedState({
        bytes: saved,
        resolution: "reject",
        scope: pre.resolutionScope,
      });
      assert.deepEqual(
        project(rejected.rows),
        project(pre.rejected),
        `${context}: rejecting every change does not give the document before back`,
      );
    }
    await assertNoteLifetime({
      before: pre.noteLifetime,
      bytes: saved,
      accepted,
      rejected,
    });
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
  rows: (
    await resolvedState({
      bytes: await save(reviewer),
      resolution,
      scope: { type: "document", story: MAIN },
    })
  ).rows,
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
