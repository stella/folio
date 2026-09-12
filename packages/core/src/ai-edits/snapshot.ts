import { panic } from "better-result";
import { sameTextFormatting } from "@stll/docx-core/model";
import type { Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import { expectParagraphAttrs } from "../prosemirror/attrs";
import { marksToTextFormatting } from "../prosemirror/conversion/fromProseDoc";
import { directParagraphAlignment } from "../prosemirror/paragraphAlignment";
import { directParagraphSpacing } from "../prosemirror/paragraphSpacing";
import {
  paragraphFormattingForRun,
  paragraphRunStyleContext,
  resolveEffectiveRunStyleFormatting,
  type RunStyleResolver,
} from "../prosemirror/runStyleFormatting";
import type { TextFormatting } from "../types/document";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import { deriveBlankBlockId, deriveBlockId, type FolioBlockId } from "../types/block-id";
import { buildCleanBlockText, type CleanBlockText } from "./clean-text";
import type {
  FolioAIBlock,
  FolioAIBlockAnchor,
  FolioAIBlockKind,
  FolioAIBlockPreviewRun,
  FolioAIBlockStructuralBoundary,
  FolioAIBlockTableLocation,
  FolioAIEditSnapshot,
  FolioAITextRangeHandle,
} from "./types";

type FolioAIEditSnapshotMetadata = {
  numberingReferenceKeys: readonly string[];
  sourceDocument: PMNode;
  storyTables: readonly FolioStoryTable[];
};

const metadataBySnapshot = new WeakMap<FolioAIEditSnapshot, FolioAIEditSnapshotMetadata>();

const metadataOf = (snapshot: FolioAIEditSnapshot): FolioAIEditSnapshotMetadata =>
  metadataBySnapshot.get(snapshot) ??
  panic("Metadata was requested for a snapshot that did not record it");

/** @internal Numbering references collected during the snapshot's document walk. */
export const numberingReferenceKeysOf = (snapshot: FolioAIEditSnapshot): readonly string[] =>
  metadataOf(snapshot).numberingReferenceKeys;

/** @internal Tables collected during the snapshot's document walk. */
export const storyTablesOf = (snapshot: FolioAIEditSnapshot): readonly FolioStoryTable[] =>
  metadataOf(snapshot).storyTables;

/** @internal The immutable ProseMirror document that produced this snapshot. */
export const sourceDocumentOf = (snapshot: FolioAIEditSnapshot): PMNode =>
  metadataOf(snapshot).sourceDocument;

export const normalizeFolioAIBlockText = (text: string): string =>
  text.replace(/\s+/gu, " ").trim();

/**
 * Whether a block carries text a reader would see.
 *
 * The snapshot holds every paragraph, blank ones included: an empty cell is
 * part of a table's shape, an empty row is a row, and a comparison that cannot
 * address them cannot describe what changed around them. A reading surface
 * wants the opposite — a model shown a document should see its content, not a
 * line per blank paragraph.
 *
 * So the two needs are split here rather than in the walk: the snapshot is
 * complete, and every surface that reads it for a person or a model states
 * that it wants content by calling this. One helper, so "what counts as
 * content" has a single answer; five inline emptiness checks would drift.
 */
export const isFolioAIContentBlock = ({ text }: Pick<FolioAIBlock, "text">): boolean =>
  normalizeFolioAIBlockText(text).length > 0;

/**
 * The block that content appended to the end of a story hangs from: the last
 * paragraph at body level.
 *
 * Not simply the last block. A story's last block is often inside a table, and
 * a paragraph cannot be appended after one: the insertion escapes to the
 * table's boundary, where the block it was anchored to is not adjacent to it
 * and its paragraph mark ends a different container's paragraph. Neither
 * container-edge rule then applies and the addition cannot be rejected
 * cleanly.
 *
 * A body may validly end with a table immediately before its final section
 * properties. Such a story has no body paragraph to anchor to, so this returns
 * `null`; callers that can place a peer beside the table use its outer boundary
 * instead.
 */
export const trailingBodyBlockId = ({ blocks }: FolioAIEditSnapshot): string | null => {
  for (let index = blocks.length - 1; index >= 0; index--) {
    const block = blocks[index];
    if (block !== undefined && block.table === undefined) {
      return block.id;
    }
  }
  return null;
};

export const hashFolioAIBlockText = (text: string): string => {
  let hash = 5381;
  for (const character of text) {
    hash = (hash * 33 + (character.codePointAt(0) ?? 0)) % 2_147_483_647;
  }
  return `h${hash.toString(36)}`;
};

const EMPTY_FOLIO_AI_BLOCK_STRUCTURAL_BOUNDARIES: readonly FolioAIBlockStructuralBoundary[] =
  Object.freeze([]);
const EMPTY_FOLIO_AI_BLOCK_STRUCTURAL_BOUNDARY_HASH = hashFolioAIBlockText(
  JSON.stringify(EMPTY_FOLIO_AI_BLOCK_STRUCTURAL_BOUNDARIES),
);

/** Canonical public projection of the clean view's zero-width structure. */
export const projectFolioAIBlockStructuralBoundaries = ({
  structuralBoundaries,
}: Pick<CleanBlockText, "structuralBoundaries">): readonly FolioAIBlockStructuralBoundary[] => {
  let projected: FolioAIBlockStructuralBoundary[] | undefined;
  for (const { clear, offset, presentInCleanView } of structuralBoundaries) {
    if (!presentInCleanView) {
      continue;
    }
    (projected ??= []).push({
      type: "pageBreak",
      offset,
      ...(clear !== undefined ? { clear } : {}),
    });
  }
  return projected ?? EMPTY_FOLIO_AI_BLOCK_STRUCTURAL_BOUNDARIES;
};

/** @internal Stable prefilter for an already-projected boundary sequence. */
export const hashFolioAIBlockStructuralBoundaryProjection = (
  structuralBoundaries: readonly FolioAIBlockStructuralBoundary[],
): string =>
  structuralBoundaries.length === 0
    ? EMPTY_FOLIO_AI_BLOCK_STRUCTURAL_BOUNDARY_HASH
    : hashFolioAIBlockText(JSON.stringify(structuralBoundaries));

/** Stable precondition fingerprint for a block's zero-width structure. */
export const hashFolioAIBlockStructuralBoundaries = (
  cleanBlock: Pick<CleanBlockText, "structuralBoundaries">,
): string =>
  hashFolioAIBlockStructuralBoundaryProjection(projectFolioAIBlockStructuralBoundaries(cleanBlock));

type CreateFolioAITextRangeHandleOptions = {
  blockId: string;
  text: string;
  startOffset: number;
  endOffset: number;
};

export const createFolioAITextRangeHandle = ({
  blockId,
  text,
  startOffset,
  endOffset,
}: CreateFolioAITextRangeHandleOptions): FolioAITextRangeHandle | null => {
  if (
    blockId.length === 0 ||
    !Number.isInteger(startOffset) ||
    !Number.isInteger(endOffset) ||
    startOffset < 0 ||
    endOffset <= startOffset ||
    endOffset > text.length
  ) {
    return null;
  }
  return {
    type: "textRange",
    story: "main",
    blockId,
    startOffset,
    endOffset,
    selectedTextHash: hashFolioAIBlockText(text.slice(startOffset, endOffset)),
  };
};

const TABLE_ROW_NODE_NAME = "tableRow";

const TABLE_ROLE_TABLE = "table";
const TABLE_ROLE_ROW = "row";

/**
 * One node on the path from the document to the block being visited: the node
 * itself, where it ends, and its index in its own parent.
 *
 * The walk below keeps this path instead of calling `doc.resolve(pos)` per
 * block. `resolve` re-descends from the root and finds each level's child by
 * scanning that level's fragment from index 0, so on a flat document of n
 * paragraphs it costs O(n) per block and the snapshot costs O(n^2). The path
 * is already known: a depth-first walk in document order visits every ancestor
 * before the block, so carrying it costs nothing and the snapshot is linear.
 */
type AncestorPathEntry = { node: PMNode; start: number; end: number; index: number };

/**
 * True for a `tableRow` marked `hidden` (OOXML `w:trPr/w:hidden` — Word never
 * renders the row). Nothing inside such a row may surface its text to the
 * AI/agent snapshot: an attacker DOCX could otherwise smuggle prompt-injection
 * instructions into a row that no human ever sees. The walk skips the row's
 * whole subtree, so a table nested inside a hidden row is hidden with it.
 */
export const isHiddenTableRow = (node: PMNode): boolean =>
  node.type.name === TABLE_ROW_NODE_NAME && node.attrs["hidden"] === true;

/** One table of a story, numbered the way {@link createFolioAIEditSnapshot} numbers it. */
export type FolioStoryTable = {
  /** Document-order index over every table, nested ones included. */
  index: number;
  /** The table node's position in the story document. */
  start: number;
  node: PMNode;
};

const STORY_TABLE_CONTAINER = "container";
const STORY_TABLE_HIDDEN_SUBTREE = "hidden-subtree";
const STORY_TABLE = "table";

type StoryTableNodeDisposition =
  | typeof STORY_TABLE_CONTAINER
  | typeof STORY_TABLE_HIDDEN_SUBTREE
  | typeof STORY_TABLE;

/** Shared visibility and table-role rule for both story projection walks. */
const classifyNonTextblockStoryTableNode = (node: PMNode): StoryTableNodeDisposition => {
  if (isHiddenTableRow(node)) {
    return STORY_TABLE_HIDDEN_SUBTREE;
  }
  return node.type.spec["tableRole"] === TABLE_ROLE_TABLE ? STORY_TABLE : STORY_TABLE_CONTAINER;
};

/**
 * Every table of one story in document order, nested tables included and
 * hidden rows' subtrees excluded.
 *
 * Uses the same traversal classification as the snapshot's integrated census,
 * so both surfaces skip and number the same nodes in the same order.
 */
export const folioStoryTables = (doc: PMNode): FolioStoryTable[] => {
  const tables: FolioStoryTable[] = [];
  doc.descendants((node, pos) => {
    if (node.isTextblock) {
      return false;
    }
    const disposition = classifyNonTextblockStoryTableNode(node);
    if (disposition === STORY_TABLE_HIDDEN_SUBTREE) {
      return false;
    }
    if (disposition === STORY_TABLE) {
      tables.push({ index: tables.length, start: pos, node });
    }
    return true;
  });
  return tables;
};

type TableLocationOptions = {
  path: readonly AncestorPathEntry[];
  /** The visited block's index in its own parent. */
  blockIndex: number;
  tableIndexByStart: ReadonlyMap<number, number>;
};

/**
 * Locate the block inside its innermost enclosing table, or `undefined` when
 * it sits outside every table.
 */
const getTableLocation = ({
  path,
  blockIndex,
  tableIndexByStart,
}: TableLocationOptions): FolioAIBlockTableLocation | undefined => {
  for (let cellDepth = path.length - 1; cellDepth > 1; cellDepth--) {
    const cell = path[cellDepth];
    if (!cell) {
      continue;
    }
    const role = cell.node.type.spec["tableRole"];
    if (role !== "cell" && role !== "header_cell") {
      continue;
    }
    const row = path[cellDepth - 1];
    const table = path[cellDepth - 2];
    if (
      !row ||
      !table ||
      row.node.type.spec["tableRole"] !== TABLE_ROLE_ROW ||
      table.node.type.spec["tableRole"] !== TABLE_ROLE_TABLE
    ) {
      return undefined;
    }
    const tableIndex = tableIndexByStart.get(table.start);
    const outerTable = path.find((entry) => entry.node.type.spec["tableRole"] === TABLE_ROLE_TABLE);
    const outerTableIndex =
      outerTable === undefined ? undefined : tableIndexByStart.get(outerTable.start);
    if (tableIndex === undefined || outerTableIndex === undefined) {
      return undefined;
    }
    const rectangle = TableMap.get(table.node).findCell(cell.start - table.start - 1);
    return {
      outerTableIndex,
      tableIndex,
      rowIndex: row.index,
      cellIndex: cell.index,
      gridColumnIndex: rectangle.left,
      columnSpan: rectangle.right - rectangle.left,
      rowSpan: rectangle.bottom - rectangle.top,
      paragraphIndex: blockIndex,
    };
  }
  return undefined;
};

const createFolioAIEditSnapshotInternal = (
  doc: PMNode,
  styleResolver: RunStyleResolver | null,
): FolioAIEditSnapshot => {
  const draftBlocks: {
    block: FolioAIBlock;
    anchor: Omit<FolioAIBlockAnchor, "hashOccurrenceCount">;
  }[] = [];
  const hashCounts = new Map<string, number>();
  const usedBlockIds = new Set<string>();
  const numberingReferenceKeys = new Set<string>();
  const tables: FolioStoryTable[] = [];
  const tableIndexByStart = new Map<number, number>();
  // The containers enclosing the node being visited, innermost last. Kept in
  // step with the walk so no block has to resolve its own position.
  const path: AncestorPathEntry[] = [];

  let blockIndex = 0;
  let blankIndex = 0;
  doc.descendants((node, pos, _parent, index) => {
    while (pos >= (path.at(-1)?.end ?? Number.POSITIVE_INFINITY)) {
      path.pop();
    }
    if (!node.isTextblock) {
      const disposition = classifyNonTextblockStoryTableNode(node);
      if (disposition === STORY_TABLE_HIDDEN_SUBTREE) {
        return false;
      }
      if (disposition === STORY_TABLE) {
        const tableIndex = tables.length;
        tables.push({ index: tableIndex, start: pos, node });
        tableIndexByStart.set(pos, tableIndex);
      }
      if (!node.isLeaf) {
        path.push({ node, start: pos, end: pos + node.nodeSize, index });
      }
      return true;
    }

    // Snapshot the AI-facing text in its post-tracked-changes
    // form: existing deletion-marked runs are skipped, existing
    // insertion-marked runs are included as plain text. The model
    // would otherwise see "shallmust" smashed together in a block
    // mid-edit and write find/replace operations against that
    // confused string. Apply uses the same clean view to resolve
    // operation positions, so the offsets stay consistent.
    const cleanBlock = buildCleanBlockText(node, pos);
    const { text } = cleanBlock;
    const structuralBoundaries = projectFolioAIBlockStructuralBoundaries(cleanBlock);
    const structuralBoundaryHash =
      hashFolioAIBlockStructuralBoundaryProjection(structuralBoundaries);
    const normalizedText = normalizeFolioAIBlockText(text);
    const textHash = hashFolioAIBlockText(normalizedText);
    hashCounts.set(textHash, (hashCounts.get(textHash) ?? 0) + 1);

    // Use the paragraph's Word `w14:paraId` (allocated by
    // `ParaIdAllocatorExtension` if the parsed DOCX didn't have one)
    // as the canonical block id everywhere: AI prompts, chip hrefs
    // (`#folio:<paraId>`), apply-tool blockIds, scrollToBlock. ParaIds
    // are stable across structural edits — no more "this chip points
    // at the wrong paragraph after an insertion-above" surprise.
    //
    // Shared with `apps/api/.../docx-blocks.ts` via `deriveBlockId`,
    // so a server-emitted citation id is always one of the shapes this
    // snapshot produces (paraId verbatim, `seq-NNNN`, or `blank-NNNN`).
    //
    // `seq-NNNN` counts the paragraphs that carry text, and nothing
    // else: that count is the published contract the server extractor
    // derives too, so a stored citation keeps naming its paragraph. A
    // paragraph with no text is numbered in the separate blank
    // sequence rather than consuming a position in it.
    const paraIdAttr: unknown = node.attrs["paraId"];
    const paraId = typeof paraIdAttr === "string" && paraIdAttr.length > 0 ? paraIdAttr : null;
    const idStabilityAttr: unknown = node.attrs["idStability"];
    const idStability = idStabilityAttr === "positional" ? "positional" : undefined;
    const isBlank = normalizedText.length === 0;
    let id: FolioBlockId;
    if (isBlank) {
      blankIndex++;
      id = deriveBlankBlockId({ paraId, index: blankIndex, taken: usedBlockIds });
    } else {
      blockIndex++;
      id = deriveBlockId({ paraId, index: blockIndex, taken: usedBlockIds });
    }
    usedBlockIds.add(id);
    const headingLevel = getHeadingLevel(node);
    const kind = getBlockKind(node, headingLevel);
    const displayLabel = getDisplayLabel(node);
    const styleId = getStyleId(node);
    const listLevel = getListLevel(node);
    const numberingReferenceKey = getNumberingReferenceKey(node);
    if (numberingReferenceKey) {
      numberingReferenceKeys.add(numberingReferenceKey);
    }
    const directAlignment = getDirectAlignment(node);
    const directSpacing = getDirectSpacing(node);
    const previewRuns = getPreviewRuns(node, styleResolver);
    const table = getTableLocation({ path, blockIndex: index, tableIndexByStart });

    draftBlocks.push({
      block: {
        id,
        kind,
        text,
        ...(idStability !== undefined && { idStability }),
        ...(headingLevel !== undefined && { headingLevel }),
        ...(displayLabel !== undefined && { displayLabel }),
        ...(styleId !== undefined && { styleId }),
        ...(listLevel !== undefined && { listLevel }),
        ...(directAlignment !== undefined && { directAlignment }),
        ...(directSpacing !== undefined && { directSpacing }),
        ...(previewRuns !== undefined && { previewRuns }),
        ...(structuralBoundaries.length > 0 && { structuralBoundaries }),
        ...(table !== undefined && { table }),
      },
      anchor: {
        id,
        from: pos,
        to: pos + node.nodeSize,
        text,
        normalizedText,
        textHash,
        structuralBoundaryHash,
      },
    });
    return true;
  });

  const blocks: FolioAIBlock[] = [];
  const anchors: Record<string, FolioAIBlockAnchor> = {};
  for (const draft of draftBlocks) {
    blocks.push(draft.block);
    anchors[draft.block.id] = {
      ...draft.anchor,
      hashOccurrenceCount: hashCounts.get(draft.anchor.textHash) ?? 0,
    };
  }

  const snapshot = { blocks, anchors };
  metadataBySnapshot.set(snapshot, {
    numberingReferenceKeys: [...numberingReferenceKeys],
    sourceDocument: doc,
    storyTables: tables,
  });
  return snapshot;
};

export const createFolioAIEditSnapshot = (doc: PMNode): FolioAIEditSnapshot =>
  createFolioAIEditSnapshotInternal(doc, null);

/** @internal Use for an EditorState that owns the document's style resolver. */
export const createFolioAIEditSnapshotWithStyleResolver = (
  doc: PMNode,
  styleResolver: RunStyleResolver | null,
): FolioAIEditSnapshot => createFolioAIEditSnapshotInternal(doc, styleResolver);

const getBlockKind = (node: PMNode, headingLevel: number | undefined): FolioAIBlockKind => {
  const listMarker: unknown = node.attrs["listMarker"];
  const numPr: unknown = node.attrs["numPr"];
  if (
    (typeof listMarker === "string" && listMarker.trim().length > 0) ||
    (numPr !== undefined && numPr !== null)
  ) {
    return "listItem";
  }

  if (headingLevel !== undefined) {
    return "heading";
  }

  return "paragraph";
};

const getHeadingLevel = (node: PMNode): number | undefined => {
  const outlineLevel: unknown = node.attrs["outlineLevel"];
  if (
    typeof outlineLevel === "number" &&
    Number.isInteger(outlineLevel) &&
    outlineLevel >= 0 &&
    outlineLevel <= 8
  ) {
    return outlineLevel + 1;
  }

  const styleId: unknown = node.attrs["styleId"];
  if (typeof styleId !== "string") {
    return undefined;
  }
  const match = /^heading(?<level>[1-9])$/iu.exec(styleId);
  const level = match?.groups?.["level"];
  return level === undefined ? undefined : Number.parseInt(level, 10);
};

const getDisplayLabel = (node: PMNode): string | undefined => {
  const listMarker: unknown = node.attrs["listMarker"];
  if (typeof listMarker === "string" && listMarker.trim().length > 0) {
    return listMarker.trim();
  }

  const styleId: unknown = node.attrs["styleId"];
  if (typeof styleId === "string" && /^heading/iu.test(styleId)) {
    return styleId;
  }

  return undefined;
};

const getListLevel = (node: PMNode): number | undefined => {
  const numPr: unknown = node.attrs["numPr"];
  if (typeof numPr !== "object" || numPr === null || !("ilvl" in numPr)) {
    return undefined;
  }
  const { ilvl } = numPr;
  return typeof ilvl === "number" && Number.isInteger(ilvl) && ilvl >= 0 ? ilvl : undefined;
};

const getNumberingReferenceKey = (node: PMNode): string | null => {
  const numPr: unknown = node.attrs["numPr"];
  if (typeof numPr !== "object" || numPr === null || !("numId" in numPr)) {
    return null;
  }
  const { numId } = numPr;
  if (typeof numId !== "number" || !Number.isInteger(numId) || numId <= 0) {
    return null;
  }
  const level = "ilvl" in numPr ? numPr.ilvl : undefined;
  if (level !== undefined && (typeof level !== "number" || !Number.isInteger(level) || level < 0)) {
    return null;
  }
  return `${String(numId)}:${String(level ?? 0)}`;
};

const getStyleId = (node: PMNode): string | undefined => {
  const styleId: unknown = node.attrs["styleId"];
  return typeof styleId === "string" && styleId.length > 0 ? styleId : undefined;
};

/** Read only authored `w:jc`, never the effective alignment resolved from a style. */
const getDirectAlignment = (node: PMNode) => directParagraphAlignment(expectParagraphAttrs(node));

/** Read only authored `w:spacing`, never effective spacing resolved from a style. */
const getDirectSpacing = (node: PMNode) => directParagraphSpacing(expectParagraphAttrs(node));

const DELETION_MARK = "deletion";
const HIDDEN_MARK = "hidden";
const RUN_FORMATTING_OVERRIDE_MARK = "runFormattingOverride";
const CHARACTER_STYLE_MARK = "characterStyle";

const nonemptyTextFormatting = (
  formatting: TextFormatting | undefined,
): TextFormatting | undefined =>
  formatting !== undefined && Object.keys(formatting).length > 0 ? formatting : undefined;

const samePreviewRunFormatting = (
  left: FolioAIBlockPreviewRun,
  right: FolioAIBlockPreviewRun,
): boolean =>
  sameTextFormatting(left.effectiveFormatting, right.effectiveFormatting) &&
  sameTextFormatting(left.authoredFormatting, right.authoredFormatting);

const getPreviewRuns = (
  node: PMNode,
  styleResolver: RunStyleResolver | null,
): FolioAIBlockPreviewRun[] | undefined => {
  const runs: FolioAIBlockPreviewRun[] = [];
  const carrierlessContext = paragraphRunStyleContext(node, null);
  let authoredContext: ReturnType<typeof paragraphRunStyleContext> | undefined;

  node.descendants((child) => {
    if (!child.isText || child.text === undefined) {
      return true;
    }
    if (
      child.marks.some((mark) => mark.type.name === DELETION_MARK || mark.type.name === HIDDEN_MARK)
    ) {
      return false;
    }

    const hasAuthorshipCarrier = child.marks.some(
      ({ type }) =>
        type.name === RUN_FORMATTING_OVERRIDE_MARK || type.name === CHARACTER_STYLE_MARK,
    );
    const context = hasAuthorshipCarrier
      ? (authoredContext ??= paragraphRunStyleContext(node, styleResolver))
      : carrierlessContext;
    const runStyleResolver = hasAuthorshipCarrier ? styleResolver : null;
    const observedFormatting = nonemptyTextFormatting(marksToTextFormatting(child.marks));
    const authoredFormatting = nonemptyTextFormatting(
      marksToTextFormatting(child.marks, {
        baseParagraphFormatting: context.baseParagraphFormatting,
        inheritedFormatting: context.paragraphFormatting,
        paragraphMarkFormatting: context.paragraphMarkFormatting,
        paragraphMarkPrecedesStyle: context.paragraphMarkPrecedesStyle,
        styleResolver: runStyleResolver,
      }),
    );
    const paragraphFormatting = paragraphFormattingForRun({
      context,
      ...(authoredFormatting !== undefined && { directFormatting: authoredFormatting }),
      marks: child.marks,
    });
    const inheritedFormatting = resolveEffectiveRunStyleFormatting({
      marks: child.marks,
      paragraphFormatting,
      styleResolver: runStyleResolver,
    });
    const effectiveFormatting = nonemptyTextFormatting(
      mergeTextFormatting(inheritedFormatting, observedFormatting),
    );
    const run: FolioAIBlockPreviewRun = {
      text: child.text,
      ...(effectiveFormatting !== undefined && { effectiveFormatting }),
      ...(authoredFormatting !== undefined && { authoredFormatting }),
    };
    const previous = runs.at(-1);
    if (previous && samePreviewRunFormatting(previous, run)) {
      runs[runs.length - 1] = { ...previous, text: previous.text + child.text };
      return false;
    }

    runs.push(run);
    return false;
  });

  if (
    runs.every(
      ({ effectiveFormatting, authoredFormatting }) =>
        effectiveFormatting === undefined && authoredFormatting === undefined,
    )
  ) {
    return undefined;
  }

  return runs;
};
