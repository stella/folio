import { panic } from "better-result";
import { Fragment } from "prosemirror-model";
import type { Mark, Node as PMNode } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import {
  type BuiltInStyleIndex,
  EMPTY_BUILT_IN_STYLE_INDEX,
  resolveHeadingLevel,
} from "../docx/builtInStyles";
import {
  expectCharacterStyleMarkAttrs,
  expectHyperlinkMarkAttrs,
  expectParagraphAttrs,
  expectRunFormattingOverrideMarkAttrs,
} from "../prosemirror/attrs";
import { marksToTextFormatting } from "../prosemirror/conversion/fromProseDoc";
import {
  type ResolvedParagraphNumbering,
  resolveParagraphNumbering,
} from "../docx/numberingReference";
import { createListLabelCounter } from "../prosemirror/listLabels";
import { paragraphNumberingAttr, readParagraphNumberingAttr } from "../prosemirror/numberingAttr";
import { readOutlineLevelAttr } from "../prosemirror/outlineLevelAttr";
import { directParagraphAlignment } from "../prosemirror/paragraphAlignment";
import { directParagraphIndentation } from "../prosemirror/paragraphIndentation";
import { directParagraphSpacing } from "../prosemirror/paragraphSpacing";
import { paragraphRunStyleContext, type RunStyleResolver } from "../prosemirror/runStyleFormatting";
import { runFormattingInlineAtomCleanText } from "../prosemirror/runFormattingInlineCarriers";
import { authoredRunFormattingFromAttrs } from "../prosemirror/runFormattingProvenance";
import {
  readAuthoredRunFormatting,
  reconcileRunFormattingMarks,
} from "../prosemirror/runFormattingReconciliation";
import { recreateProseNodeWithParagraphPropertySource } from "../docx/paragraphPropertySource";
import { isAltChunkMarkup } from "../docx/altChunk";
import {
  OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC,
  isOpaqueNestedRowMarkup,
  opaqueRevisionCarrierName,
} from "../docx/opaqueCarrier";
import type { TextFormatting } from "../types/document";
import { deriveBlankBlockId, deriveBlockId, type FolioBlockId } from "../types/block-id";
import { splitsSurrogatePair } from "./character-boundaries";
import {
  buildCleanBlockText,
  type CleanBlockText,
  type CleanTextStructuralBoundary,
} from "./clean-text";
import {
  createNoteReferenceLabeler,
  type NoteReferenceLabeler,
  type NoteReferenceLabels,
} from "./note-references";
import type {
  FolioAIBlock,
  FolioAIBlockAnchor,
  FolioAIBlockKind,
  FolioAIBlockPreviewRun,
  FolioAIBlockStructuralBoundary,
  FolioAIBlockTableLocation,
  FolioAIEditSnapshot,
  FolioAIInlineFormatting,
  FolioAITextRangeHandle,
} from "./types";

type FolioAIEditSnapshotMetadata = {
  numberingReferenceKeys: readonly string[];
  sourceDocument: PMNode;
  styleResolver: RunStyleResolver | null;
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

/** @internal Style context that produced the snapshot's authored run properties. */
export const styleResolverOf = (snapshot: FolioAIEditSnapshot): RunStyleResolver | null =>
  metadataOf(snapshot).styleResolver;

/**
 * Derive a comparison-only view with concrete numbering references rebound.
 * Rebuild from the remapped source document so blocks, anchors, table nodes,
 * and operation templates retain one canonical set of numbering references.
 */
export const remapFolioAIEditSnapshotNumberingReferences = (
  snapshot: FolioAIEditSnapshot,
  numIdMap: ReadonlyMap<number, number>,
): FolioAIEditSnapshot => {
  if (numIdMap.size === 0) {
    return snapshot;
  }
  const metadata = metadataOf(snapshot);
  const remapNode = (node: PMNode): PMNode => {
    const numPr = readParagraphNumberingAttr(node.attrs["numPr"]);
    if (numPr?.kind !== "reference") {
      return node;
    }
    const remappedNumId = numIdMap.get(numPr.numId);
    if (remappedNumId === undefined) {
      return node;
    }
    return recreateProseNodeWithParagraphPropertySource(node, {
      attrs: {
        ...node.attrs,
        numPr: paragraphNumberingAttr({ ...numPr, numId: remappedNumId }),
      },
    });
  };
  const remapped = remapDocument(metadata.sourceDocument, remapNode);
  if (remapped === metadata.sourceDocument) {
    return snapshot;
  }
  return createFolioAIEditSnapshotInternal(remapped, metadata.styleResolver);
};

const remapDocument = (doc: PMNode, remapNode: (node: PMNode) => PMNode): PMNode => {
  const rewrite = (node: PMNode): PMNode => {
    if (node.isText) return remapNode(node);
    const children: PMNode[] = [];
    let changed = false;
    node.forEach((child) => {
      const next = rewrite(child);
      if (next !== child) changed = true;
      children.push(next);
    });
    const withChildren = changed
      ? recreateProseNodeWithParagraphPropertySource(node, {
          content: Fragment.fromArray(children),
        })
      : node;
    return remapNode(withChildren);
  };
  return rewrite(doc);
};

const externalHrefCanCrossPackage = (href: string): boolean =>
  href.length > 0 && !href.startsWith("#");

const hyperlinkCanCrossPackage = ({ href, rId }: ReturnType<typeof expectHyperlinkMarkAttrs>) =>
  externalHrefCanCrossPackage(href) ||
  (href.length === 0 && !(typeof rId === "string" && rId.length > 0));

/**
 * Detach external hyperlink relationship ids before copying a story into a
 * different package. The serializer allocates relationship ids for the
 * receiving part; bookmark targets remain package-local and are refused.
 */
export const detachFolioAIEditSnapshotExternalHyperlinks = (
  snapshot: FolioAIEditSnapshot,
): FolioAIEditSnapshot | null => {
  const metadata = metadataOf(snapshot);
  let portable = true;
  metadata.sourceDocument.descendants((node) => {
    const emptyHyperlinks = node.attrs["_emptyHyperlinks"];
    if (Array.isArray(emptyHyperlinks)) {
      for (const hyperlink of emptyHyperlinks) {
        if (
          typeof hyperlink !== "object" ||
          hyperlink === null ||
          "anchor" in hyperlink ||
          ("rId" in hyperlink &&
            (typeof hyperlink.href !== "string" || !externalHrefCanCrossPackage(hyperlink.href)))
        ) {
          portable = false;
          return false;
        }
      }
    }
    for (const mark of node.marks) {
      if (mark.type.name !== "hyperlink") continue;
      if (!hyperlinkCanCrossPackage(expectHyperlinkMarkAttrs(mark))) {
        portable = false;
        return false;
      }
    }
    return true;
  });
  if (!portable) return null;
  const remapNode = (node: PMNode): PMNode => {
    const emptyHyperlinks = node.attrs["_emptyHyperlinks"];
    const attrs = Array.isArray(emptyHyperlinks)
      ? {
          ...node.attrs,
          _emptyHyperlinks: emptyHyperlinks.map(({ rId: _rId, ...hyperlink }) => hyperlink),
        }
      : node.attrs;
    let marks: readonly Mark[] = node.marks;
    for (const mark of node.marks) {
      if (mark.type.name !== "hyperlink") continue;
      const { rId: _rId, ...hyperlink } = expectHyperlinkMarkAttrs(mark);
      marks = marks.map((candidate) =>
        candidate === mark ? mark.type.create(hyperlink) : candidate,
      );
    }
    if (attrs === node.attrs && marks === node.marks) return node;
    if (node.isText) return node.mark(marks);
    return recreateProseNodeWithParagraphPropertySource(node, { attrs, marks });
  };
  const remapped = remapDocument(metadata.sourceDocument, remapNode);
  return remapped === metadata.sourceDocument
    ? snapshot
    : createFolioAIEditSnapshotInternal(remapped, metadata.styleResolver);
};

type RemapFolioAIEditSnapshotStyleReferencesOptions = {
  snapshot: FolioAIEditSnapshot;
  styleIdMap: ReadonlyMap<string, string>;
  defaultParagraphStyleId: string | undefined;
  importedStyleResolver: RunStyleResolver | null;
  reconcileAuthoredFormatting?: boolean;
};

/** Rebind imported style identifiers while retaining the source formatting context. */
export const remapFolioAIEditSnapshotStyleReferences = ({
  snapshot,
  styleIdMap,
  defaultParagraphStyleId,
  importedStyleResolver,
  reconcileAuthoredFormatting = false,
}: RemapFolioAIEditSnapshotStyleReferencesOptions): FolioAIEditSnapshot => {
  if (
    defaultParagraphStyleId === undefined &&
    (styleIdMap.size === 0 || [...styleIdMap].every(([source, target]) => source === target))
  ) {
    return snapshot;
  }
  const metadata = metadataOf(snapshot);
  const remapNode = (node: PMNode): PMNode => {
    const styleId = node.attrs["styleId"];
    const remappedStyleId = typeof styleId === "string" ? styleIdMap.get(styleId) : undefined;
    if (remappedStyleId !== undefined) {
      return recreateProseNodeWithParagraphPropertySource(node, {
        attrs: { ...node.attrs, styleId: remappedStyleId },
      });
    }
    if (node.isTextblock && typeof styleId !== "string" && defaultParagraphStyleId !== undefined) {
      return recreateProseNodeWithParagraphPropertySource(node, {
        attrs: { ...node.attrs, styleId: defaultParagraphStyleId },
      });
    }
    if (!node.isText) return node;
    let marks: readonly Mark[] = node.marks;
    for (const mark of node.marks) {
      if (mark.type.name !== "characterStyle") continue;
      const attrs = expectCharacterStyleMarkAttrs(mark);
      const remapped = styleIdMap.get(attrs.styleId);
      if (remapped === undefined) continue;
      marks = marks.map((candidate) =>
        candidate === mark ? mark.type.create({ ...attrs, styleId: remapped }) : candidate,
      );
    }
    return marks === node.marks ? node : node.mark(marks);
  };
  const remapped = remapDocument(metadata.sourceDocument, remapNode);
  if (remapped === metadata.sourceDocument) return snapshot;
  if (!reconcileAuthoredFormatting) {
    return createFolioAIEditSnapshotInternal(remapped, importedStyleResolver);
  }
  const rebindAuthoredInlineFormatting = (source: PMNode, candidate: PMNode): PMNode => {
    if (source.childCount !== candidate.childCount) {
      return panic("Style remapping changed the document structure");
    }
    if (source.isInline) {
      const hasAuthoredFormatting = source.marks.some(
        ({ type }) => type.name === "runFormattingOverride" || type.name === "characterStyle",
      );
      if (!hasAuthoredFormatting) return candidate;
      // Inline nodes are reconciled by their paragraph parent below, where the
      // two style cascades are available.
      return candidate;
    }
    const sourceContext =
      source.type.name === "paragraph"
        ? paragraphRunStyleContext(source, metadata.styleResolver)
        : undefined;
    const candidateContext =
      candidate.type.name === "paragraph"
        ? paragraphRunStyleContext(candidate, importedStyleResolver)
        : undefined;
    const children: PMNode[] = [];
    let changed = false;
    source.forEach((sourceChild, _offset, index) => {
      const candidateChild = candidate.child(index);
      let next = rebindAuthoredInlineFormatting(sourceChild, candidateChild);
      if (sourceContext && candidateContext && sourceChild.isInline) {
        const hasAuthoredFormatting = sourceChild.marks.some(
          ({ type }) => type.name === "runFormattingOverride" || type.name === "characterStyle",
        );
        if (hasAuthoredFormatting) {
          const authoredFormatting = readAuthoredRunFormatting({
            context: sourceContext,
            marks: sourceChild.marks,
            styleResolver: metadata.styleResolver,
          });
          if (authoredFormatting.styleId !== undefined) {
            authoredFormatting.styleId =
              styleIdMap.get(authoredFormatting.styleId) ?? authoredFormatting.styleId;
          }
          const marks = reconcileRunFormattingMarks({
            authoredFormatting,
            context: candidateContext,
            node: candidateChild,
            styleResolver: importedStyleResolver,
          });
          next = candidateChild.isText
            ? candidateChild.mark(marks)
            : recreateProseNodeWithParagraphPropertySource(candidateChild, { marks });
        }
      }
      if (next !== candidateChild) changed = true;
      children.push(next);
    });
    return changed
      ? recreateProseNodeWithParagraphPropertySource(candidate, {
          content: Fragment.fromArray(children),
        })
      : candidate;
  };
  const rebound = rebindAuthoredInlineFormatting(metadata.sourceDocument, remapped);
  return createFolioAIEditSnapshotInternal(rebound, importedStyleResolver);
};

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

/**
 * Canonical public projection of the clean view's inline structure.
 *
 * A field boundary stays internal: it marks text the reader already sees, so it
 * belongs to range resolution rather than to a block's published structure. A
 * note reference is published: its marker reads like text, and a reader has to
 * be able to tell it is not.
 */
export const projectFolioAIBlockStructuralBoundaries = ({
  structuralBoundaries,
}: Pick<CleanBlockText, "structuralBoundaries">): readonly FolioAIBlockStructuralBoundary[] => {
  let projected: FolioAIBlockStructuralBoundary[] | undefined;
  for (const boundary of structuralBoundaries) {
    if (boundary.type === "noteReference") {
      (projected ??= []).push({
        type: "noteReference",
        noteType: boundary.noteType,
        offset: boundary.offset,
        length: boundary.length,
      });
      continue;
    }
    if (boundary.type !== "pageBreakRun" || !boundary.presentInCleanView) {
      continue;
    }
    (projected ??= []).push({
      type: "pageBreak",
      offset: boundary.offset,
      ...(boundary.clear !== undefined ? { clear: boundary.clear } : {}),
    });
  }
  return projected ?? EMPTY_FOLIO_AI_BLOCK_STRUCTURAL_BOUNDARIES;
};

const hashFolioAIBlockStructuralBoundaryProjection = (
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
    endOffset > text.length ||
    // Half an emoji is no text a range can name: see `character-boundaries.ts`.
    splitsSurrogatePair(text, startOffset) ||
    splitsSurrogatePair(text, endOffset)
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

const noteReferenceLabelsByStory = new WeakMap<PMNode, NoteReferenceLabels>();

/**
 * The marker every footnote/endnote reference of a story reads as: numbered in
 * reading order over the blocks {@link createFolioAIEditSnapshot} reads (a
 * hidden row's references are as absent as its text), so the applier, the
 * range resolvers and the snapshot all read the same text.
 *
 * The story is numbered on the first label asked for, not here: a reader
 * whose walk meets no reference (most documents, and every revision-stats
 * read) never pays for a second walk of the story.
 */
export const collectNoteReferenceLabels = (doc: PMNode): NoteReferenceLabels => {
  // A document node is immutable, so its numbering is too.
  const cached = noteReferenceLabelsByStory.get(doc);
  if (cached !== undefined) {
    return cached;
  }
  const labels = numberNoteReferencesLazily([doc]);
  noteReferenceLabelsByStory.set(doc, labels);
  return labels;
};

/** Number the stories' references in reading order, one story after another, on first use. */
const numberNoteReferencesLazily = (stories: readonly PMNode[]): NoteReferenceLabels => {
  let numbered: NoteReferenceLabels | undefined;
  const numberStories = (): NoteReferenceLabels => {
    const labeler = createNoteReferenceLabeler();
    for (const doc of stories) {
      doc.descendants((node) => {
        if (isHiddenTableRow(node)) {
          return false;
        }
        if (node.isTextblock) {
          labeler.numberBlock(node);
          return false;
        }
        return true;
      });
    }
    return labeler;
  };
  return { labelOf: (reference) => (numbered ??= numberStories()).labelOf(reference) };
};

/**
 * @internal The markers a comparison reads a revised story's references as:
 * each note the base story references keeps the base's marker, and a note
 * only the revised story references is numbered after them.
 *
 * Each story's own reading-order numbering would renumber every later marker
 * of the revised story as soon as a reference is added or removed before it,
 * and a text comparison would then read a reference both stories keep as an
 * edit of it. Numbering the revised story on from the base keeps a kept
 * reference's marker identical, and makes a reference the comparison brings
 * write the marker the base story (which the edits resolve against) gives it.
 */
export const alignedNoteReferenceLabels = (base: PMNode, revised: PMNode): NoteReferenceLabels =>
  numberNoteReferencesLazily([base, revised]);

/**
 * @internal The same snapshot with every note reference read through `labels`
 * rather than the story's own numbering. Only the text and structural
 * boundaries can change: block ids, anchors and the source document do not.
 */
export const relabelFolioAIEditSnapshotNoteReferences = (
  snapshot: FolioAIEditSnapshot,
  labels: NoteReferenceLabels,
): FolioAIEditSnapshot => {
  const metadata = metadataOf(snapshot);
  if (
    !snapshot.blocks.some((block) =>
      block.structuralBoundaries?.some(({ type }) => type === "noteReference"),
    )
  ) {
    return snapshot;
  }
  return createFolioAIEditSnapshotInternal(metadata.sourceDocument, metadata.styleResolver, labels);
};

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
  noteReferenceLabels?: NoteReferenceLabels,
): FolioAIEditSnapshot => {
  // Resolved once: the walk below classifies every block, and the index is
  // what lets a localized heading style reach the model as a heading.
  const builtInStyles = styleResolver?.builtInStyles ?? EMPTY_BUILT_IN_STYLE_INDEX;
  // Numbered block by block as the walk reads them, so each reference marker
  // carries the number the page shows (and the Markdown export writes), not the
  // package id; the same numbering `collectNoteReferenceLabels` gives. A
  // comparison supplies its own labels instead.
  const noteReferences: NoteReferenceLabeler = noteReferenceLabels
    ? { labelOf: noteReferenceLabels.labelOf, numberBlock: () => {} }
    : createNoteReferenceLabeler();
  const draftBlocks: {
    block: FolioAIBlock;
    anchor?: Omit<FolioAIBlockAnchor, "hashOccurrenceCount">;
  }[] = [];
  const hashCounts = new Map<string, number>();
  const usedBlockIds = new Set<string>();
  const numberingReferenceKeys = new Set<string>();
  // Labels are counted in this walk from the document as it stands, not read
  // from the markers the parser resolved at open: those go stale with the
  // first list edit.
  const nextListLabel = createListLabelCounter();
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
      const preserved = node.attrs["_preserved"];
      const preservedChildren =
        preserved !== null && typeof preserved === "object" && "children" in preserved
          ? preserved.children
          : undefined;
      if (Array.isArray(preservedChildren)) {
        for (const [carrierIndex, child] of preservedChildren.entries()) {
          if (child === null || typeof child !== "object" || !("xml" in child)) {
            continue;
          }
          const xml = child.xml;
          if (typeof xml !== "string") {
            continue;
          }
          const diagnostic = getOpaqueCarrierDiagnostic(xml);
          if (diagnostic === undefined) {
            continue;
          }
          draftBlocks.push({
            block: {
              id: `opaque-${String(pos)}-${String(carrierIndex)}`,
              kind: "diagnostic",
              text: diagnostic.text,
              diagnostic: { type: "opaqueCarrier", carrier: diagnostic.carrier },
            },
          });
        }
      }
      if (node.type.name === "preservedBlock") {
        const xml: unknown = node.attrs["xml"];
        if (typeof xml === "string") {
          const readerText: unknown = node.attrs["readerText"];
          const diagnostic = getOpaqueCarrierDiagnostic(
            xml,
            typeof readerText === "string" ? readerText : undefined,
          );
          if (diagnostic !== undefined) {
            const id = `opaque-${String(pos)}`;
            draftBlocks.push({
              block: {
                id,
                kind: "diagnostic",
                text: diagnostic.text,
                diagnostic: { type: "opaqueCarrier", carrier: diagnostic.carrier },
              },
            });
          }
        }
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
    noteReferences.numberBlock(node);
    const cleanBlock = buildCleanBlockText(node, pos, { fieldResults: "text", noteReferences });
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
    const headingLevel = getHeadingLevel(node, builtInStyles);
    const listLabel =
      node.type.name === "paragraph" ? nextListLabel(expectParagraphAttrs(node)) : undefined;
    const kind = getBlockKind(headingLevel, listLabel);
    const displayLabel = getDisplayLabel(node, listLabel, kind === "heading");
    const styleId = getStyleId(node);
    const directOutlineLevel =
      node.type.name === "paragraph"
        ? expectParagraphAttrs(node)._originalFormatting?.outlineLevel
        : undefined;
    const listLevel = getListLevel(node);
    const listReference = getListReference(node);
    const numberingReferenceKey = getNumberingReferenceKey(node);
    if (numberingReferenceKey) {
      numberingReferenceKeys.add(numberingReferenceKey);
    }
    const directAlignment = getDirectAlignment(node);
    const directSpacing = getDirectSpacing(node);
    const directIndentation = getDirectIndentation(node);
    const previewRuns = getPreviewRuns({ node, nodeFrom: pos, cleanBlock, styleResolver });
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
        ...(directOutlineLevel !== undefined && { directOutlineLevel }),
        ...(listLevel !== undefined && { listLevel }),
        ...(listReference !== undefined && { listReference }),
        ...(directAlignment !== undefined && { directAlignment }),
        ...(directSpacing !== undefined && { directSpacing }),
        ...(directIndentation !== undefined && { directIndentation }),
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
    if (draft.anchor === undefined) {
      continue;
    }
    anchors[draft.block.id] = {
      ...draft.anchor,
      hashOccurrenceCount: hashCounts.get(draft.anchor.textHash) ?? 0,
    };
  }

  const snapshot = { blocks, anchors };
  metadataBySnapshot.set(snapshot, {
    numberingReferenceKeys: [...numberingReferenceKeys],
    sourceDocument: doc,
    styleResolver,
    storyTables: tables,
  });
  return snapshot;
};

const getOpaqueCarrierDiagnostic = (
  xml: string,
  readerText?: string,
): { carrier: string; text: string } | undefined => {
  if (isAltChunkMarkup(xml)) {
    return {
      carrier: "w:altChunk",
      text:
        readerText === undefined
          ? "[Unsupported w:altChunk content]"
          : `[Unsupported w:altChunk content] ${readerText}`,
    };
  }
  const elementName = opaqueRevisionCarrierName(xml);
  if (elementName !== undefined) {
    return {
      carrier: elementName,
      text: OPAQUE_REVISION_CARRIER_READER_DIAGNOSTIC,
    };
  }
  return isOpaqueNestedRowMarkup(xml)
    ? { carrier: "w:tr", text: "[Unsupported nested w:tr content]" }
    : undefined;
};

export const createFolioAIEditSnapshot = (doc: PMNode): FolioAIEditSnapshot =>
  createFolioAIEditSnapshotInternal(doc, null);

/** @internal Use for an EditorState that owns the document's style resolver. */
export const createFolioAIEditSnapshotWithStyleResolver = (
  doc: PMNode,
  styleResolver: RunStyleResolver | null,
): FolioAIEditSnapshot => createFolioAIEditSnapshotInternal(doc, styleResolver);

/**
 * What a reader sees the block as, heading first: a numbered heading (`1.
 * Scope`, numbered through its style or its own `w:numPr`) is a heading that
 * shows a number, with the number in `displayLabel` and its level in
 * `listLevel`. A paragraph is a list item when it shows a marker, which is
 * when it has a label. One whose numbering shows none (a `w:vanish` level, a
 * level its list does not define, or `w:numId="0"` cancelling its style's
 * numbering) reads as prose and is a paragraph, keeping its `listLevel` and
 * `listReference`.
 */
const getBlockKind = (
  headingLevel: number | undefined,
  listLabel: string | undefined,
): FolioAIBlockKind => {
  if (headingLevel !== undefined) {
    return "heading";
  }
  return listLabel === undefined ? "paragraph" : "listItem";
};

/** The block's 1-based heading level, as {@link resolveHeadingLevel} classifies it. */
const getHeadingLevel = (node: PMNode, styles: BuiltInStyleIndex): number | undefined => {
  const styleId: unknown = node.attrs["styleId"];
  const level = resolveHeadingLevel(
    {
      outlineLevel: readOutlineLevelAttr(node.attrs["outlineLevel"]),
      styleId: typeof styleId === "string" ? styleId : null,
    },
    styles,
  );
  return level === undefined ? undefined : level + 1;
};

const getDisplayLabel = (
  node: PMNode,
  listLabel: string | undefined,
  isHeading: boolean,
): string | undefined => {
  // The number a reader sees beside the text, heading or not. A marker its
  // level hides (`w:vanish`) is not one, and has no label.
  if (listLabel !== undefined) {
    return listLabel;
  }

  // An unnumbered heading's label is its style id — what the model sees and
  // refers back to. It labels a block the classifier already decided is a
  // heading, so a localized id such as `Nadpis1` reaches the model instead of
  // being dropped for not starting with "heading".
  const styleId: unknown = node.attrs["styleId"];
  if (isHeading && typeof styleId === "string") {
    return styleId;
  }

  return undefined;
};

/**
 * The level the paragraph's `<w:numPr>` states, and only that. An absent
 * `w:ilvl` renders as level zero but is not a stated zero, and a caller that
 * writes it back would turn an untouched paragraph into one stating a level
 * its source never did.
 */
const getListLevel = (node: PMNode): number | undefined => {
  const numPr = readParagraphNumberingAttr(node.attrs["numPr"]);
  return numPr === null || numPr.kind === "none" ? undefined : numPr.ilvl;
};

/**
 * The numbering a paragraph attr states, read through the model's own reader.
 *
 * The two functions below used to test the reserved id relationally
 * (`numId > 0`, `numId <= 0`), which is the one spelling that disagreed with
 * the other four: a malformed package's negative id is a dangling reference
 * everywhere else and "not numbered" here.
 */
const statedNumbering = (node: PMNode): ResolvedParagraphNumbering =>
  resolveParagraphNumbering(readParagraphNumberingAttr(node.attrs["numPr"]) ?? undefined);

const getListReference = (node: PMNode): FolioAIBlock["listReference"] | undefined => {
  const numbering = statedNumbering(node);
  return numbering.kind === "reference"
    ? { numId: numbering.numId, level: numbering.ilvl }
    : undefined;
};

const getNumberingReferenceKey = (node: PMNode): string | null => {
  const numbering = statedNumbering(node);
  return numbering.kind === "reference"
    ? `${String(numbering.numId)}:${String(numbering.ilvl)}`
    : null;
};

const getStyleId = (node: PMNode): string | undefined => {
  const styleId: unknown = node.attrs["styleId"];
  return typeof styleId === "string" && styleId.length > 0 ? styleId : undefined;
};

/** Read only authored `w:jc`, never the effective alignment resolved from a style. */
const getDirectAlignment = (node: PMNode) => directParagraphAlignment(expectParagraphAttrs(node));

/** Read only authored `w:spacing`, never effective spacing resolved from a style. */
const getDirectSpacing = (node: PMNode) => directParagraphSpacing(expectParagraphAttrs(node));

/** Read only authored `w:ind`, never effective indentation resolved from a style. */
const getDirectIndentation = (node: PMNode) =>
  directParagraphIndentation(expectParagraphAttrs(node));

type PreviewRunStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSizePt?: number;
  color?: string;
};

const DELETION_MARK = "deletion";
const HIDDEN_MARK = "hidden";
const RUN_FORMATTING_OVERRIDE_MARK = "runFormattingOverride";
const CHARACTER_STYLE_MARK = "characterStyle";

type GetPreviewRunsOptions = {
  node: PMNode;
  nodeFrom: number;
  cleanBlock: CleanBlockText;
  styleResolver: RunStyleResolver | null;
};

const getPreviewRuns = ({
  node,
  nodeFrom,
  cleanBlock,
  styleResolver,
}: GetPreviewRunsOptions): FolioAIBlockPreviewRun[] | undefined => {
  const runs: FolioAIBlockPreviewRun[] = [];
  const defaultStyle = getDefaultPreviewRunStyle(node);
  let paragraphStyleContext: ReturnType<typeof paragraphRunStyleContext> | undefined;
  let cleanOffset = 0;

  node.descendants((child, relativePosition) => {
    const start = nodeFrom + 1 + relativePosition;
    // A note reference previews as the marker the clean text shows for it.
    const noteReference = cleanBlock.structuralBoundaries.find(
      (boundary): boundary is Extract<CleanTextStructuralBoundary, { type: "noteReference" }> =>
        boundary.type === "noteReference" && boundary.from === start,
    );
    let text: string | null | undefined;
    if (noteReference !== undefined) {
      text = cleanBlock.text.slice(
        noteReference.offset,
        noteReference.offset + noteReference.length,
      );
    } else {
      text = child.isText ? child.text : runFormattingInlineAtomCleanText(child);
    }
    if (text === undefined || text === null || text.length === 0) return true;
    if (
      child.marks.some((mark) => mark.type.name === DELETION_MARK || mark.type.name === HIDDEN_MARK)
    ) {
      return false;
    }
    while ((cleanBlock.offsets[cleanOffset] ?? Number.POSITIVE_INFINITY) < start) {
      cleanOffset++;
    }
    // A text node advances one PM position per character; an atom, and a note
    // reference's marker, hold all of their characters at one position.
    const lastCharacterPosition =
      child.isText && noteReference === undefined ? start + text.length - 1 : start;
    if (
      cleanBlock.offsets[cleanOffset] !== start ||
      cleanBlock.offsets[cleanOffset + text.length - 1] !== lastCharacterPosition
    ) {
      return false;
    }
    cleanOffset += text.length;

    const style = getPreviewRunStyle(child.marks, defaultStyle);
    const hasAuthorshipCarrier = child.marks.some(
      ({ type }) =>
        type.name === RUN_FORMATTING_OVERRIDE_MARK || type.name === CHARACTER_STYLE_MARK,
    );
    let directFormatting: PreviewRunStyle;
    if (!hasAuthorshipCarrier) {
      directFormatting = getDirectPreviewRunStyleFromMarks(child.marks, defaultStyle);
    } else {
      paragraphStyleContext ??= paragraphRunStyleContext(node, styleResolver);
      const directTextFormatting = marksToTextFormatting(child.marks, {
        baseParagraphFormatting: paragraphStyleContext.baseParagraphFormatting,
        inheritedFormatting: paragraphStyleContext.paragraphFormatting,
        paragraphMarkFormatting: paragraphStyleContext.paragraphMarkFormatting,
        paragraphMarkPrecedesStyle: paragraphStyleContext.paragraphMarkPrecedesStyle,
        styleResolver,
      });
      directFormatting = getDirectPreviewRunStyle(child.marks, directTextFormatting, styleResolver);
    }
    const previous = runs.at(-1);
    if (
      previous &&
      samePreviewRunStyle(previous, style) &&
      sameDirectFormatting(previous.directFormatting, directFormatting)
    ) {
      previous.text += text;
      return false;
    }

    runs.push({
      text,
      ...style,
      ...(!isEmptyPreviewRunStyle(directFormatting) && { directFormatting }),
    });
    return false;
  });

  if (runs.every(isUnstyledPreviewRun)) {
    return undefined;
  }

  return runs;
};

const getDefaultPreviewRunStyle = (node: PMNode): PreviewRunStyle => {
  const formatting: unknown = node.attrs["defaultTextFormatting"];
  if (typeof formatting !== "object" || formatting === null) {
    return {};
  }

  return {
    ...getBooleanTextFormatting(formatting),
    ...getFontSizeTextFormatting(formatting),
    ...getFontFamilyTextFormatting(formatting),
    ...getColorTextFormatting(formatting),
  };
};

const getPreviewRunStyle = (
  marks: readonly Mark[],
  defaultStyle: PreviewRunStyle,
): PreviewRunStyle => {
  const style: PreviewRunStyle = { ...defaultStyle };

  for (const mark of marks) {
    switch (mark.type.name) {
      case "bold":
        style.bold = true;
        break;
      case "italic":
        style.italic = true;
        break;
      case "underline":
        if (isUnderlineEnabled(mark.attrs["style"])) {
          style.underline = true;
        }
        break;
      case "strike":
        style.strike = true;
        break;
      case "fontSize": {
        const size = Number(mark.attrs["size"]);
        if (Number.isFinite(size) && size > 0) {
          style.fontSizePt = size / 2;
        }
        break;
      }
      case "fontFamily": {
        const fontFamily = getFontFamilyFromAttrs(mark.attrs);
        if (fontFamily !== undefined) {
          style.fontFamily = fontFamily;
        }
        break;
      }
      case "textColor": {
        const color = getColorFromAttrs(mark.attrs);
        if (color !== undefined) {
          style.color = color;
        }
        break;
      }
      default:
        break;
    }
  }

  const overrideMark = marks.find(({ type }) => type.name === RUN_FORMATTING_OVERRIDE_MARK);
  if (!overrideMark) {
    return style;
  }
  const overrides = expectRunFormattingOverrideMarkAttrs(overrideMark);
  if (overrides.bold === false) {
    delete style.bold;
  }
  if (overrides.italic === false) {
    delete style.italic;
  }
  if (overrides.underline === "none") {
    delete style.underline;
  }
  const hasDoubleStrike = marks.some(
    (mark) => mark.type.name === "strike" && mark.attrs["double"] === true,
  );
  if (overrides.strike === false && !hasDoubleStrike) {
    delete style.strike;
  }
  if (overrides.color === "auto") {
    delete style.color;
  }

  return style;
};

const getDirectPreviewRunStyleFromMarks = (
  marks: readonly Mark[],
  inheritedStyle: PreviewRunStyle,
): PreviewRunStyle => {
  const markedStyle = getPreviewRunStyle(marks, {});
  const directStyle: PreviewRunStyle = {};
  for (const property of ["bold", "italic", "underline", "strike"] as const) {
    if (Boolean(markedStyle[property]) !== Boolean(inheritedStyle[property])) {
      directStyle[property] = Boolean(markedStyle[property]);
    }
  }
  if (
    markedStyle.fontFamily !== undefined &&
    markedStyle.fontFamily !== inheritedStyle.fontFamily
  ) {
    directStyle.fontFamily = markedStyle.fontFamily;
  }
  if (
    markedStyle.fontSizePt !== undefined &&
    markedStyle.fontSizePt !== inheritedStyle.fontSizePt
  ) {
    directStyle.fontSizePt = markedStyle.fontSizePt;
  }
  if (markedStyle.color !== undefined && markedStyle.color !== inheritedStyle.color) {
    directStyle.color = markedStyle.color;
  }
  return directStyle;
};

const getDirectPreviewRunStyle = (
  marks: readonly Mark[],
  formatting: TextFormatting,
  styleResolver: RunStyleResolver | null,
): PreviewRunStyle => {
  let directFormatting = formatting;
  if (!styleResolver) {
    const overrideMark = marks.find(({ type }) => type.name === RUN_FORMATTING_OVERRIDE_MARK);
    const authoredFormatting = overrideMark
      ? authoredRunFormattingFromAttrs(expectRunFormattingOverrideMarkAttrs(overrideMark))
      : undefined;
    if (authoredFormatting !== undefined) {
      directFormatting = authoredFormatting;
    } else if (marks.some(({ type }) => type.name === CHARACTER_STYLE_MARK)) {
      directFormatting = {};
    }
  }

  const directStyle: PreviewRunStyle = {};
  for (const property of ["bold", "italic"] as const) {
    if (directFormatting[property] !== undefined) {
      directStyle[property] = directFormatting[property];
    }
  }
  if (directFormatting.underline !== undefined) {
    directStyle.underline = isUnderlineEnabled(directFormatting.underline);
  }
  if (directFormatting.strike !== undefined) {
    directStyle.strike = directFormatting.strike;
  } else if (directFormatting.doubleStrike === true) {
    directStyle.strike = true;
  }
  const fontFamily = directFormatting.fontFamily
    ? getFontFamilyFromAttrs(directFormatting.fontFamily)
    : undefined;
  if (fontFamily !== undefined) {
    directStyle.fontFamily = fontFamily;
  }
  const fontSizePt = getFontSizeTextFormatting(directFormatting).fontSizePt;
  if (fontSizePt !== undefined) {
    directStyle.fontSizePt = fontSizePt;
  }
  const color = directFormatting.color ? getColorFromAttrs(directFormatting.color) : undefined;
  if (color !== undefined) {
    directStyle.color = color;
  }

  return directStyle;
};

const getBooleanTextFormatting = (formatting: object): PreviewRunStyle => ({
  ...(Reflect.get(formatting, "bold") === true && { bold: true }),
  ...(Reflect.get(formatting, "italic") === true && { italic: true }),
  ...(isUnderlineEnabled(Reflect.get(formatting, "underline")) && {
    underline: true,
  }),
  ...(Reflect.get(formatting, "strike") === true && { strike: true }),
});

const isUnderlineEnabled = (underline: unknown): boolean => {
  if (underline === undefined || underline === null || underline === false) {
    return false;
  }
  if (underline === true) {
    return true;
  }
  if (typeof underline === "string") {
    return underline !== "none";
  }
  if (typeof underline !== "object") {
    return false;
  }

  const style: unknown = Reflect.get(underline, "style");
  return style !== "none";
};

const getFontSizeTextFormatting = (formatting: object): PreviewRunStyle => {
  const fontSize = Number(Reflect.get(formatting, "fontSize"));
  if (!Number.isFinite(fontSize) || fontSize <= 0) {
    return {};
  }
  return { fontSizePt: fontSize / 2 };
};

const getFontFamilyTextFormatting = (formatting: object): PreviewRunStyle => {
  const fontFamilyValue: unknown = Reflect.get(formatting, "fontFamily");
  if (typeof fontFamilyValue !== "object" || fontFamilyValue === null) {
    return {};
  }

  const fontFamily = getFontFamilyFromAttrs(fontFamilyValue);
  return fontFamily === undefined ? {} : { fontFamily };
};

const getColorTextFormatting = (formatting: object): PreviewRunStyle => {
  const colorValue: unknown = Reflect.get(formatting, "color");
  if (typeof colorValue !== "object" || colorValue === null) {
    return {};
  }

  const color = getColorFromAttrs(colorValue);
  return color === undefined ? {} : { color };
};

const getFontFamilyFromAttrs = (attrs: object): string | undefined => {
  const ascii: unknown = Reflect.get(attrs, "ascii");
  if (typeof ascii === "string" && ascii.length > 0) {
    return ascii;
  }

  const hAnsi: unknown = Reflect.get(attrs, "hAnsi");
  if (typeof hAnsi === "string" && hAnsi.length > 0) {
    return hAnsi;
  }

  return undefined;
};

const getColorFromAttrs = (attrs: object): string | undefined => {
  const rgb: unknown = Reflect.get(attrs, "rgb") ?? Reflect.get(attrs, "val");
  if (typeof rgb !== "string" || !/^[0-9a-fA-F]{6}$/u.test(rgb)) {
    return undefined;
  }

  return `#${rgb}`;
};

const samePreviewRunStyle = (run: FolioAIBlockPreviewRun, style: PreviewRunStyle): boolean =>
  run.bold === style.bold &&
  run.italic === style.italic &&
  run.underline === style.underline &&
  run.strike === style.strike &&
  run.fontFamily === style.fontFamily &&
  run.fontSizePt === style.fontSizePt &&
  run.color === style.color;

const isEmptyPreviewRunStyle = ({
  bold,
  italic,
  underline,
  strike,
  fontFamily,
  fontSizePt,
  color,
}: FolioAIInlineFormatting): boolean =>
  bold === undefined &&
  italic === undefined &&
  underline === undefined &&
  strike === undefined &&
  fontFamily === undefined &&
  fontSizePt === undefined &&
  color === undefined;

const sameDirectFormatting = (
  left: FolioAIInlineFormatting | undefined,
  right: PreviewRunStyle,
): boolean =>
  (left === undefined && isEmptyPreviewRunStyle(right)) ||
  (left !== undefined &&
    left.bold === right.bold &&
    left.italic === right.italic &&
    left.underline === right.underline &&
    left.strike === right.strike &&
    left.fontFamily === right.fontFamily &&
    left.fontSizePt === right.fontSizePt &&
    left.color === right.color);

const isUnstyledPreviewRun = ({ directFormatting, ...style }: FolioAIBlockPreviewRun): boolean =>
  isEmptyPreviewRunStyle(style) &&
  (directFormatting === undefined || isEmptyPreviewRunStyle(directFormatting));
