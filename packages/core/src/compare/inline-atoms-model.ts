/**
 * Restore the inline atoms a comparison's text operations leave out, on the
 * document model: fields, embedded pictures and page breaks the revised story
 * has where the reviewed story has none (or the other way round), and the note
 * references whose markers those operations wrote as text.
 *
 * The reviewed story is read as accepting every change would leave it, so it
 * lines up block for block with the revised story; each character of that view
 * keeps the paragraph and offset it occupies in the story as it stands, and
 * the result is a batch of document operations addressed there. Nothing here
 * edits a document: the batch applies as one commit, in order.
 */

import { panic } from "better-result";

import { DRAWING_RAW_XML_MODES, relationshipIdOf } from "@stll/docx-core/model";
import {
  applyDocumentOp,
  type DeleteRangeOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  findStoryBody,
  type InsertContentOp,
  type NewIds,
  type OpStory,
  packageResourcesOpOf,
  revisionIdDemand,
  type RevisionStamp,
  type TextPosition,
} from "@stll/docx-core/ops";

import type { FolioRevisionStamp } from "../ai-edits/apply";
import type {
  NoteReferenceIdentity,
  NoteReferenceLabels,
} from "../ai-edits/note-references";
import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { drawingXmlWithoutIdentity } from "../docx/drawingIdNormalization";
import { rebindDrawingImageRelationship } from "../docx/drawingRelationships";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import type {
  BlockContent,
  ComplexField,
  DrawingContent,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  SimpleField,
  TextFormatting,
  TrackedChangeInfo,
} from "../types/document";
import {
  type DeepReadonly,
  readModel,
  type ReadonlyDocument,
  writableCopy,
} from "../types/documentModel";
import { canonicalJson } from "../utils/canonicalJson";
import {
  createNoteReferenceNumbering,
  noteReferenceMarker,
  type NoteReferenceNumbering,
} from "../utils/noteReferenceLabels";
import { base64ToBytes } from "../utils/textEncoding";
import {
  type AlignableAtom,
  type AtomActionDisposition,
  atomActionRangeCount,
  fieldResultOwnershipOf,
  insertionAnchors,
  matchingAtoms,
  MISSING_NOTE_DETAIL,
  ownsFieldResult,
  sameBlockTopology,
  sameSupportedAtoms,
} from "./inline-atom-alignment";

/** One story of a document model. */
export type InlineAtomStory = {
  document: ReadonlyDocument;
  story: OpStory;
};

export type MatchInlineAtomsOnModelOptions = InlineAtomStory & {
  /**
   * The revised story, accepted, and the snapshot the comparison planned
   * against, whose block ids name the revised blocks a change describes.
   */
  target: InlineAtomStory & { snapshot: FolioAIEditSnapshot };
  revisionStamp: FolioRevisionStamp;
  /** First id reserved for this comparison, before planned operations allocated ids. */
  originalRevisionIdSeed: number;
  author: string;
  maxRanges: number;
  /** The markers the comparison's operations wrote note references as. */
  noteReferenceLabels: NoteReferenceLabels;
  /**
   * The notes the base package defines, as `footnote:<id>` / `endnote:<id>`.
   * A reference the comparison writes may only name one of them. Defaults to
   * the notes the story already references.
   */
  baseNotes?: ReadonlySet<string>;
};

export type MatchInlineAtomsOnModelResult =
  | {
      status: "matched";
      /** Applies as one batch, in order, to `document`. */
      operations: readonly DocumentOp[];
      nextRevisionId: number;
      changedTargetBlockIds: readonly string[];
      rangeCount: number;
    }
  | {
      status: "unalignable";
      /** Why, when the cause is more specific than an alignment failure. */
      detail?: string;
    }
  | { status: "budget-exceeded" };

type InlineNode = ParagraphContent | RunContent;
type ReadonlyInline = DeepReadonly<InlineNode>;
type ReadonlyParagraph = DeepReadonly<Paragraph>;
type ReadonlyRun = DeepReadonly<Run>;
type ReadonlyField = DeepReadonly<SimpleField | ComplexField>;
type ReadonlyDrawing = DeepReadonly<DrawingContent>;

/** A place in a paragraph's offset space (see `TextPosition`). */
type Gap = { offset: number; zeroWidthBefore: number };

/** A place in the story as it stands, ordered by paragraph, then offset. */
type LivePoint = Gap & { paraId: string; order: number };

const comparePoints = (left: LivePoint, right: LivePoint): number =>
  left.order - right.order ||
  left.offset - right.offset ||
  left.zeroWidthBefore - right.zeroWidthBefore;

type LeafWidth = "characters" | "unit" | "zeroWidth";

/** One leaf of a paragraph's content tree, with what holds it. */
type ModelLeaf = {
  node: ReadonlyInline;
  width: LeafWidth;
  /** The run holding a run child. */
  run: ReadonlyRun | undefined;
  /** Inside a tracked deletion or move away: accepting removes it. */
  removed: boolean;
  /** The tracked insertion or move holding the leaf, innermost. */
  insertion: DeepReadonly<TrackedChangeInfo> | undefined;
  before: Gap;
  after: Gap;
};

const childNodesOf = (
  node: ReadonlyInline,
): readonly ReadonlyInline[] | undefined => {
  switch (node.type) {
    case "run":
      return node.content;
    case "hyperlink":
      return node.children;
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
    case "inlineSdt":
    case "inlineWrapper":
      return node.content;
    default:
      return undefined;
  }
};

/** Units of `offsets.ts`: what a leaf adds to its paragraph's offset space. */
const leafWidthOf = (node: ReadonlyInline): LeafWidth => {
  switch (node.type) {
    case "text":
      return node.text === "" ? "zeroWidth" : "characters";
    case "noteMarker":
    case "renderedPageBreak":
    case "bookmarkStart":
    case "bookmarkEnd":
    case "commentRangeStart":
    case "commentRangeEnd":
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd":
      return "zeroWidth";
    case "preservedInline":
      return node.text === "" ? "zeroWidth" : "unit";
    default:
      return childNodesOf(node) === undefined ? "unit" : "zeroWidth";
  }
};

/** Every leaf of a paragraph in document order, as `leafSpans` counts them. */
const paragraphLeaves = (paragraph: ReadonlyParagraph): ModelLeaf[] => {
  const leaves: ModelLeaf[] = [];
  let position = 0;
  let zeroWidthSeen = 0;
  type Holders = Pick<ModelLeaf, "run" | "removed" | "insertion">;
  const walk = (
    items: readonly ReadonlyInline[],
    { run, removed, insertion }: Holders,
  ): void => {
    for (const node of items) {
      const children = childNodesOf(node);
      if (children !== undefined && children.length > 0) {
        walk(children, {
          run: node.type === "run" ? node : run,
          removed:
            removed || node.type === "deletion" || node.type === "moveFrom",
          insertion:
            node.type === "insertion" || node.type === "moveTo"
              ? node.info
              : insertion,
        });
        continue;
      }
      const width = leafWidthOf(node);
      const before = { offset: position, zeroWidthBefore: zeroWidthSeen };
      if (width === "zeroWidth") zeroWidthSeen += 1;
      else {
        position += node.type === "text" ? node.text.length : 1;
        zeroWidthSeen = 0;
      }
      leaves.push({
        node,
        width,
        run,
        removed,
        insertion,
        before,
        after: { offset: position, zeroWidthBefore: zeroWidthSeen },
      });
    }
  };
  walk(paragraph.content, {
    run: undefined,
    removed: false,
    insertion: undefined,
  });
  return leaves;
};

/** One paragraph of a story with its leaves. */
type StoryParagraph = {
  paragraph: ReadonlyParagraph;
  paraId: string;
  order: number;
  leaves: readonly ModelLeaf[];
};

/**
 * Paragraphs that accepting every change leaves as one: each but the last has
 * a mark the change removes, and the last keeps its id.
 */
type AcceptedChain = readonly StoryParagraph[];

type StoryReading = {
  /** In the order the editor's text blocks stand. */
  chains: readonly AcceptedChain[];
  /** Every paragraph, rows a change removes included. */
  paragraphs: readonly StoryParagraph[];
};

const isShapeWithText = (node: ReadonlyInline): boolean =>
  node.type === "shape" && node.shape.textBody !== undefined;

const textBoxBlocksOf = (
  leaves: readonly ModelLeaf[],
): readonly DeepReadonly<BlockContent>[] =>
  leaves.flatMap(({ node }) =>
    node.type === "shape" && node.shape.textBody !== undefined
      ? node.shape.textBody.content
      : [],
  );

/**
 * A paragraph holding only text-box anchors stands as its text boxes alone,
 * unless something else of it has to stay: its section or a zero-width marker.
 */
const standsAsItsTextBoxes = (
  paragraph: ReadonlyParagraph,
  leaves: readonly ModelLeaf[],
): boolean =>
  paragraph.sectionProperties === undefined &&
  leaves.some(({ node }) => isShapeWithText(node)) &&
  leaves.every(({ node }) => isShapeWithText(node));

const removesMark = (paragraph: ReadonlyParagraph): boolean =>
  paragraph.pPrMark?.kind === "del" || paragraph.pPrMark?.kind === "moveFrom";

const readStory = (blocks: readonly DeepReadonly<BlockContent>[]) => {
  const chains: AcceptedChain[] = [];
  const paragraphs: StoryParagraph[] = [];
  let pending: StoryParagraph[] = [];
  const visit = (
    list: readonly DeepReadonly<BlockContent>[],
    removed: boolean,
  ): void => {
    for (const [index, block] of list.entries()) {
      switch (block.type) {
        case "paragraph": {
          const paraId =
            block.paraId ??
            panic("A canonical paragraph has no paragraph id", { index });
          const leaves = paragraphLeaves(block);
          const read = {
            paragraph: block,
            paraId,
            order: paragraphs.length,
            leaves,
          };
          paragraphs.push(read);
          if (removed) break;
          const textBoxes = textBoxBlocksOf(leaves);
          if (!standsAsItsTextBoxes(block, leaves)) {
            pending.push(read);
            const next = list[index + 1];
            const joins =
              removesMark(block) &&
              block.sectionProperties === undefined &&
              next?.type === "paragraph";
            if (!joins) {
              chains.push(pending);
              pending = [];
            }
          }
          if (textBoxes.length > 0) visit(textBoxes, false);
          break;
        }
        case "table":
          for (const row of block.rows) {
            const rowRemoved =
              removed || row.structuralChange?.type === "tableRowDeletion";
            for (const cell of row.cells)
              visit(
                cell.content,
                rowRemoved ||
                  cell.structuralChange?.type === "tableCellDeletion",
              );
          }
          break;
        case "blockSdt":
        case "blockCustomXml":
          visit(block.content, removed);
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
          const unreachable: never = block;
          return unreachable;
        }
      }
    }
  };
  visit(blocks, false);
  if (pending.length > 0) chains.push(pending);
  return { chains, paragraphs } satisfies StoryReading;
};

const readStoryOf = ({ document, story }: InlineAtomStory): StoryReading => {
  const body = findStoryBody(readModel(document), story);
  return readStory(body?.content ?? []);
};

const fieldResultRuns = (field: ReadonlyField): readonly ReadonlyRun[] =>
  field.type === "simpleField"
    ? field.content.filter((item) => item.type === "run")
    : field.fieldResult;

/**
 * The run formatting a field's result shows: its first formatted result run,
 * or the field's own when it has no result at all.
 */
const fieldFormattingOf = (
  field: ReadonlyField,
): DeepReadonly<TextFormatting> | undefined => {
  const results = fieldResultRuns(field);
  const first = results.find((run) => run.formatting !== undefined);
  if (first?.formatting !== undefined) return first.formatting;
  return field.type === "complexField" && results.length === 0
    ? field.formatting
    : undefined;
};

const CONTROL_TEXT_BY_RUN_CONTENT = {
  tab: "\t",
  softHyphen: "\u00AD",
  noBreakHyphen: "\u2011",
} as const;

/** What a field's result reads as: its text, tabs and line breaks. */
const fieldResultText = (field: ReadonlyField): string => {
  let text = "";
  for (const run of fieldResultRuns(field)) {
    for (const content of run.content) {
      if (content.type === "text") text += content.text;
      else if (content.type === "break" && content.breakType !== "page")
        text += "\n";
      else if (
        content.type === "tab" ||
        content.type === "softHyphen" ||
        content.type === "noBreakHyphen"
      )
        text += CONTROL_TEXT_BY_RUN_CONTENT[content.type];
    }
  }
  return text;
};

/** A run whose leaves the editor rejoins into one `w:r`, which marks it. */
const runCarriesIdentity = (run: ReadonlyRun): boolean =>
  run.preservedAttributes !== undefined ||
  run.formatting?.preserved !== undefined ||
  (run.formatting !== undefined && Object.keys(run.formatting).length === 0) ||
  run.content.some(
    (content) => content.type === "break" && content.breakType === "page",
  ) ||
  (run.content.length > 1 &&
    run.content.some((content) => content.type !== "text"));

/**
 * Whether a field's result is plain text in one formatting, which the editor
 * holds as one field atom. Any other result is structured content around which
 * a comparison edits as it edits text.
 */
const isPlainField = (field: ReadonlyField): boolean => {
  if (
    field.type === "simpleField" &&
    field.content.some((item) => item.type !== "run")
  )
    return false;
  let formatting: string | undefined;
  for (const run of fieldResultRuns(field)) {
    if (runCarriesIdentity(run)) return false;
    for (const content of run.content) {
      if (content.type !== "text") return false;
      if (content.text === "") continue;
      const own = canonicalJson(run.formatting);
      formatting ??= own;
      if (formatting !== own) return false;
    }
  }
  return true;
};

/** A field's document facts, as the editor's field atom holds them. */
const fieldFacts = (field: ReadonlyField): Record<string, unknown> => {
  const displayText = fieldResultText(field);
  const results = fieldResultRuns(field);
  const emptyResultRuns =
    displayText === "" &&
    results.length > 0 &&
    (field.type === "complexField" ||
      field.content.length === results.length) &&
    results.every((run) =>
      run.content.every((item) => item.type === "text" && item.text === ""),
    )
      ? results
      : undefined;
  const fieldResultIsFallback =
    field.type === "complexField" ? field.fieldResultIsFallback : undefined;
  // OOXML cannot put w:fldSimple inside a revision: a tracked simple field is
  // written as field-character runs, so the two forms are the same field
  // unless the result was synthesized rather than authored.
  return {
    fieldType: field.fieldType,
    instruction: field.instruction,
    displayText,
    fldLock: field.fldLock,
    dirty: field.dirty,
    fieldResultIsFallback,
    formFieldDataXml:
      field.type === "complexField" ? field.formFieldDataXml : undefined,
    emptyResultRuns,
    ...(fieldResultIsFallback === undefined
      ? {}
      : { fieldKind: field.type === "simpleField" ? "simple" : "complex" }),
  };
};

/** The relationship a comparison's picture takes until save gives it its own. */
const COMPARISON_IMAGE_RELATIONSHIP =
  relationshipIdOf("rId_img_compare") ??
  panic("The comparison's picture relationship names nothing.");

/**
 * A picture detached from its package, or `null` when the comparison cannot
 * carry it: no embedded media, a linked picture, or markup it cannot rebind.
 */
const detachedDrawing = (
  drawing: ReadonlyDrawing,
): DeepReadonly<DrawingContent> | null => {
  const { image } = drawing;
  if (
    image.src?.startsWith("data:") !== true ||
    image.hlinkRId !== undefined ||
    image.hlinkHref !== undefined
  )
    return null;
  const rawXml =
    drawing.rawXml === undefined
      ? undefined
      : rebindDrawingImageRelationship({
          xml: drawing.rawXml,
          previousId: image.rId === undefined ? "" : String(image.rId),
          nextId: COMPARISON_IMAGE_RELATIONSHIP,
        });
  if (rawXml === null) return null;
  return {
    ...drawing,
    image: { ...image, rId: COMPARISON_IMAGE_RELATIONSHIP },
    ...(rawXml === undefined ? {} : { rawXml }),
  };
};

const ownFormatting = (run: ReadonlyRun | undefined) =>
  run?.formatting !== undefined && Object.keys(run.formatting).length > 0
    ? run.formatting
    : undefined;

/** A picture's document facts: not the drawing id, which only numbers it. */
const drawingKey = (
  drawing: ReadonlyDrawing,
  run: ReadonlyRun | undefined,
): string =>
  canonicalJson({
    type: "image",
    image: { ...drawing.image, id: undefined },
    rawXml:
      drawing.rawXml === undefined
        ? undefined
        : drawingXmlWithoutIdentity(drawing.rawXml),
    rawXmlMode: drawing.rawXmlMode,
    rawImageFingerprint:
      drawing.rawXmlMode === DRAWING_RAW_XML_MODES.PRESERVE_ONLY
        ? undefined
        : drawing.rawImageFingerprint,
    runFormatting: ownFormatting(run),
  });

/** A run holding one atom of another run, with that run's formatting. */
const runWith = (
  run: ReadonlyRun | undefined,
  content: DeepReadonly<RunContent>,
): DeepReadonly<Run> => ({
  type: "run",
  ...(run?.formatting === undefined ? {} : { formatting: run.formatting }),
  content: [content],
});

/** An atom the comparison restores, or `null` for any other leaf. */
type SupportedAtomReading = {
  key: string;
  resultText: string | null;
  /** The content a revised atom is inserted as. */
  content: DeepReadonly<ParagraphContent>;
};

const supportedAtomOf = (leaf: ModelLeaf): SupportedAtomReading | null => {
  const { node, run } = leaf;
  switch (node.type) {
    case "simpleField":
    case "complexField":
      if (!isPlainField(node)) return null;
      return {
        key: canonicalJson({ type: "field", ...fieldFacts(node) }),
        resultText: fieldResultText(node),
        content: node,
      };
    case "break": {
      if (node.breakType !== "page") return null;
      return {
        key: canonicalJson({ type: "pageBreakRun", clear: node.clear }),
        resultText: null,
        content: runWith(run, node),
      };
    }
    case "drawing": {
      const detached = detachedDrawing(node);
      if (detached === null) return null;
      return {
        key: drawingKey(detached, run),
        resultText: null,
        content: runWith(run, detached),
      };
    }
    default:
      return null;
  }
};

/**
 * The topology key of an inline atom the comparison does not restore, or
 * `null` for text and for what the editor holds as no atom at all.
 */
const unsupportedTopologyKeyOf = (leaf: ModelLeaf): string | null => {
  const { node, run } = leaf;
  switch (node.type) {
    case "text":
    case "renderedPageBreak":
    case "footnoteRef":
    case "endnoteRef":
    case "softHyphen":
    case "noBreakHyphen":
      return null;
    case "simpleField":
    case "complexField":
      return `structuredField:${canonicalJson(fieldFacts(node))}`;
    case "drawing":
      return `image:${drawingKey(node, run)}`;
    case "shape":
      // The editor anchors a text box by its run's own properties alone; its
      // content stands as blocks of their own.
      return node.shape.textBody === undefined
        ? `shape:${canonicalJson({ shape: node, runFormatting: ownFormatting(run) })}`
        : `textBoxAnchor:${canonicalJson(ownFormatting(run))}`;
    case "break":
      return `hardBreak:${canonicalJson(node)}`;
    default:
      return childNodesOf(node) === undefined
        ? `${node.type}:${canonicalJson(node)}`
        : null;
  }
};

const isHidden = (leaf: ModelLeaf): boolean => {
  switch (leaf.node.type) {
    case "simpleField":
    case "complexField":
      return fieldFormattingOf(leaf.node)?.hidden === true;
    default:
      return leaf.run?.formatting?.hidden === true;
  }
};

/** A leaf of an accepted block: its paragraph's place in the chain, and its own there. */
type BlockLeaf = {
  leaf: ModelLeaf;
  paragraph: StoryParagraph;
  paragraphIndex: number;
  leafIndex: number;
};

const pointBefore = ({ leaf, paragraph }: BlockLeaf): LivePoint => ({
  paraId: paragraph.paraId,
  order: paragraph.order,
  ...leaf.before,
});

const pointAfter = ({ leaf, paragraph }: BlockLeaf): LivePoint => ({
  paraId: paragraph.paraId,
  order: paragraph.order,
  ...leaf.after,
});

const paragraphStart = (paragraph: StoryParagraph): LivePoint => ({
  paraId: paragraph.paraId,
  order: paragraph.order,
  offset: 0,
  zeroWidthBefore: 0,
});

type PastRemovedOptions = {
  chain: AcceptedChain;
  /** The chain paragraph and leaf the point stands before. */
  paragraphIndex: number;
  leafIndex: number;
  point: LivePoint;
};

/**
 * A point of the accepted block in the story as it stands, placed as the
 * review's own position mapping places it: past what accepting removes right
 * there, and past the paragraph end accepting joins away.
 */
const pastRemoved = ({
  chain,
  paragraphIndex,
  leafIndex,
  point,
}: PastRemovedOptions): LivePoint => {
  let current = point;
  let paragraphAt = paragraphIndex;
  let leafAt = leafIndex;
  for (;;) {
    const paragraph = chain[paragraphAt];
    if (paragraph === undefined) return current;
    const leaf = paragraph.leaves[leafAt];
    if (leaf === undefined) {
      const next = chain[paragraphAt + 1];
      if (next === undefined) return current;
      paragraphAt += 1;
      leafAt = 0;
      current = paragraphStart(next);
      continue;
    }
    if (!leaf.removed) return current;
    current = pointAfter({
      leaf,
      paragraph,
      paragraphIndex: paragraphAt,
      leafIndex: leafAt,
    });
    leafAt += 1;
  }
};

/** The point just past a leaf of an accepted block. */
const pastLeaf = (chain: AcceptedChain, at: BlockLeaf): LivePoint =>
  pastRemoved({
    chain,
    paragraphIndex: at.paragraphIndex,
    leafIndex: at.leafIndex + 1,
    point: pointAfter(at),
  });

type ModelAtom = AlignableAtom & SupportedAtomReading & { at: BlockLeaf };

/** A footnote or endnote reference read as its marker, over `length` clean-text characters. */
type NoteReferenceAt = NoteReferenceIdentity & {
  offset: number;
  length: number;
  reference: DeepReadonly<RunContent>;
};

type FieldResults = "text" | "omitted";

/**
 * How a block's note references read. `"sourceText"` is the editor's own text
 * for a reference (its package id); labels read it as the marker (`[^1]`) the
 * comparison planned in.
 */
type NoteReferenceReading = NoteReferenceLabels | "sourceText";

/** One block as accepting every change leaves it, read for alignment. */
type ModelBlock = {
  blockType: string;
  chain: AcceptedChain;
  cleanText: string;
  /** Where each clean-text character stands in the story, then the block's end. */
  offsets: readonly LivePoint[];
  /** The leaf each clean-text character belongs to. */
  characters: readonly { at: BlockLeaf; last: boolean }[];
  supported: readonly ModelAtom[];
  unsupportedTopology: readonly { offset: number; key: string }[];
  /** Empty unless the block was read with note reference labels. */
  noteReferences: readonly NoteReferenceAt[];
  /** The paragraph that keeps its id once the block is accepted. */
  survivor: StoryParagraph;
};

/** The point just past clean-text character `index` of a block. */
const pastCharacter = (
  block: ModelBlock,
  index: number,
): LivePoint | undefined => {
  const character = block.characters[index];
  const before = block.offsets[index];
  if (character === undefined || before === undefined) return undefined;
  return character.last
    ? pastLeaf(block.chain, character.at)
    : { ...before, offset: before.offset + 1, zeroWidthBefore: 0 };
};

type ReadBlockOptions = {
  chain: AcceptedChain;
  fieldResults: FieldResults;
  noteReferences: NoteReferenceReading;
};

/** The editor's clean text of an accepted block, and its atoms, on the model. */
const readBlock = ({
  chain,
  fieldResults,
  noteReferences,
}: ReadBlockOptions): ModelBlock => {
  const survivor = chain.at(-1) ?? panic("An accepted block has a paragraph");
  let text = "";
  const offsets: LivePoint[] = [];
  const characters: { at: BlockLeaf; last: boolean }[] = [];
  const supported: ModelAtom[] = [];
  const unsupportedTopology: { offset: number; key: string }[] = [];
  const references: NoteReferenceAt[] = [];
  const labels = noteReferences === "sourceText" ? undefined : noteReferences;
  let localNumbering: NoteReferenceNumbering | undefined;
  let lastContributing: BlockLeaf | undefined;
  /** Characters a unit shows: each anchors at the unit, which has no inside. */
  const append = (shown: string, at: BlockLeaf) => {
    if (shown === "") return;
    const before = pointBefore(at);
    for (let index = 0; index < shown.length; index++) {
      offsets.push(before);
      characters.push({ at, last: index === shown.length - 1 });
    }
    text += shown;
    lastContributing = at;
  };
  for (const [paragraphIndex, paragraph] of chain.entries()) {
    for (const [leafIndex, leaf] of paragraph.leaves.entries()) {
      if (leaf.removed) continue;
      const at = { leaf, paragraph, paragraphIndex, leafIndex };
      const { node } = leaf;
      const atom = supportedAtomOf(leaf);
      if (atom !== null) supported.push({ ...atom, offset: text.length, at });
      else {
        const key = unsupportedTopologyKeyOf(leaf);
        if (key !== null)
          unsupportedTopology.push({ offset: text.length, key });
      }
      if (isHidden(leaf)) continue;
      switch (node.type) {
        case "text": {
          const before = pointBefore(at);
          for (let index = 0; index < node.text.length; index++) {
            offsets.push(
              index === 0
                ? before
                : {
                    ...before,
                    offset: before.offset + index,
                    zeroWidthBefore: 0,
                  },
            );
            characters.push({ at, last: index === node.text.length - 1 });
          }
          text += node.text;
          lastContributing = at;
          break;
        }
        case "footnoteRef":
        case "endnoteRef": {
          const reference = {
            noteType: node.type === "endnoteRef" ? "endnote" : "footnote",
            noteId: String(node.id),
          } as const satisfies NoteReferenceIdentity;
          if (labels === undefined) {
            append(reference.noteId, at);
            break;
          }
          const label =
            labels.labelOf(reference) ??
            noteReferenceMarker(
              reference.noteType,
              (localNumbering ??= createNoteReferenceNumbering()).next(
                reference.noteType,
                reference.noteId,
              ).displayNumber,
            );
          references.push({
            ...reference,
            offset: text.length,
            length: label.length,
            reference: node,
          });
          append(label, at);
          break;
        }
        case "tab":
        case "softHyphen":
        case "noBreakHyphen":
          append(CONTROL_TEXT_BY_RUN_CONTENT[node.type], at);
          break;
        case "break":
          if (node.breakType !== "page") append("\n", at);
          break;
        case "simpleField":
        case "complexField":
          // A plain field's result is its atom's text, which an alignment
          // coordinate leaves out; a structured one reads as its text.
          if (atom === null || fieldResults === "text")
            append(fieldResultText(node), at);
          break;
        default:
          break;
      }
    }
  }
  const first = chain.at(0) ?? survivor;
  offsets.push(
    lastContributing === undefined
      ? pastRemoved({
          chain,
          paragraphIndex: 0,
          leafIndex: 0,
          point: paragraphStart(first),
        })
      : pastLeaf(chain, lastContributing),
  );
  return {
    blockType: "paragraph",
    chain,
    cleanText: text,
    offsets,
    characters,
    supported,
    unsupportedTopology,
    noteReferences: references,
    survivor,
  };
};

const readBlocks = (
  reading: StoryReading,
  fieldResults: FieldResults,
  noteReferences: NoteReferenceReading,
): ModelBlock[] =>
  reading.chains.map((chain) =>
    readBlock({ chain, fieldResults, noteReferences }),
  );

const noteReferenceKey = ({
  noteType,
  noteId,
}: NoteReferenceIdentity): string => `${noteType}:${noteId}`;

const allLeaves = (reading: StoryReading): ModelLeaf[] =>
  reading.paragraphs.flatMap(({ leaves }) => leaves);

/** Every note a story references, tracked-deleted references included. */
const referencedNotesOf = (reading: StoryReading): ReadonlySet<string> => {
  const notes = new Set<string>();
  for (const { node } of allLeaves(reading))
    if (node.type === "footnoteRef" || node.type === "endnoteRef")
      notes.add(
        noteReferenceKey({
          noteType: node.type === "endnoteRef" ? "endnote" : "footnote",
          noteId: String(node.id),
        }),
      );
  return notes;
};

const hasNoteReference = (reading: StoryReading): boolean =>
  allLeaves(reading).some(
    ({ node }) => node.type === "footnoteRef" || node.type === "endnoteRef",
  );

/** Whether a story holds an atom of a kind the comparison restores. */
const hasSupportedAtom = (reading: StoryReading): boolean =>
  allLeaves(reading).some(({ node }) => {
    switch (node.type) {
      case "simpleField":
      case "complexField":
        return isPlainField(node);
      case "break":
        return node.breakType === "page";
      case "drawing":
        return true;
      default:
        return false;
    }
  });

/**
 * Whether this comparison inserted text that may stand for a note reference.
 * Operations carry clean text, where a reference reads as its marker, so an
 * operation that brings a reference writes the marker as text.
 */
const wroteNoteReferenceMarker = (
  reading: StoryReading,
  originalRevisionIdSeed: number,
): boolean =>
  allLeaves(reading).some(
    ({ node, insertion }) =>
      node.type === "text" &&
      node.text.includes("[^") &&
      insertion !== undefined &&
      insertion.id >= originalRevisionIdSeed,
  );

/** Leaves of one paragraph that overlap `[from, to)`. */
const leavesBetween = (
  paragraph: StoryParagraph,
  from: number,
  to: number,
): ModelLeaf[] =>
  paragraph.leaves.filter(
    ({ before, after, width }) =>
      (width === "zeroWidth" && before.offset > from && before.offset < to) ||
      (width !== "zeroWidth" && before.offset < to && after.offset > from),
  );

type TextRangeDispositionOptions = {
  paragraphs: ReadonlyMap<string, StoryParagraph>;
  from: LivePoint;
  to: LivePoint;
  expectedText: string;
  originalRevisionIdSeed: number;
};

/**
 * Whether `[from, to)` is exactly `expectedText`, all of it this comparison's
 * own insertion or all of it retained content, and nothing but text.
 */
const textRangeDisposition = ({
  paragraphs,
  from,
  to,
  expectedText,
  originalRevisionIdSeed,
}: TextRangeDispositionOptions): AtomActionDisposition | null => {
  if (from.paraId !== to.paraId || from.offset >= to.offset) return null;
  const paragraph = paragraphs.get(from.paraId);
  if (paragraph === undefined) return null;
  let text = "";
  let disposition: AtomActionDisposition | undefined;
  for (const leaf of leavesBetween(paragraph, from.offset, to.offset)) {
    if (leaf.node.type !== "text" || leaf.removed) return null;
    const start = Math.max(from.offset, leaf.before.offset);
    const end = Math.min(to.offset, leaf.after.offset);
    text += leaf.node.text.slice(
      start - leaf.before.offset,
      end - leaf.before.offset,
    );
    let current: AtomActionDisposition = "retained";
    if (leaf.insertion !== undefined) {
      if (leaf.insertion.id < originalRevisionIdSeed) return null;
      current = "inserted";
    }
    if (disposition !== undefined && disposition !== current) return null;
    disposition = current;
  }
  return text === expectedText ? (disposition ?? null) : null;
};

const atomDisposition = (
  leaf: ModelLeaf,
  originalRevisionIdSeed: number,
): AtomActionDisposition | null => {
  if (leaf.removed) return null;
  if (leaf.insertion === undefined) return "retained";
  return leaf.insertion.id >= originalRevisionIdSeed ? "inserted" : null;
};

type InsertAction = {
  kind: "insert";
  from: LivePoint;
  content: DeepReadonly<ParagraphContent>;
  targetBlockId: string | undefined;
};
type DeleteAction = {
  kind: "delete";
  from: LivePoint;
  to: LivePoint;
  disposition: AtomActionDisposition;
  targetBlockId: string | undefined;
};
type ReplaceTextAction = {
  kind: "replace-text";
  from: LivePoint;
  to: LivePoint;
  content: DeepReadonly<ParagraphContent>;
  textDisposition: AtomActionDisposition;
  targetBlockId: string | undefined;
};
/**
 * Turn the marker text a comparison operation wrote (`[^1]`) back into the
 * note reference the target holds there, inside that same insertion.
 */
type NoteReferenceAction = {
  kind: "note-reference";
  from: LivePoint;
  to: LivePoint;
  content: DeepReadonly<ParagraphContent>;
  insertion: DeepReadonly<TrackedChangeInfo>;
  targetBlockId: string | undefined;
};
type ReplaceAtomAction = {
  kind: "replace-atom";
  from: LivePoint;
  to: LivePoint;
  text: string;
  atomDisposition: AtomActionDisposition;
  formatting: DeepReadonly<TextFormatting> | undefined;
  targetBlockId: string | undefined;
};
type Action =
  | InsertAction
  | DeleteAction
  | ReplaceTextAction
  | ReplaceAtomAction
  | NoteReferenceAction;

const actionRangeCount = (action: Action): number => {
  switch (action.kind) {
    case "replace-text":
      return atomActionRangeCount(action.textDisposition);
    case "replace-atom":
      return atomActionRangeCount(action.atomDisposition);
    case "insert":
    case "delete":
    case "note-reference":
      return atomActionRangeCount(null);
    default: {
      const unreachable: never = action;
      return unreachable;
    }
  }
};

/** A revised block's id in the snapshot the comparison planned against. */
const targetBlockIdOf = (
  snapshot: FolioAIEditSnapshot,
  block: ModelBlock,
): string | undefined =>
  Object.hasOwn(snapshot.anchors, block.survivor.paraId)
    ? block.survivor.paraId
    : undefined;

/** One edit already in the batch, in its paragraph's offsets. */
type AppliedEdit = { offset: number; deleted: number; inserted: number };

/**
 * Where a point of the story as it stood lands after the batch's earlier
 * edits, with the review's affinity: `1` past what was inserted there, `-1`
 * before it.
 */
const mapPoint = (
  edits: ReadonlyMap<string, readonly AppliedEdit[]>,
  point: LivePoint,
  assoc: 1 | -1,
): LivePoint => {
  let { offset, zeroWidthBefore } = point;
  for (const edit of edits.get(point.paraId) ?? []) {
    const end = edit.offset + edit.deleted;
    if (offset < edit.offset) continue;
    if (offset > end) {
      offset += edit.inserted - edit.deleted;
      continue;
    }
    let side: 1 | -1 = assoc;
    if (edit.deleted > 0 && offset === edit.offset) side = -1;
    else if (edit.deleted > 0 && offset === end) side = 1;
    const mapped = side < 0 ? edit.offset : edit.offset + edit.inserted;
    if (mapped !== offset) zeroWidthBefore = 0;
    offset = mapped;
  }
  return { ...point, offset, zeroWidthBefore };
};

/**
 * Restore field, picture and page-break atoms omitted by text-only comparison
 * operations, and the note references whose markers they wrote as text, as
 * document operations on the reviewed story.
 */
export const matchInlineAtomsOnModel = ({
  document,
  story,
  target,
  revisionStamp,
  originalRevisionIdSeed,
  author,
  maxRanges,
  noteReferenceLabels,
  baseNotes,
}: MatchInlineAtomsOnModelOptions): MatchInlineAtomsOnModelResult => {
  if (!Number.isSafeInteger(maxRanges) || maxRanges < 0)
    return { status: "budget-exceeded" };
  const live = readStoryOf({ document, story });
  const revised = readStoryOf(target);
  // An operation that brings a note reference writes its marker, which is
  // the reference's clean text. Reading both sides' references with the
  // markers the plan was written in lines that text up with the reference
  // it stands for.
  const restoresNoteReferences =
    wroteNoteReferenceMarker(live, originalRevisionIdSeed) &&
    hasNoteReference(revised);
  const noteReferences: NoteReferenceReading = restoresNoteReferences
    ? noteReferenceLabels
    : "sourceText";
  const unchanged = {
    status: "matched",
    operations: [],
    nextRevisionId: revisionStamp.idSeed,
    changedTargetBlockIds: [],
    rangeCount: 0,
  } as const;
  if (
    !restoresNoteReferences &&
    !hasSupportedAtom(live) &&
    !hasSupportedAtom(revised)
  )
    return unchanged;

  const liveBlocks = readBlocks(live, "text", noteReferences);
  const targetBlocks = readBlocks(revised, "text", noteReferences);
  if (liveBlocks.length !== targetBlocks.length)
    return { status: "unalignable" };
  const paragraphs = new Map(
    live.paragraphs.map((paragraph) => [paragraph.paraId, paragraph] as const),
  );

  const actions: Action[] = [];
  let rangeCount = 0;
  const planAction = (action: Action): boolean => {
    const nextRangeCount = rangeCount + actionRangeCount(action);
    if (nextRangeCount > maxRanges) return false;
    rangeCount = nextRangeCount;
    actions.push(action);
    return true;
  };
  let referencedNotes: ReadonlySet<string> | undefined;
  let omittedLiveBlocks: readonly ModelBlock[] | undefined;
  let omittedTargetBlocks: readonly ModelBlock[] | undefined;
  for (const [index, fullLive] of liveBlocks.entries()) {
    const fullTarget = targetBlocks[index];
    if (!fullTarget) return { status: "unalignable" };
    const targetBlockId = targetBlockIdOf(target.snapshot, fullTarget);
    if (
      (fullLive.noteReferences.length > 0 ||
        fullTarget.noteReferences.length > 0) &&
      fullLive.cleanText === fullTarget.cleanText
    ) {
      const kept = new Map(
        fullLive.noteReferences.map((reference) => [
          reference.offset,
          reference,
        ]),
      );
      for (const reference of fullTarget.noteReferences) {
        const liveReference = kept.get(reference.offset);
        if (liveReference) {
          kept.delete(reference.offset);
          if (
            noteReferenceKey(liveReference) !== noteReferenceKey(reference) ||
            liveReference.length !== reference.length
          )
            return { status: "unalignable" };
          continue;
        }
        // The marker is text this comparison wrote. It may only stand for a
        // note the base package defines: a reference to one it does not have
        // would point at nothing in the redline.
        referencedNotes ??= baseNotes ?? referencedNotesOf(live);
        if (!referencedNotes.has(noteReferenceKey(reference)))
          return { status: "unalignable", detail: MISSING_NOTE_DETAIL };
        const from = fullLive.offsets[reference.offset];
        const to = pastCharacter(
          fullLive,
          reference.offset + reference.length - 1,
        );
        const written = fullLive.characters[reference.offset]?.at.leaf;
        if (from === undefined || to === undefined || written === undefined)
          return { status: "unalignable" };
        const disposition = textRangeDisposition({
          paragraphs,
          from,
          to,
          expectedText: fullTarget.cleanText.slice(
            reference.offset,
            reference.offset + reference.length,
          ),
          originalRevisionIdSeed,
        });
        if (disposition !== "inserted" || written.insertion === undefined)
          return { status: "unalignable" };
        if (
          !planAction({
            kind: "note-reference",
            from,
            to,
            // The marker's own run properties, as the comparison wrote them.
            content: runWith(written.run, reference.reference),
            insertion: written.insertion,
            targetBlockId,
          })
        )
          return { status: "budget-exceeded" };
      }
      // A reference the reviewed story keeps where the target has none.
      if (kept.size > 0) return { status: "unalignable" };
    }
    if (fullLive.supported.length === 0 && fullTarget.supported.length === 0)
      continue;
    // An unsupported zero-width neighbor cannot make an unchanged supported
    // carrier need reconciliation. Package-local comment ids are the common
    // case: they may be rebound independently while the page break beside
    // them remains identical.
    if (sameSupportedAtoms(fullLive, fullTarget)) continue;
    let liveBlock = fullLive;
    let targetBlock = fullTarget;
    let fieldResults: FieldResults = "text";
    if (!sameBlockTopology(liveBlock, targetBlock)) {
      omittedLiveBlocks ??= readBlocks(live, "omitted", noteReferences);
      omittedTargetBlocks ??= readBlocks(revised, "omitted", noteReferences);
      const omittedLive = omittedLiveBlocks[index];
      const omittedTarget = omittedTargetBlocks[index];
      if (
        !omittedLive ||
        !omittedTarget ||
        !sameBlockTopology(omittedLive, omittedTarget)
      )
        return { status: "unalignable" };
      liveBlock = omittedLive;
      targetBlock = omittedTarget;
      fieldResults = "omitted";
    }
    const matches = matchingAtoms({
      live: liveBlock.supported,
      target: targetBlock.supported,
    });
    const matchedLive = new Set(matches.map(([liveIndex]) => liveIndex));
    const matchedTarget = new Set(
      matches.map(([, targetIndex]) => targetIndex),
    );
    const anchors = insertionAnchors({
      live: liveBlock.supported,
      target: targetBlock.supported,
      matches,
    });
    const liveFieldResultOwnership = fieldResultOwnershipOf(
      liveBlock.supported,
    );
    const targetFieldResultOwnership = fieldResultOwnershipOf(
      targetBlock.supported,
    );
    for (const [liveIndex, atom] of liveBlock.supported.entries()) {
      if (matchedLive.has(liveIndex)) continue;
      const from = pointBefore(atom.at);
      const to = pointAfter(atom.at);
      const { resultText } = atom;
      if (
        fieldResults === "text" &&
        resultText &&
        !ownsFieldResult({
          ownership: targetFieldResultOwnership,
          offset: atom.offset,
          resultText,
        })
      ) {
        const disposition = atomDisposition(
          atom.at.leaf,
          originalRevisionIdSeed,
        );
        if (disposition === null) return { status: "unalignable" };
        const field = atom.at.leaf.node;
        if (
          !planAction({
            kind: "replace-atom",
            from,
            to,
            text: resultText,
            atomDisposition: disposition,
            formatting:
              field.type === "simpleField" || field.type === "complexField"
                ? fieldFormattingOf(field)
                : undefined,
            targetBlockId,
          })
        )
          return { status: "budget-exceeded" };
        continue;
      }
      const disposition = atomDisposition(atom.at.leaf, originalRevisionIdSeed);
      if (disposition === null) return { status: "unalignable" };
      if (!planAction({ kind: "delete", from, to, disposition, targetBlockId }))
        return { status: "budget-exceeded" };
    }
    for (const [targetIndex, atom] of targetBlock.supported.entries()) {
      if (matchedTarget.has(targetIndex)) continue;
      const anchor = anchors.get(targetIndex);
      const anchorAtom =
        anchor === undefined
          ? undefined
          : liveBlock.supported[anchor.liveIndex];
      let from: LivePoint | undefined;
      if (anchor === undefined || anchorAtom === undefined)
        from = liveBlock.offsets[atom.offset];
      else
        from =
          anchor.side === "before"
            ? pointBefore(anchorAtom.at)
            : pastLeaf(liveBlock.chain, anchorAtom.at);
      if (from === undefined) return { status: "unalignable" };
      const { resultText } = atom;
      if (
        fieldResults === "text" &&
        resultText &&
        !ownsFieldResult({
          ownership: liveFieldResultOwnership,
          offset: atom.offset,
          resultText,
        })
      ) {
        const to = liveBlock.offsets[atom.offset + resultText.length];
        if (to === undefined) return { status: "unalignable" };
        const textDisposition = textRangeDisposition({
          paragraphs,
          from,
          to,
          expectedText: resultText,
          originalRevisionIdSeed,
        });
        if (textDisposition === null) return { status: "unalignable" };
        if (
          !planAction({
            kind: "replace-text",
            from,
            to,
            content: atom.content,
            textDisposition,
            targetBlockId,
          })
        )
          return { status: "budget-exceeded" };
        continue;
      }
      if (
        !planAction({
          kind: "insert",
          from,
          content: atom.content,
          targetBlockId,
        })
      )
        return { status: "budget-exceeded" };
    }
  }
  return writeActions({
    document,
    actions,
    story,
    author,
    revisionStamp,
    rangeCount,
  });
};

type WriteActionsOptions = {
  document: ReadonlyDocument;
  actions: readonly Action[];
  story: OpStory;
  author: string;
  revisionStamp: FolioRevisionStamp;
  rangeCount: number;
};

const PICTURE_DATA_URL =
  /^data:(image\/[a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+)$/iu;

/** The picture of a revised atom, when the atom inserts one. */
const insertedPictureOf = (
  content: DeepReadonly<ParagraphContent>,
): DeepReadonly<DrawingContent> | undefined => {
  if (content.type !== "run") return undefined;
  const drawing = content.content.at(0);
  return drawing?.type === "drawing" &&
    drawing.image.rId === COMPARISON_IMAGE_RELATIONSHIP
    ? drawing
    : undefined;
};

/**
 * The batch that carries the planned actions out, from the end of the story
 * back, each edit addressed where the earlier ones left its point. Every
 * operation is applied to a copy as it is written, so the batch names the
 * ids the records it cuts take, and a picture brings its media and
 * relationship ahead of it.
 */
const writeActions = ({
  document,
  actions,
  story,
  author,
  revisionStamp,
  rangeCount,
}: WriteActionsOptions): MatchInlineAtomsOnModelResult => {
  const operations: DocumentOp[] = [];
  let current = readModel(document);
  const edits = new Map<string, AppliedEdit[]>();
  const changedTargetBlockIds = new Set<string>();
  let nextRevisionId = revisionStamp.idSeed;
  let refusal: string | undefined;
  const stamp = (): RevisionStamp => ({
    id: nextRevisionId++,
    author,
    date: revisionStamp.date,
  });
  const position = (point: LivePoint): TextPosition => ({
    story,
    blockId: point.paraId,
    offset: point.offset,
    zeroWidthBefore: point.zeroWidthBefore,
  });
  const record = (point: LivePoint, edit: AppliedEdit) => {
    const list = edits.get(point.paraId);
    if (list === undefined) edits.set(point.paraId, [edit]);
    else list.push(edit);
  };
  /** Apply one operation to the copy and keep it, or note why it cannot apply. */
  const emit = (op: DocumentOp): boolean => {
    if (refusal !== undefined) return false;
    const applied = applyDocumentOp(current, op);
    if (applied.isErr()) {
      refusal = applied.error.message;
      return false;
    }
    current = applied.value.document;
    operations.push(op);
    return true;
  };
  /** The ids the records an edit cuts in two take, after its stamp's. */
  const freshIds = (
    op: InsertContentOp | DeleteRangeOp,
  ): NewIds | undefined => {
    const demand = revisionIdDemand(current, op);
    if (demand.isErr()) {
      refusal ??= demand.error.message;
      return undefined;
    }
    if (demand.value === 0) return undefined;
    return {
      revision: Array.from({ length: demand.value }, () => nextRevisionId++),
    };
  };
  /** A picture's media and relationship, added to the package for it. */
  const withPictureResources = (
    picture: DeepReadonly<DrawingContent>,
  ): DeepReadonly<DrawingContent> | undefined => {
    const { src } = picture.image;
    const data = src === undefined ? null : PICTURE_DATA_URL.exec(src);
    const mimeType = data?.at(1);
    const payload = data?.at(2);
    if (src === undefined || mimeType === undefined || payload === undefined) {
      refusal ??= "A restored picture has no embedded image bytes.";
      return undefined;
    }
    const relationships = new Map(current.package.relationships);
    let relationshipIndex = relationships.size + 1;
    while (relationships.has(`rId${String(relationshipIndex)}`))
      relationshipIndex += 1;
    const id = relationshipIdOf(`rId${String(relationshipIndex)}`);
    const media = new Map(current.package.media);
    const extension =
      mimeType.toLowerCase() === "image/jpeg" ? "jpg" : mimeType.slice(6);
    let mediaIndex = media.size + 1;
    while (media.has(`word/media/compare${String(mediaIndex)}.${extension}`))
      mediaIndex += 1;
    const path = `word/media/compare${String(mediaIndex)}.${extension}`;
    if (id === undefined) {
      refusal ??= "A restored picture's relationship id names nothing.";
      return undefined;
    }
    media.set(path, {
      path,
      mimeType,
      data: base64ToBytes(payload).slice().buffer,
      dataUrl: src,
    });
    relationships.set(id, {
      id,
      type: RELATIONSHIP_TYPES.image,
      target: path.slice("word/".length),
    });
    const rawXml =
      picture.rawXml === undefined
        ? undefined
        : rebindDrawingImageRelationship({
            xml: picture.rawXml,
            previousId: COMPARISON_IMAGE_RELATIONSHIP,
            nextId: id,
          });
    if (rawXml === null) {
      refusal ??= "A restored picture's markup names no single image.";
      return undefined;
    }
    if (
      !emit(
        packageResourcesOpOf({
          before: current,
          after: {
            ...current,
            package: { ...current.package, relationships, media },
          },
        }),
      )
    )
      return undefined;
    // A drawing id from another package does not number a drawing here.
    const { id: _drawingId, ...image } = picture.image;
    return {
      ...picture,
      image: { ...image, rId: id },
      ...(rawXml === undefined ? {} : { rawXml }),
    };
  };
  const insert = (
    at: LivePoint,
    content: DeepReadonly<ParagraphContent>,
    revision: RevisionStamp,
  ): boolean => {
    const picture = insertedPictureOf(content);
    let inserted = content;
    if (picture !== undefined && content.type === "run") {
      const placed = withPictureResources(picture);
      if (placed === undefined) return false;
      inserted = { ...content, content: [placed] };
    }
    const op: InsertContentOp = {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: position(at),
      slice: {
        content: [writableCopy<ParagraphContent>(inserted)],
        openStart: 0,
        openEnd: 0,
      },
      revision,
    };
    const newIds = freshIds(op);
    if (refusal !== undefined) return false;
    if (!emit(newIds === undefined ? op : { ...op, newIds })) return false;
    record(at, { offset: at.offset, deleted: 0, inserted: 1 });
    return true;
  };
  /** Remove `[from, to)`: the comparison's own insertion goes, anything else is marked deleted. */
  const remove = (
    from: LivePoint,
    to: LivePoint,
    disposition: AtomActionDisposition,
  ): boolean => {
    switch (disposition) {
      case "inserted": {
        const op: DeleteRangeOp = {
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: position(from),
          to: position(to),
        };
        if (!emit(op)) return false;
        record(from, {
          offset: from.offset,
          deleted: to.offset - from.offset,
          inserted: 0,
        });
        return true;
      }
      case "retained": {
        const op: DeleteRangeOp = {
          type: DOCUMENT_OP_TYPES.DELETE_RANGE,
          from: position(from),
          to: position(to),
          revision: stamp(),
        };
        const newIds = freshIds(op);
        if (refusal !== undefined) return false;
        return emit(newIds === undefined ? op : { ...op, newIds });
      }
      default: {
        const unreachable: never = disposition;
        return unreachable;
      }
    }
  };
  const ordered = actions.toSorted(
    (left, right) =>
      comparePoints(right.from, left.from) || (left.kind === "delete" ? -1 : 1),
  );
  for (const action of ordered) {
    let written: boolean;
    switch (action.kind) {
      case "insert":
        written = insert(
          mapPoint(edits, action.from, 1),
          action.content,
          stamp(),
        );
        break;
      case "replace-text":
      case "replace-atom": {
        const from = mapPoint(edits, action.from, 1);
        const to = mapPoint(edits, action.to, -1);
        if (from.paraId !== to.paraId || from.offset >= to.offset)
          return { status: "unalignable" };
        written =
          remove(
            from,
            to,
            action.kind === "replace-text"
              ? action.textDisposition
              : action.atomDisposition,
          ) &&
          insert(
            mapPoint(edits, action.from, -1),
            action.kind === "replace-text"
              ? action.content
              : {
                  type: "run",
                  ...(action.formatting === undefined
                    ? {}
                    : { formatting: action.formatting }),
                  content: [{ type: "text", text: action.text }],
                },
            stamp(),
          );
        break;
      }
      case "note-reference": {
        const from = mapPoint(edits, action.from, 1);
        const to = mapPoint(edits, action.to, -1);
        if (from.paraId !== to.paraId || from.offset >= to.offset)
          return { status: "unalignable" };
        // The reference takes the place of the marker inside the insertion
        // that wrote it, alike in author and date.
        written =
          insert(from, action.content, {
            id: nextRevisionId++,
            author: action.insertion.author,
            date: action.insertion.date ?? revisionStamp.date,
          }) &&
          remove(
            { ...from, offset: from.offset + 1, zeroWidthBefore: 0 },
            { ...to, offset: to.offset + 1 },
            "inserted",
          );
        break;
      }
      case "delete": {
        const from = mapPoint(edits, action.from, 1);
        const to = mapPoint(edits, action.to, -1);
        if (from.paraId !== to.paraId || to.offset - from.offset !== 1)
          return { status: "unalignable" };
        written = remove(from, to, action.disposition);
        break;
      }
      default: {
        const unreachable: never = action;
        return unreachable;
      }
    }
    if (!written)
      return {
        status: "unalignable",
        detail: `inline atoms cannot be restored as document operations: ${refusal ?? "an operation was refused"}`,
      };
    if (action.targetBlockId) changedTargetBlockIds.add(action.targetBlockId);
  }
  return {
    status: "matched",
    operations,
    nextRevisionId,
    changedTargetBlockIds: [...changedTargetBlockIds],
    rangeCount,
  };
};
