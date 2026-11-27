import { panic, Result } from "better-result";
import { Fragment, type Node as ProseNode, Slice } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { CellSelection, tableNodeTypes } from "prosemirror-tables";

import {
  DRAWING_RAW_XML_MODES,
  NUMBER_FORMATS,
  paragraphNumberingReference,
  relationshipIdOf,
} from "@stll/docx-core/model";
import {
  type DocumentOp,
  type EditorIntent,
  OP_STORIES,
  type OpStory,
  packageResourcesOpOf,
  type PasteAnnotations,
  type PastedBlock,
  type TextPosition,
} from "@stll/docx-core/ops";

import { groundForDocumentOpRefusal } from "../canonical/canonicalRefusalGrounds";
import {
  type CanonicalCommit,
  type CanonicalSession,
  CanonicalSessionError,
} from "../canonical/canonicalSession";
import {
  storyPartPath,
  storyRelationshipsKey,
  withRelationshipIds,
} from "../docx/copiedPackageParts";
import {
  takeDrawingId,
  usedDrawingIds,
  withDrawingId,
} from "../docx/drawingIdNormalization";
import { imageRawXmlFingerprint } from "../docx/imageRawXml";
import {
  createNumberingIdAllocator,
  numberingNumIdsTaken,
} from "../docx/numberingIds";
import { copyParagraphIndentationProjection } from "../docx/paragraphPropertySource";
import {
  type InlineContentSlot,
  visitInlineContentSlots,
  visitParagraphRuns,
  withInlineContentRewritten,
} from "../docx/paragraphTraversal";
import { captureVerbatimXml } from "../docx/verbatimCapture";
import {
  findAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  getTextContent,
  OFFICE_RELATIONSHIP_NAMESPACE_URIS,
  OOXML_NAMESPACE_SCOPE,
  parseXml,
  type XmlElement,
} from "../docx/xmlParser";
import { canonicalSelectionRange } from "../prosemirror/canonicalSelectionRange";
import { proseDocToBlocks } from "../prosemirror/conversion/fromProseDoc";
import { completeNumberingForDoc } from "../prosemirror/listInstanceReferences";
import { marksToTextFormatting } from "../prosemirror/runFormattingFromMarks";
import { pastedCells } from "../prosemirror/tableCellPaste";
import {
  CANONICAL_CONDITION,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";
import type {
  BlockContent,
  CounterFormat,
  Document,
  DrawingContent,
  NumberFormat,
  NumberingDefinitions,
  Paragraph,
  ParagraphContent,
  Relationship,
  Run,
  Table,
} from "../types/document";
import { readModel } from "../types/documentModel";
import { bytesToDataUrl } from "../utils/base64";
import { base64ToBytes } from "../utils/textEncoding";
import {
  type ClipboardPackageParts,
  createPartCopier,
  storyRelationshipsOf,
} from "./canonicalClipboardParts";
import {
  flattenClipboardStyleReferences,
  importClipboardStyles,
} from "./canonicalClipboardResources";

/** A refusal under the gap, rule or condition it belongs to. */
const refuse = (message: string, ground: CanonicalRefusalGround) =>
  Result.err(new CanonicalSessionError({ ground, message, reason: "refused" }));
/** Clipboard content the paste cannot read or resolve as given. */
const INVALID_CLIPBOARD: CanonicalRefusalGround = {
  condition: CANONICAL_CONDITION.invalidRequest,
};

/**
 * The ground of a clipboard block the paste cannot place: every block kind
 * the editor shows pastes, so one it cannot place is a slice it cannot read.
 */
const UNPLACED_BLOCK: CanonicalRefusalGround = INVALID_CLIPBOARD;

const tableRole = (node: ProseNode | null): unknown =>
  node?.type.spec["tableRole"];

/**
 * A table the clipboard holds part of, opened `left` levels deep at its start
 * and `right` at its end, as the table of the cells it holds: a first or last
 * row cut short keeps the cells copied and is filled out with empty ones.
 * Undefined when the part holds no cells.
 */
const closedTable = (
  table: ProseNode,
  left: number,
  right: number,
): ProseNode | undefined => {
  // Opened one level, the table holds whole rows.
  if (left <= 1 && right <= 1) return table;
  const cells = pastedCells(new Slice(Fragment.from(table), left, right));
  if (cells === null) return undefined;
  const rowType = tableNodeTypes(table.type.schema).row;
  return table.type.create(
    table.attrs,
    cells.rows.map((row, index) =>
      rowType.create(
        index < table.childCount ? table.child(index).attrs : null,
        row,
      ),
    ),
    table.marks,
  );
};

/**
 * A slice that opens or closes inside a table, its edges closed at the table:
 * it pastes the table of the cells it holds, as a selection of cells pastes.
 * Selected cells, rows or cells without their table, paste as a table
 * of their own. The slice is returned as given when no table edge is open,
 * and undefined when an open table edge holds no cells.
 */
const withClosedTableEdges = (slice: Slice): Slice | undefined => {
  const { content } = slice;
  const first = content.firstChild;
  const last = content.lastChild;
  const role = tableRole(first);
  if (role === "row" || role === "cell" || role === "header_cell") {
    if (first === null) return slice;
    const cells = pastedCells(slice);
    if (cells === null) return undefined;
    const { table, row } = tableNodeTypes(first.type.schema);
    return new Slice(
      Fragment.from(
        table.create(
          null,
          cells.rows.map((cellsOfRow) => row.create(null, cellsOfRow)),
        ),
      ),
      0,
      0,
    );
  }
  const opensTable = slice.openStart > 1 && tableRole(first) === "table";
  const closesTable = slice.openEnd > 1 && tableRole(last) === "table";
  if (!opensTable && !closesTable) return slice;
  const blocks: ProseNode[] = [];
  content.forEach((node) => blocks.push(node));
  const lastIndex = blocks.length - 1;
  const edges = new Set<number>();
  if (opensTable) edges.add(0);
  if (closesTable) edges.add(lastIndex);
  for (const index of edges) {
    const node = blocks[index];
    if (node === undefined) continue;
    const closed = closedTable(
      node,
      opensTable && index === 0 ? slice.openStart : 0,
      closesTable && index === lastIndex ? slice.openEnd : 0,
    );
    if (closed === undefined) return undefined;
    blocks[index] = closed;
  }
  return new Slice(
    Fragment.from(blocks),
    opensTable ? 1 : slice.openStart,
    closesTable ? 1 : slice.openEnd,
  );
};

/**
 * The slice with each container it is open through at an edge given up for
 * the content copied from it: a copy that starts or ends inside a block
 * content control, custom XML element, text box or inline content control
 * holds part of it, which pastes as that content, as a copy from inside a
 * paragraph pastes as its text. A block wrapper open at an edge is given up
 * however deep the slice opens through it, its first or last paragraph then
 * pasting whole; a paragraph open only at depth one keeps its inline content.
 */
const withOpenContainersUnwrapped = (slice: Slice): Slice => {
  /**
   * One edge unwrapped. A node the edge gives up that is also the other
   * edge's, the only child where both edges run through it, takes a depth
   * from the other edge too.
   */
  const unwrapEdge = (
    fragment: Fragment,
    open: number,
    other: number,
    edge: "start" | "end",
  ): { fragment: Fragment; open: number; other: number } => {
    let current = fragment;
    let depth = open;
    let otherDepth = other;
    let shared = true;
    while (depth > 0) {
      const node = edge === "start" ? current.firstChild : current.lastChild;
      if (node === null || node.type.name === "table") break;
      const children: ProseNode[] = [];
      current.forEach((child) => children.push(child));
      if (node.type.name !== "paragraph") {
        if (node.isTextblock || node.isLeaf || node.childCount === 0) break;
        // An inline node at the top of a slice open one level holds its
        // content whole; a block wrapper open at all holds part of its own.
        if (depth === 1 && node.isInline) break;
        const inner: ProseNode[] = [];
        node.forEach((child) => inner.push(child));
        current = Fragment.from(
          edge === "start"
            ? [...inner, ...children.slice(1)]
            : [...children.slice(0, -1), ...inner],
        );
        shared &&= children.length === 1 && otherDepth > 0;
        if (shared) otherDepth -= 1;
        depth -= 1;
        continue;
      }
      // Open at depth one, a paragraph holds its inline content whole.
      if (depth === 1) break;
      const inline =
        edge === "start" ? node.content.firstChild : node.content.lastChild;
      if (inline === null || inline.isText || inline.isLeaf) break;
      const siblings: ProseNode[] = [];
      node.content.forEach((child) => siblings.push(child));
      const lifted: ProseNode[] = [];
      inline.forEach((child) => lifted.push(child));
      const paragraph = node.copy(
        Fragment.from(
          edge === "start"
            ? [...lifted, ...siblings.slice(1)]
            : [...siblings.slice(0, -1), ...lifted],
        ),
      );
      current = Fragment.from(
        edge === "start"
          ? [paragraph, ...children.slice(1)]
          : [...children.slice(0, -1), paragraph],
      );
      shared &&=
        children.length === 1 && siblings.length === 1 && otherDepth > 1;
      if (shared) otherDepth -= 1;
      depth -= 1;
    }
    return { fragment: current, open: depth, other: otherDepth };
  };
  const start = unwrapEdge(
    slice.content,
    slice.openStart,
    slice.openEnd,
    "start",
  );
  const end = unwrapEdge(start.fragment, start.other, start.open, "end");
  return new Slice(end.fragment, end.other, end.open);
};

/** Range markers between blocks hold no text; a paste leaves them behind. */
const BETWEEN_BLOCK_MARKERS: ReadonlySet<BlockContent["type"]> = new Set([
  "bookmarkStart",
  "bookmarkEnd",
  "commentRangeStart",
  "commentRangeEnd",
  "moveFromRangeStart",
  "moveFromRangeEnd",
  "moveToRangeStart",
  "moveToRangeEnd",
]);

/** A converted clipboard block the paste places, or `undefined` for a marker it leaves. */
const pastedBlockOf = (block: BlockContent): PastedBlock | undefined => {
  switch (block.type) {
    case "paragraph":
    case "table":
    case "blockSdt":
    case "blockCustomXml":
    case "preservedBlock":
      return block;
    default:
      return BETWEEN_BLOCK_MARKERS.has(block.type)
        ? undefined
        : panic(`A clipboard block of kind ${block.type} has no placement.`);
  }
};

/** `blocks` without the range markers between them, at any depth. */
const withoutBlockMarkers = (blocks: BlockContent[]): void => {
  for (let index = blocks.length - 1; index >= 0; index -= 1) {
    const block = blocks[index];
    if (block === undefined) continue;
    if (BETWEEN_BLOCK_MARKERS.has(block.type)) {
      blocks.splice(index, 1);
      continue;
    }
    if (block.type === "blockSdt" || block.type === "blockCustomXml")
      withoutBlockMarkers(block.content);
  }
};
type ClipboardParagraphPlace = {
  paragraph: Paragraph;
  replace: (next: Paragraph) => void;
};

/**
 * A pasted block's paragraphs in document order, a table's cell paragraphs
 * and a container's among them, each with the place it fills.
 */
const clipboardParagraphs = (
  block: PastedBlock,
  replace: (next: Paragraph) => void = () => {},
): ClipboardParagraphPlace[] | undefined => {
  if (block.type === "paragraph") return [{ paragraph: block, replace }];
  if (block.type === "preservedBlock") return [];
  if (block.type === "blockSdt" || block.type === "blockCustomXml") {
    const places: ClipboardParagraphPlace[] = [];
    for (const [index, inner] of block.content.entries()) {
      const pasted = pastedBlockOf(inner);
      if (pasted === undefined) continue;
      const found = clipboardParagraphs(pasted, (next) => {
        block.content[index] = next;
      });
      if (found === undefined) return undefined;
      places.push(...found);
    }
    return places;
  }
  const places: ClipboardParagraphPlace[] = [];
  for (const row of block.rows)
    for (const cell of row.cells)
      for (const [index, inner] of cell.content.entries()) {
        // A range marker in a cell holds no paragraph.
        const pasted = pastedBlockOf(inner);
        if (pasted === undefined) continue;
        const found = clipboardParagraphs(pasted, (next) => {
          cell.content[index] = next;
        });
        if (found === undefined) return undefined;
        places.push(...found);
      }
  return places;
};

/**
 * A table's verbatim captures describe the markup it was read from; a pasted
 * copy is new content the serializer writes from its model.
 */
/** Every table among the blocks, those nested in cells and containers included. */
const clipboardTables = (blocks: readonly BlockContent[]): Table[] =>
  blocks.flatMap((block) => {
    switch (block.type) {
      case "table":
        return [
          block,
          ...clipboardTables(
            block.rows.flatMap((row) =>
              row.cells.flatMap((cell) => cell.content),
            ),
          ),
        ];
      case "blockSdt":
      case "blockCustomXml":
        return clipboardTables(block.content);
      default:
        return [];
    }
  });

/** Every paragraph among the blocks, at any depth. */
const nestedParagraphs = (blocks: readonly BlockContent[]): Paragraph[] =>
  blocks.flatMap((block) => {
    switch (block.type) {
      case "paragraph":
        return [block];
      case "table":
        return nestedParagraphs(
          block.rows.flatMap((row) =>
            row.cells.flatMap((cell) => cell.content),
          ),
        );
      case "blockSdt":
      case "blockCustomXml":
        return nestedParagraphs(block.content);
      default:
        return [];
    }
  });

/** A table's own captures; the tables in its cells are their own. */
const withoutTableCaptures = (table: Table): void => {
  if (table.formatting !== undefined) {
    delete table.formatting.sourceXml;
    delete table.formatting.gridSourceXml;
  }
  for (const row of table.rows) {
    if (row.formatting !== undefined) delete row.formatting.sourceXml;
    for (const cell of row.cells)
      if (cell.formatting !== undefined) delete cell.formatting.sourceXml;
  }
};

/**
 * Every table among the blocks freed of its verbatim captures, those nested
 * in cells and containers, a container in a cell among them, included.
 */
const withoutContainedTableCaptures = (blocks: readonly BlockContent[]) => {
  for (const table of clipboardTables(blocks)) withoutTableCaptures(table);
};

/** A pasted block's own copy, each paragraph keeping its indentation provenance. */
const cloneClipboardBlock = (block: PastedBlock): PastedBlock => {
  const clone = structuredClone(block);
  withoutContainedTableCaptures([clone]);
  if (clone.type === "blockSdt" || clone.type === "blockCustomXml")
    withoutBlockMarkers(clone.content);
  const sources = clipboardParagraphs(block) ?? [];
  const targets = clipboardParagraphs(clone) ?? [];
  for (const [index, { paragraph: target }] of targets.entries()) {
    const source = sources[index]?.paragraph;
    if (source !== undefined)
      copyParagraphIndentationProjection({ target, source });
  }
  return clone;
};

const IMAGE_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const HYPERLINK_RELATIONSHIP =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const MAX_CLIPBOARD_IMAGE_BYTES = 16 * 1024 * 1024;

const isNumberFormat = (format: CounterFormat): format is NumberFormat =>
  NUMBER_FORMATS.some((value) => value === format);

const numberFormatFields = (format: CounterFormat) => {
  if (isNumberFormat(format)) return { numFmt: format };
  switch (format) {
    case "decimalZero3":
      return { numFmt: "custom", numFmtFormat: "000" } as const;
    case "decimalZero4":
      return { numFmt: "custom", numFmtFormat: "0000" } as const;
    case "decimalZero5":
      return { numFmt: "custom", numFmtFormat: "00000" } as const;
    default: {
      const exhaustive: never = format;
      return exhaustive;
    }
  }
};

type ImportNumberingOptions = {
  destination: NumberingDefinitions | undefined;
  source: NumberingDefinitions | undefined;
  paragraphs: Paragraph[];
  mode: "copy" | "move";
  sourceKind: "package" | "normalizedSlice";
  extraIds?: ReadonlySet<number>;
  styleIds?: ReadonlyMap<string, string>;
};

/** Clipboard instance ids belong to their source package, so imported lists get fresh ids. */
const importNumbering = ({
  destination,
  source,
  paragraphs,
  mode,
  sourceKind,
  extraIds,
  styleIds,
}: ImportNumberingOptions) => {
  if (mode === "move")
    return Result.ok({
      numbering: destination,
      numIds: new Map<number, number>(),
    });
  const references = paragraphs.flatMap((paragraph) => {
    const formats = [
      paragraph.formatting,
      ...(paragraph.propertyChanges ?? []).map(
        (change) => change.previousFormatting,
      ),
    ];
    return formats.flatMap((formatting) =>
      formatting?.numPr?.kind === "reference" ? [formatting.numPr.numId] : [],
    );
  });
  references.push(...(extraIds ?? []));
  if (references.length === 0)
    return Result.ok({
      numbering: destination,
      numIds: new Map<number, number>(),
    });
  if (source === undefined)
    return refuse(
      "The clipboard list has no numbering definition.",
      INVALID_CLIPBOARD,
    );
  const sourceDefinitions = source;
  if (
    references.some(
      (id) => !sourceDefinitions.nums.some(({ numId }) => numId === id),
    )
  )
    return refuse(
      "The clipboard list has no numbering definition.",
      INVALID_CLIPBOARD,
    );
  const used = new Set(references);
  const sourceNums = source.nums.filter(({ numId }) => used.has(numId));
  const usedAbstracts = new Set(
    sourceNums.map(({ abstractNumId }) => abstractNumId),
  );
  source = {
    ...source,
    nums: structuredClone(sourceNums),
    abstractNums: structuredClone(
      source.abstractNums.filter(({ abstractNumId }) =>
        usedAbstracts.has(abstractNumId),
      ),
    ),
  };
  if (sourceKind === "normalizedSlice") {
    for (const paragraph of paragraphs) {
      const rendering = paragraph.listRendering;
      if (rendering === undefined) continue;
      const instance = source.nums.find(
        ({ numId }) => numId === rendering.numId,
      );
      const definition = source.abstractNums.find(
        ({ abstractNumId }) => abstractNumId === instance?.abstractNumId,
      );
      if (definition === undefined)
        return refuse(
          "The clipboard list has no abstract definition.",
          INVALID_CLIPBOARD,
        );
      const level = definition.levels.find(
        ({ ilvl }) => ilvl === rendering.level,
      );
      if (level === undefined)
        return refuse(
          "The clipboard list level could not be imported.",
          INVALID_CLIPBOARD,
        );
      Object.assign(level, {
        ...numberFormatFields(
          rendering.numFmt ?? (rendering.isBullet ? "bullet" : "decimal"),
        ),
        lvlText: rendering.markerTemplate ?? rendering.marker,
        start: rendering.levelStarts?.at(rendering.level) ?? level.start ?? 1,
        ...(rendering.isLegal === undefined
          ? {}
          : { isLgl: rendering.isLegal }),
        ...(rendering.markerAlignment === undefined
          ? {}
          : { lvlJc: rendering.markerAlignment }),
        ...(rendering.markerSuffix === undefined
          ? {}
          : { suffix: rendering.markerSuffix }),
        ...(rendering.levelTabs === undefined
          ? {}
          : { pPr: { ...level.pPr, tabs: rendering.levelTabs } }),
        ...(rendering.markerFormatting === undefined
          ? {}
          : { rPr: rendering.markerFormatting }),
      });
      if (
        rendering.markerAllCaps !== undefined ||
        rendering.markerHidden !== undefined
      ) {
        level.rPr = {
          ...level.rPr,
          ...(rendering.markerAllCaps === undefined
            ? {}
            : { allCaps: rendering.markerAllCaps }),
          ...(rendering.markerHidden === undefined
            ? {}
            : { hidden: rendering.markerHidden }),
        };
      }
    }
  }
  const numAllocator = createNumberingIdAllocator(
    "num",
    numberingNumIdsTaken(destination),
  );
  const abstractAllocator = createNumberingIdAllocator(
    "abstract",
    (destination?.abstractNums ?? []).map(({ abstractNumId }) => abstractNumId),
  );
  const nums = new Map(
    source.nums.map(({ numId }) => [numId, numAllocator.next()]),
  );
  const abstracts = new Map(
    source.abstractNums.map(({ abstractNumId }) => [
      abstractNumId,
      abstractAllocator.next(),
    ]),
  );
  for (const paragraph of paragraphs) {
    const formats = [
      paragraph.formatting,
      ...(paragraph.propertyChanges ?? []).map(
        (change) => change.previousFormatting,
      ),
    ];
    for (const formatting of formats) {
      if (formatting?.numPr?.kind !== "reference") continue;
      const numId = nums.get(formatting.numPr.numId);
      if (numId === undefined)
        return refuse(
          "The clipboard list instance could not be imported.",
          INVALID_CLIPBOARD,
        );
      formatting.numPr = paragraphNumberingReference({
        numId,
        ...(formatting.numPr.ilvl === undefined
          ? {}
          : { ilvl: formatting.numPr.ilvl }),
      });
    }
    if (paragraph.listRendering !== undefined) {
      const numId = nums.get(paragraph.listRendering.numId);
      if (numId !== undefined) paragraph.listRendering.numId = numId;
      const abstractId = paragraph.listRendering.abstractNumId;
      if (abstractId !== undefined) {
        const abstractNumId = abstracts.get(abstractId);
        if (abstractNumId !== undefined)
          paragraph.listRendering.abstractNumId = abstractNumId;
      }
    }
  }
  const importedNums = [];
  for (const instance of source.nums) {
    const numId = nums.get(instance.numId);
    const abstractNumId = abstracts.get(instance.abstractNumId);
    if (numId === undefined || abstractNumId === undefined)
      return refuse(
        "The clipboard numbering definition could not be imported.",
        INVALID_CLIPBOARD,
      );
    const imported = { ...instance, numId, abstractNumId };
    for (const override of imported.levelOverrides ?? []) {
      if (override.lvl?.pStyle === undefined || styleIds === undefined)
        continue;
      const mapped = styleIds.get(override.lvl.pStyle);
      if (mapped === undefined)
        return refuse(
          "A clipboard numbering style dependency could not be imported.",
          INVALID_CLIPBOARD,
        );
      override.lvl.pStyle = mapped;
    }
    importedNums.push(imported);
  }
  const importedAbstracts = [];
  for (const definition of source.abstractNums) {
    const abstractNumId = abstracts.get(definition.abstractNumId);
    if (abstractNumId === undefined)
      return refuse(
        "The clipboard numbering definition could not be imported.",
        INVALID_CLIPBOARD,
      );
    const imported = { ...definition, abstractNumId };
    if (styleIds !== undefined) {
      if (imported.numStyleLink !== undefined) {
        const mapped = styleIds.get(imported.numStyleLink);
        if (mapped === undefined)
          return refuse(
            "A clipboard numbering style dependency could not be imported.",
            INVALID_CLIPBOARD,
          );
        imported.numStyleLink = mapped;
      }
      if (imported.styleLink !== undefined) {
        const mapped = styleIds.get(imported.styleLink);
        if (mapped === undefined)
          return refuse(
            "A clipboard numbering style dependency could not be imported.",
            INVALID_CLIPBOARD,
          );
        imported.styleLink = mapped;
      }
      for (const level of imported.levels) {
        if (level.pStyle === undefined) continue;
        const mapped = styleIds.get(level.pStyle);
        if (mapped === undefined)
          return refuse(
            "A clipboard numbering style dependency could not be imported.",
            INVALID_CLIPBOARD,
          );
        level.pStyle = mapped;
      }
    }
    importedAbstracts.push(imported);
  }
  return Result.ok({
    numbering: {
      ...destination,
      nums: [...(destination?.nums ?? []), ...importedNums],
      abstractNums: [
        ...(destination?.abstractNums ?? []),
        ...importedAbstracts,
      ],
    },
    numIds: nums,
  });
};

type ClipboardAnnotationsOptions = {
  paragraphs: readonly Paragraph[];
  source: Document | undefined;
};
/**
 * The paragraphs without their comment markers, comment references and note
 * references: their text stays, and nothing names a comment or note.
 */
const withoutAnnotationMarkers = (paragraphs: readonly Paragraph[]): void => {
  for (const paragraph of paragraphs) {
    const marked: { content: unknown[]; index: number }[] = [];
    visitInlineContentSlots(paragraph, ({ content, index, item }) => {
      if (
        item.type === "commentRangeStart" ||
        item.type === "commentRangeEnd" ||
        item.type === "commentReference"
      )
        marked.push({ content, index });
    });
    // Each array from its last marker back, so no removal moves another.
    for (const { content, index } of marked.toReversed())
      content.splice(index, 1);
    const emptied: { content: unknown[]; index: number }[] = [];
    visitInlineContentSlots(paragraph, ({ content, index, item }) => {
      if (item.type !== "run") return;
      const kept = item.content.filter(
        (child) => child.type !== "footnoteRef" && child.type !== "endnoteRef",
      );
      if (kept.length === item.content.length) return;
      if (kept.length === 0) emptied.push({ content, index });
      else item.content = kept;
    });
    for (const { content, index } of emptied.toReversed())
      content.splice(index, 1);
  }
};

/**
 * The notes and comments the clipboard's references and markers name, read
 * from the document they were copied from; the paste copies each under a
 * fresh id. Without that document the ids could name the destination's own
 * notes and comments: the paste keeps the text and leaves the markers and
 * references out, as a copy made outside the editor holds no comment or note.
 */
export const clipboardAnnotations = ({
  paragraphs,
  source,
}: ClipboardAnnotationsOptions): Result<
  PasteAnnotations | undefined,
  CanonicalSessionError
> => {
  const comments = new Set<number>();
  const notes = new Set<string>();
  for (const paragraph of paragraphs) {
    visitInlineContentSlots(paragraph, ({ item }) => {
      if (
        item.type === "commentRangeStart" ||
        item.type === "commentRangeEnd" ||
        item.type === "commentReference"
      )
        comments.add(item.id);
    });
    visitParagraphRuns(paragraph, (run) => {
      for (const item of run.content)
        if (item.type === "footnoteRef" || item.type === "endnoteRef")
          notes.add(
            `${item.type === "footnoteRef" ? "footnote" : "endnote"}:${String(item.id)}`,
          );
    });
  }
  if (comments.size === 0 && notes.size === 0) return Result.ok(undefined);
  if (source === undefined) {
    withoutAnnotationMarkers(paragraphs);
    return Result.ok(undefined);
  }
  const annotations = {
    notes: [
      ...(source.package.footnotes ?? []),
      ...(source.package.endnotes ?? []),
    ].filter(({ type, id }) => notes.has(`${type}:${String(id)}`)),
    comments: (source.package.document.comments ?? []).filter(({ id }) =>
      comments.has(id),
    ),
  };
  if (
    annotations.notes.length !== notes.size ||
    annotations.comments.length !== comments.size
  )
    return refuse(
      "A clipboard comment or note reference names none in the document it was copied from.",
      INVALID_CLIPBOARD,
    );
  return Result.ok(annotations);
};

/** The relationship ids markup names, in any office relationship attribute. */
const relationshipIdsIn = (
  xml: string,
): Result<string[], CanonicalSessionError> => {
  const parsed = Result.try({
    try: () => parseXml(`<root>${xml}</root>`, OOXML_NAMESPACE_SCOPE),
    catch: () =>
      new CanonicalSessionError({
        ground: INVALID_CLIPBOARD,
        reason: "refused",
        message: "The clipboard block markup is malformed.",
      }),
  });
  if (parsed.isErr()) return parsed;
  const ids: string[] = [];
  const visit = (element: XmlElement): void => {
    for (const name of Object.keys(element.attributes ?? {})) {
      const attribute = findAttributeByNamespaceUri(
        element,
        OFFICE_RELATIONSHIP_NAMESPACE_URIS,
        getLocalName(name),
      );
      if (attribute !== null && attribute.name === name)
        ids.push(attribute.value);
    }
    for (const child of element.elements ?? []) visit(child);
  };
  visit(parsed.value);
  return Result.ok(ids);
};

/** Whether two relationships name the same target in the same way. */
const sameRelationship = (
  left: { type: string; target: string; targetMode?: string } | undefined,
  right: { type: string; target: string; targetMode?: string } | undefined,
): boolean =>
  left !== undefined &&
  right !== undefined &&
  left.type === right.type &&
  left.target === right.target &&
  left.targetMode === right.targetMode;

/**
 * The image relationship ids and the markup a range of the editor's own
 * document holds. A move carries what it drags within the part whose
 * relationships it names, so what this range holds stays that part's own.
 */
const rangeHoldings = (
  doc: ProseNode,
  { from, to }: { from: number; to: number },
): { relationshipIds: ReadonlySet<string>; markup: ReadonlySet<string> } => {
  const relationshipIds = new Set<string>();
  const markup = new Set<string>();
  doc.nodesBetween(from, to, (node) => {
    const rId: unknown = node.attrs["rId"];
    if (node.type.name === "image" && typeof rId === "string")
      relationshipIds.add(rId);
    for (const value of Object.values(node.attrs))
      if (typeof value === "string" && value.startsWith("<")) markup.add(value);
    return true;
  });
  return { relationshipIds, markup };
};

/** A picture source a paste imports: image bytes, or an HTTP image URL. */
const IMPORTABLE_PICTURE_SOURCE =
  /^(?:data:image\/(?:png|jpeg|gif|webp);base64,|https?:\/\/)/u;

const WORDPROCESSING_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Control characters a paragraph's text cannot hold, a tab aside. */
// oxlint-disable-next-line no-control-regex -- matches the characters XML text cannot hold
const UNWRITABLE_TEXT = /[\u0000-\u0008\u000B-\u001F\uFFFE\uFFFF]/gu;

/** A paragraph holding `text`, a tab written as a tab. */
const textParagraph = (text: string): Paragraph => {
  const content: Run["content"] = [];
  for (const [index, part] of text
    .replaceAll(UNWRITABLE_TEXT, "")
    .split("\t")
    .entries()) {
    if (index > 0) content.push({ type: "tab" });
    if (part.length > 0) content.push({ type: "text", text: part });
  }
  return {
    type: "paragraph",
    content: content.length === 0 ? [] : [{ type: "run", content }],
  };
};

/**
 * The text of each paragraph kept block markup holds, in order, or of each
 * line of the text read from the part it names: none when it holds none.
 */
const keptBlockTexts = (
  xml: string,
  readerText: string | undefined,
): string[] => {
  if (readerText !== undefined)
    return readerText
      .split(/\r\n|\r|\n/u)
      .filter((line) => line.trim().length > 0);
  const parsed = Result.try(() =>
    parseXml(`<root>${xml}</root>`, OOXML_NAMESPACE_SCOPE),
  );
  if (parsed.isErr()) return [];
  const texts: string[] = [];
  const isWord = (element: XmlElement, name: string) =>
    getLocalName(element.name ?? "") === name &&
    getNamespaceUri(element) === WORDPROCESSING_NAMESPACE;
  const textOf = (element: XmlElement): string =>
    (element.elements ?? [])
      .map((child) => {
        if (isWord(child, "t")) return getTextContent(child);
        if (isWord(child, "tab")) return "\t";
        // Deleted text is not the paragraph's text.
        if (isWord(child, "del") || isWord(child, "moveFrom")) return "";
        return textOf(child);
      })
      .join("");
  const visit = (element: XmlElement): void => {
    for (const child of element.elements ?? []) {
      if (isWord(child, "p")) {
        texts.push(textOf(child));
        continue;
      }
      visit(child);
    }
  };
  visit(parsed.value);
  return texts.some((text) => text.trim().length > 0) ? texts : [];
};

/**
 * `paragraph` without the runs of `emptied` left empty, and without each
 * wrapper their removal leaves empty: markup a paste left out leaves no empty
 * run or wrapper behind.
 */
const withoutEmptiedRuns = (
  paragraph: Paragraph,
  emptied: ReadonlySet<Run>,
): void => {
  const shrunk = new Set<ParagraphContent[]>();
  const leftEmpty = (item: ParagraphContent): boolean => {
    if (item.type === "run")
      return emptied.has(item) && item.content.length === 0;
    let empty = false;
    withInlineContentRewritten(item, (children) => {
      empty = children.length === 0 && shrunk.has(children);
      return children;
    });
    return empty;
  };
  for (;;) {
    const removed: InlineContentSlot[] = [];
    visitInlineContentSlots(paragraph, (slot) => {
      if (leftEmpty(slot.item)) removed.push(slot);
    });
    if (removed.length === 0) return;
    for (const { content, index } of removed.toReversed()) {
      content.splice(index, 1);
      shrunk.add(content);
    }
  }
};

/** Clipboard slices are input payloads; their paragraph identities and private captures are not owners. */
type PrepareCanonicalPasteOptions = {
  session: CanonicalSession;
  state: EditorState;
  slice: Slice;
  plain?: boolean;
  moveTarget?: number;
  pasteTarget?: number;
  sourceDocument?: Document;
  /**
   * The parts `sourceDocument`'s package holds that the copied markup names,
   * read ahead of the paste: with them, markup naming parts pastes into
   * another package.
   */
  sourceParts?: ClipboardPackageParts;
  /** The story of `sourceDocument` the slice was copied from; the body when absent. */
  sourceStory?: OpStory;
  /** The story `state` shows, which the paste lands in; the body when absent. */
  story?: OpStory;
  moveSource?: { from: number; to: number };
};

/**
 * What a paste did in place of keeping markup whose package parts cannot
 * come with it: a clipboard from outside the editor, or a copy whose source
 * package could not be read.
 */
export const CANONICAL_PASTE_WARNING_CODES = {
  /** A drawing or embedded object pasted as the picture that previews it. */
  keptMarkupAsPicture: "kept-markup-as-picture",
  /** Kept block markup pasted as the paragraphs of text it holds. */
  keptMarkupAsText: "kept-markup-as-text",
  /** Kept markup holding neither a picture nor text, left out. */
  keptMarkupDropped: "kept-markup-dropped",
} as const;

/** One substitution a paste made, as data. */
export type CanonicalPasteWarning = {
  code: (typeof CANONICAL_PASTE_WARNING_CODES)[keyof typeof CANONICAL_PASTE_WARNING_CODES];
  /**
   * The part the markup was pasted into: `document`, `footnotes`,
   * `endnotes`, `comments`, or a header's or footer's `header:<rId>` or
   * `footer:<rId>`.
   */
  part: string;
  /** The kept element's name as the markup spells it, such as `w:altChunk`. */
  element: string;
};

/** A prepared paste, with what it pasted in place of markup it could not keep. */
export type CanonicalPasteCommit = CanonicalCommit & {
  readonly warnings: readonly CanonicalPasteWarning[];
};

export const prepareCanonicalPaste = ({
  session,
  state,
  slice: copied,
  plain = false,
  moveTarget,
  pasteTarget,
  sourceDocument,
  sourceParts,
  sourceStory = OP_STORIES.MAIN,
  story = OP_STORIES.MAIN,
  moveSource,
}: PrepareCanonicalPasteOptions): Result<
  CanonicalPasteCommit,
  CanonicalSessionError
> => {
  // Declared: the inferred union of Ok/Err branches is emitted in a nondeterministic order.
  if (session.isComposing)
    return refuse("Composition must finish before using the clipboard.", {
      condition: CANONICAL_CONDITION.composing,
    });
  if (copied.content.size === 0)
    return Result.err(
      new CanonicalSessionError({
        ground: INVALID_CLIPBOARD,
        message: "The clipboard contains no content.",
        reason: "noChange",
      }),
    );
  // A copy that opens or closes inside a table pastes the table of the
  // cells it holds; one that opens through a container pastes the content
  // it copied from it.
  const closed = withClosedTableEdges(copied);
  if (closed === undefined)
    return refuse("The clipboard table holds no cells.", INVALID_CLIPBOARD);
  // A table the copy opened through a container is reached only once the
  // container is given up; it pastes the cells it holds as a top-level one does.
  const slice = withClosedTableEdges(withOpenContainersUnwrapped(closed));
  if (slice === undefined)
    return refuse("The clipboard table holds no cells.", INVALID_CLIPBOARD);
  if (slice.openStart > 1 || slice.openEnd > 1)
    return refuse(
      "Clipboard tables and nested block containers require canonical table editing.",
      UNPLACED_BLOCK,
    );
  const paragraphType = state.schema.nodes["paragraph"];
  if (paragraphType === undefined)
    return refuse("Clipboard input requires a paragraph schema.", {
      condition: CANONICAL_CONDITION.internal,
    });
  let inline = true;
  let blocksOnly = true;
  slice.content.forEach((node) => {
    inline &&= node.isInline;
    blocksOnly &&= node.isBlock;
  });
  if (!inline && !blocksOnly)
    return refuse(
      "Clipboard tables and embedded blocks require canonical table editing.",
      UNPLACED_BLOCK,
    );
  // A same-parent document slice omits its paragraph; it still represents open inline edges.
  const importedSlice = inline
    ? new Slice(Fragment.from(paragraphType.create(null, slice.content)), 1, 1)
    : slice;
  const convertedBlocks = Result.try({
    // Imported fragments mint new paragraph identities; matching source ids
    // must not attach a durable paragraph owner from either document.
    try: () =>
      proseDocToBlocks(
        state.schema.topNodeType.create(null, importedSlice.content),
        [],
      ),
    catch: (cause) =>
      new CanonicalSessionError({
        ground: INVALID_CLIPBOARD,
        reason: "refused",
        message: `Clipboard normalization failed: ${cause instanceof Error ? cause.message : String(cause)}`,
      }),
  });
  if (convertedBlocks.isErr()) return convertedBlocks;
  const blocks: PastedBlock[] = [];
  for (const converted of convertedBlocks.value) {
    // A range marker between blocks names a range the paste does not hold
    // whole; the paste leaves it behind, as it does a bookmark it splits.
    const block = pastedBlockOf(converted);
    if (block === undefined) continue;
    const clone = cloneClipboardBlock(block);
    if (clipboardParagraphs(clone) === undefined)
      return refuse(
        "The clipboard contains an unsupported block.",
        UNPLACED_BLOCK,
      );
    blocks.push(clone);
  }
  if (blocks.length === 0)
    return Result.err(
      new CanonicalSessionError({
        ground: INVALID_CLIPBOARD,
        message: "The clipboard contains no content.",
        reason: "noChange",
      }),
    );
  // Every pasted paragraph, a table's cell paragraphs among them, with the
  // place it fills: an import that returns new paragraphs puts them there.
  const places = blocks.flatMap(
    (block, index) =>
      clipboardParagraphs(block, (next) => {
        blocks[index] = next;
      }) ?? [],
  );
  const paragraphs = places.map(({ paragraph }) => paragraph);
  const tables = clipboardTables(blocks);
  if (moveTarget === undefined && (plain || sourceDocument === undefined)) {
    flattenClipboardStyleReferences(paragraphs);
    // A table style named outside a package this paste can read is not ours.
    for (const table of tables) delete table.formatting?.styleId;
  }
  // Over selected cells the paste clears every cell and lands in the first.
  const overCells =
    state.selection instanceof CellSelection &&
    pasteTarget === undefined &&
    moveSource === undefined &&
    moveTarget === undefined;
  if (state.selection instanceof CellSelection && !overCells)
    return refuse(
      "Pasting over selected cells requires a per-cell paste operation.",
      INVALID_CLIPBOARD,
    );
  const selectionRange = overCells
    ? { from: state.selection.from, to: state.selection.to }
    : canonicalSelectionRange(state);
  // The projection of the story the paste lands in, which `state` shows.
  const projected = session.projectStory(story);
  if (projected.isErr()) return projected;
  const projection = projected.value;
  const from = overCells
    ? Result.ok(undefined)
    : projection.addressAt(
        pasteTarget ?? moveSource?.from ?? selectionRange.from,
      );
  if (from.isErr()) return from;
  const to = overCells
    ? Result.ok(undefined)
    : projection.addressAt(pasteTarget ?? moveSource?.to ?? selectionRange.to);
  if (to.isErr()) return to;
  // The notes and comments the blocks name, copied from their source: their
  // bodies import styles and numbering with the blocks' own paragraphs.
  const named =
    moveTarget === undefined
      ? clipboardAnnotations({ paragraphs, source: sourceDocument })
      : Result.ok(undefined);
  if (named.isErr()) return named;
  const annotations =
    named.value === undefined
      ? undefined
      : {
          notes: structuredClone([...named.value.notes]),
          comments: structuredClone([...named.value.comments]),
        };
  const annotationBodies: ClipboardParagraphPlace[] = [];
  for (const note of annotations?.notes ?? [])
    for (const [index, block] of note.content.entries())
      if (block.type === "paragraph")
        annotationBodies.push({
          paragraph: block,
          replace: (next) => {
            note.content[index] = next;
          },
        });
  for (const comment of annotations?.comments ?? [])
    for (const [index, paragraph] of comment.content.entries())
      annotationBodies.push({
        paragraph,
        replace: (next) => {
          comment.content[index] = next;
        },
      });
  const annotationParagraphs = annotationBodies.map(
    ({ paragraph }) => paragraph,
  );
  // The session's own document, read to build the paste's operations.
  const destination = readModel(session.document);
  let styles = destination.package.styles;
  let styleIds: ReadonlyMap<string, string> | undefined;
  let extraIds: ReadonlySet<number> | undefined;
  if (sourceDocument !== undefined && moveTarget === undefined && !plain) {
    const importedStyles = importClipboardStyles({
      destination,
      source: sourceDocument,
      paragraphs,
      annotationParagraphs,
      tableStyleIds: tables.flatMap(({ formatting }) =>
        formatting?.styleId === undefined ? [] : [formatting.styleId],
      ),
    });
    if (importedStyles.isErr()) {
      const { reason, message } = importedStyles.error;
      return refuse(message, groundForDocumentOpRefusal(reason));
    }
    styles = importedStyles.value.styles;
    for (const [index, place] of places.entries()) {
      const imported =
        importedStyles.value.paragraphs[index] ??
        panic("An imported clipboard paragraph is missing.");
      place.replace(imported);
      paragraphs[index] = imported;
    }
    for (const [offset, body] of annotationBodies.entries()) {
      const imported =
        importedStyles.value.annotationParagraphs[offset] ??
        panic("An imported copied note or comment paragraph is missing.");
      body.replace(imported);
      annotationParagraphs[offset] = imported;
    }
    styleIds = importedStyles.value.styleIds;
    extraIds = importedStyles.value.numberingIds;
    for (const { formatting } of tables)
      if (formatting?.styleId !== undefined)
        formatting.styleId =
          importedStyles.value.styleIds.get(formatting.styleId) ??
          panic("An imported table style has no destination identity.");
  }
  const sourceNumbering =
    sourceDocument?.package.numbering ??
    completeNumberingForDoc(
      undefined,
      state.schema.topNodeType.create(null, importedSlice.content),
    );
  const importedNumbering = importNumbering({
    destination: destination.package.numbering,
    source: sourceNumbering,
    paragraphs: [...paragraphs, ...annotationParagraphs],
    mode: moveTarget === undefined ? "copy" : "move",
    sourceKind: sourceDocument === undefined ? "normalizedSlice" : "package",
    ...(extraIds === undefined ? {} : { extraIds }),
    ...(styleIds === undefined ? {} : { styleIds }),
  });
  if (importedNumbering.isErr()) return importedNumbering;
  const numbering = importedNumbering.value.numbering;
  if (styles !== undefined && styleIds !== undefined) {
    const importedStyleIds = new Set(styleIds.values());
    for (const style of styles.styles) {
      if (
        !importedStyleIds.has(style.styleId) ||
        style.pPr?.numPr?.kind !== "reference"
      )
        continue;
      const numId = importedNumbering.value.numIds.get(style.pPr.numPr.numId);
      if (numId === undefined)
        return refuse(
          "A clipboard style numbering dependency could not be imported.",
          INVALID_CLIPBOARD,
        );
      style.pPr.numPr = paragraphNumberingReference({
        numId,
        ...(style.pPr.numPr.ilvl === undefined
          ? {}
          : { ilvl: style.pPr.numPr.ilvl }),
      });
    }
  }
  const relationships = new Map(destination.package.relationships);
  const media = new Map(destination.package.media);
  const copiedParts = new Map(destination.package.copiedParts);
  // The note, comment, header and footer parts' relationships, each copied
  // once kept markup pasted into the part names a copy.
  const annotationRelationships = new Map(
    destination.package.annotationRelationships,
  );
  const changedAnnotationParts = new Set<string>();
  let importedResource = false;
  let copiedPart = false;
  let imageIndex = 1;
  let relationshipIndex = 1;
  // A picture pasted into a note, a comment, a header or a footer takes a
  // main document id that the save writes into that part's relationships, so
  // an id is fresh only where no part holds it.
  const heldRelationshipId = (id: string) =>
    relationships.has(id) ||
    [...annotationRelationships.values()].some((held) => held.has(id));
  const freshRelationshipId = () => {
    while (heldRelationshipId(`rId${relationshipIndex}`))
      relationshipIndex += 1;
    return `rId${relationshipIndex++}`;
  };
  // The relationships the copied pictures' and links' ids name: the source
  // body's, or none for a note's, a comment's, a header's or a footer's,
  // which import by their bytes and targets.
  let sourceRelationships =
    sourceStory === OP_STORIES.MAIN
      ? sourceDocument?.package.relationships
      : undefined;
  // What the paste put in place of kept markup it could not keep.
  const warnings: CanonicalPasteWarning[] = [];
  const importExternalLink = (
    href: string | undefined,
    previousId: string | undefined,
  ) => {
    const source =
      previousId === undefined
        ? undefined
        : sourceRelationships?.get(previousId);
    let target = href;
    if (target === undefined || target.length === 0)
      target = source?.targetMode === "External" ? source.target : undefined;
    if (target === undefined || target.length === 0)
      return refuse(
        "The clipboard hyperlink has no resolved target.",
        INVALID_CLIPBOARD,
      );
    const id = freshRelationshipId();
    relationships.set(id, {
      id,
      type: HYPERLINK_RELATIONSHIP,
      target,
      targetMode: "External",
    });
    importedResource = true;
    return Result.ok({ id, href: target });
  };
  const importCapturedLinks = (xml: string, remapped: Map<string, string>) => {
    const parsed = Result.try({
      try: () => parseXml(xml, OOXML_NAMESPACE_SCOPE),
      catch: () =>
        new CanonicalSessionError({
          ground: INVALID_CLIPBOARD,
          reason: "refused",
          message: "The clipboard hyperlink metadata is malformed.",
        }),
    });
    if (parsed.isErr()) return parsed;
    const visit = (
      element: XmlElement,
    ): Result<void, CanonicalSessionError> => {
      for (const name of Object.keys(element.attributes ?? {})) {
        const attribute = findAttributeByNamespaceUri(
          element,
          OFFICE_RELATIONSHIP_NAMESPACE_URIS,
          getLocalName(name),
        );
        if (attribute === null || attribute.name !== name) continue;
        if (getLocalName(name) !== "id")
          return refuse(
            "Clipboard hyperlink metadata references an unsupported package part.",
            INVALID_CLIPBOARD,
          );
        let id = remapped.get(attribute.value);
        if (id === undefined) {
          const imported = importExternalLink(undefined, attribute.value);
          if (imported.isErr()) return imported;
          id = imported.value.id;
          remapped.set(attribute.value, id);
        }
        if (element.attributes !== undefined) element.attributes[name] = id;
      }
      for (const child of element.elements ?? []) {
        const result = visit(child);
        if (result.isErr()) return result;
      }
      return Result.ok(undefined);
    };
    const imported = visit(parsed.value);
    if (imported.isErr()) return imported;
    return Result.ok(
      (parsed.value.elements ?? []).map(captureVerbatimXml).join(""),
    );
  };
  // Drawing ids the destination already uses, read once an opaque drawing
  // needs a fresh one.
  let drawingIds: Set<number> | undefined;
  // A move's blocks are the dragged range of this session's own document,
  // which no source package describes: the images and kept markup that range
  // holds keep the relationships they name.
  const moved =
    moveTarget === undefined
      ? undefined
      : rangeHoldings(state.doc, moveSource ?? selectionRange);
  // New inline images have no opaque XML dependencies. Existing package images
  // retain their relationship only when its resolved source belongs to this package.
  const importResources = (
    paragraph: Paragraph,
  ): Result<void, CanonicalSessionError> => {
    if (moveTarget === undefined) {
      const links: Extract<
        Paragraph["content"][number],
        { type: "hyperlink" }
      >[] = [];
      visitInlineContentSlots(paragraph, ({ item }) => {
        if (item.type === "hyperlink") links.push(item);
      });
      for (const link of links) {
        if (link.anchor !== undefined) {
          delete link.rId;
          continue;
        }
        const imported = importExternalLink(link.href, link.rId);
        if (imported.isErr()) return imported;
        link.href = imported.value.href;
        link.rId = imported.value.id;
      }
    }
    const runs: Run[] = [];
    visitParagraphRuns(paragraph, (run) => runs.push(run));
    // The runs a drawing was left out of, removed once they hold nothing.
    const emptied = new Set<Run>();
    for (const item of runs) {
      if (plain)
        item.formatting = marksToTextFormatting(
          state.storedMarks ??
            state.doc
              .resolve(pasteTarget ?? moveSource?.from ?? selectionRange.from)
              .marks(),
        );
      const dropped: number[] = [];
      for (const [index, original] of item.content.entries()) {
        if (original.type !== "drawing") continue;
        let child: Extract<DrawingContent, { rawXmlMode?: never }>;
        if (original.rawXmlMode !== undefined || embedsObject(original)) {
          const kept = keptOpaqueDrawing(original);
          if (kept.isErr()) return kept;
          if (kept.value !== undefined) {
            item.content[index] = kept.value;
            continue;
          }
          // Its parts cannot come with it: it pastes as the picture that
          // previews it, or, without one, not at all.
          const xml = original.rawXml ?? "";
          if (!hasPictureBytes(original.image)) {
            warn(CANONICAL_PASTE_WARNING_CODES.keptMarkupDropped, xml);
            dropped.push(index);
            continue;
          }
          warn(CANONICAL_PASTE_WARNING_CODES.keptMarkupAsPicture, xml);
          child = { type: "drawing", image: original.image };
          item.content[index] = child;
        } else child = original;
        delete child.rawXml;
        delete child.rawImageFingerprint;
        const image = child.image;
        delete image.id;
        if (
          moveTarget === undefined &&
          (image.hlinkRId !== undefined || image.hlinkHref !== undefined)
        ) {
          const previousId = image.hlinkRId;
          const imported = importExternalLink(image.hlinkHref, previousId);
          if (imported.isErr()) return imported;
          image.hlinkHref = imported.value.href;
          image.hlinkRId = imported.value.id;
          const remapped = new Map<string, string>();
          if (previousId !== undefined)
            remapped.set(previousId, imported.value.id);
          if (image.hlinkClickSource !== undefined) {
            const rebound = importCapturedLinks(
              image.hlinkClickSource.xml,
              remapped,
            );
            if (rebound.isErr()) return rebound;
            image.hlinkClickSource = {
              xml: rebound.value,
              rId: imported.value.id,
            };
          }
          if (image.hlinkHoverXml !== undefined) {
            const rebound = importCapturedLinks(image.hlinkHoverXml, remapped);
            if (rebound.isErr()) return rebound;
            image.hlinkHoverXml = rebound.value;
          }
        } else if (
          moveTarget === undefined &&
          image.hlinkHoverXml !== undefined
        ) {
          const rebound = importCapturedLinks(image.hlinkHoverXml, new Map());
          if (rebound.isErr()) return rebound;
          image.hlinkHoverXml = rebound.value;
        }
        if (
          moved !== undefined &&
          image.rId !== undefined &&
          moved.relationshipIds.has(image.rId)
        )
          continue;
        const relationship =
          image.rId === undefined ? undefined : relationships.get(image.rId);
        const ownedPath = relationship?.target.startsWith("/")
          ? relationship.target.slice(1)
          : `word/${relationship?.target ?? ""}`;
        const ownedMedia =
          relationship === undefined || sourceRelationships === undefined
            ? undefined
            : media.get(ownedPath);
        if (
          ownedMedia !== undefined &&
          (moveTarget !== undefined ||
            (image.src !== undefined && ownedMedia.dataUrl === image.src))
        )
          continue;
        const sourceRelationship =
          image.rId === undefined
            ? undefined
            : sourceRelationships?.get(image.rId);
        const sourcePath = sourceRelationship?.target.startsWith("/")
          ? sourceRelationship.target.slice(1)
          : `word/${sourceRelationship?.target ?? ""}`;
        const sourceMedia = sourceDocument?.package.media?.get(sourcePath);
        if (
          sourceMedia !== undefined &&
          sourceMedia.data.byteLength > MAX_CLIPBOARD_IMAGE_BYTES
        )
          return refuse(
            "The clipboard image exceeds the import size limit.",
            INVALID_CLIPBOARD,
          );
        const src =
          image.src ??
          (sourceMedia === undefined
            ? undefined
            : bytesToDataUrl(
                new Uint8Array(sourceMedia.data),
                sourceMedia.mimeType,
              )) ??
          (sourceRelationship?.targetMode === "External"
            ? sourceRelationship.target
            : undefined);
        if (src === undefined)
          return refuse(
            "The clipboard image has no source bytes or URL.",
            INVALID_CLIPBOARD,
          );
        image.src = src;
        const id = freshRelationshipId();
        const data =
          /^data:(image\/(?:png|jpeg|gif|webp));base64,([A-Za-z0-9+/=\s]+)$/u.exec(
            src,
          );
        if (data !== null) {
          const mimeType = data.at(1);
          const payload = data.at(2);
          if (mimeType === undefined || payload === undefined)
            return refuse(
              "The clipboard image data is malformed.",
              INVALID_CLIPBOARD,
            );
          if (payload.length > Math.ceil((MAX_CLIPBOARD_IMAGE_BYTES * 4) / 3))
            return refuse(
              "The clipboard image exceeds the import size limit.",
              INVALID_CLIPBOARD,
            );
          const decoded = Result.try({
            try: () => base64ToBytes(payload),
            catch: () =>
              new CanonicalSessionError({
                ground: INVALID_CLIPBOARD,
                message: "The clipboard image data is malformed.",
                reason: "refused",
              }),
          });
          if (decoded.isErr()) return decoded;
          const extension =
            mimeType === "image/jpeg" ? "jpg" : mimeType.slice(6);
          while (media.has(`word/media/clipboard${imageIndex}.${extension}`))
            imageIndex += 1;
          const path = `word/media/clipboard${imageIndex++}.${extension}`;
          media.set(path, {
            path,
            mimeType,
            data: decoded.value.buffer,
            dataUrl: src,
          });
          relationships.set(id, {
            id,
            type: IMAGE_RELATIONSHIP,
            target: path.slice(5),
          });
        } else if (/^https?:\/\//u.test(src)) {
          relationships.set(id, {
            id,
            type: IMAGE_RELATIONSHIP,
            target: src,
            targetMode: "External",
          });
        } else
          return refuse(
            "Clipboard images require image bytes or an HTTP image URL.",
            INVALID_CLIPBOARD,
          );
        const relationshipId = relationshipIdOf(id);
        if (relationshipId === undefined)
          return refuse(
            "The clipboard image relationship id is invalid.",
            INVALID_CLIPBOARD,
          );
        image.rId = relationshipId;
        // Drawing identities from another package must not claim local ownership.
        delete image.id;
        importedResource = true;
      }
      for (const index of dropped.toReversed()) item.content.splice(index, 1);
      if (dropped.length > 0) emptied.add(item);
    }
    withoutEmptiedRuns(paragraph, emptied);
    return Result.ok(undefined);
  };
  /**
   * Where kept markup's ids resolve: the relationships the source named them
   * in, those the destination resolves them against, and the copier that
   * copies the parts they name from the source package, read ahead of the
   * paste, under fresh names and relationship ids. `part` names the note,
   * comment, header or footer part the markup lands in, and is absent for
   * the body.
   */
  type KeptMarkupScope = {
    part?: string;
    source: ReadonlyMap<string, Relationship> | undefined;
    destination: Map<string, Relationship>;
    copier: ReturnType<typeof createPartCopier> | undefined;
  };
  const copiedPaths = new Map<string, string>();
  const copierOf = ({
    origin,
    into,
    sourcePart,
    destinationPart,
  }: {
    origin: ReadonlyMap<string, Relationship> | undefined;
    into: Map<string, Relationship>;
    sourcePart: string | undefined;
    destinationPart: string | undefined;
  }) => {
    if (
      sourceDocument === undefined ||
      origin === undefined ||
      sourcePart === undefined ||
      destinationPart === undefined
    )
      return undefined;
    // Without the parts read from the source package, the copier copies
    // relationships to external targets alone.
    const parts: ClipboardPackageParts =
      sourceParts?.source === sourceDocument
        ? sourceParts
        : { source: sourceDocument, parts: new Map() };
    return createPartCopier({
      parts,
      sourceRelationships: origin,
      sourcePart,
      destination,
      relationships: into,
      destinationPart,
      copiedParts,
      freshRelationshipId,
      copiedPaths,
    });
  };
  // The story parts whose relationships this paste holds its own copy of.
  const ownParts = new Set<string>();
  /** A story part's relationships, the paste's own copy of them. */
  const partRelationships = (part: string): Map<string, Relationship> => {
    const held = annotationRelationships.get(part);
    if (held !== undefined && ownParts.has(part)) return held;
    const into = new Map(held);
    annotationRelationships.set(part, into);
    ownParts.add(part);
    return into;
  };
  // The slice's markup names the relationships of the story it was copied
  // from, and lands in those of the story the paste goes into: the main
  // document's for the body, the story part's own otherwise.
  const copiedFrom =
    sourceDocument === undefined
      ? undefined
      : storyRelationshipsOf(sourceDocument, sourceStory);
  const landingPart = storyRelationshipsKey(story);
  const landing =
    landingPart === undefined ? relationships : partRelationships(landingPart);
  let scope: KeptMarkupScope = {
    ...(landingPart === undefined ? {} : { part: landingPart }),
    source: copiedFrom?.relationships,
    destination: landing,
    copier: copierOf({
      origin: copiedFrom?.relationships,
      into: landing,
      sourcePart: copiedFrom?.partPath,
      destinationPart: storyPartPath(destination, story),
    }),
  };
  const annotationScopes = new Map<string, KeptMarkupScope>();
  /** The scope of markup in a copied note or comment: its part's own. */
  const annotationScope = (part: string): KeptMarkupScope => {
    const held = annotationScopes.get(part);
    if (held !== undefined) return held;
    // A part without its own relationships file resolves its ids through the
    // main document's, as the parser reads it.
    const origin =
      sourceDocument?.package.annotationRelationships?.get(part) ??
      sourceDocument?.package.relationships;
    // Copying into such a part gives it a relationships file of its own,
    // which the parser then reads instead of the main document's: the new
    // table starts from the main document's relationships, so the ids the
    // part already names keep resolving to what they named.
    const hadOwn = annotationRelationships.has(part);
    const into = partRelationships(part);
    if (!hadOwn)
      for (const [id, relationship] of relationships)
        if (!into.has(id)) into.set(id, relationship);
    const created = {
      part,
      source: origin,
      destination: into,
      copier: copierOf({
        origin,
        into,
        sourcePart: `word/${part}.xml`,
        destinationPart: `word/${part}.xml`,
      }),
    };
    annotationScopes.set(part, created);
    return created;
  };
  const samePackage =
    sourceDocument !== undefined &&
    sourceDocument.originalBuffer !== undefined &&
    sourceDocument.originalBuffer === destination.originalBuffer;
  /** Whether every id names in the destination what it named in the source. */
  const resolvesInPlace = (ids: readonly string[]) =>
    samePackage &&
    ids.every((id) =>
      sameRelationship(scope.source?.get(id), scope.destination.get(id)),
    );
  /**
   * Whether a picture's captured markup is an embedded object around it: the
   * markup names a part besides the picture's own. It pastes as the object
   * where the parts it names paste with it, and as the picture otherwise.
   */
  const embedsObject = (drawing: DrawingContent): boolean => {
    const xml = drawing.rawXml;
    if (xml === undefined) return false;
    const referenced = relationshipIdsIn(xml);
    if (referenced.isErr()) return false;
    const own = new Set([drawing.image.rId, drawing.image.hlinkRId]);
    return !referenced.value.every((id) => own.has(id));
  };
  /** Whether a picture has bytes or a URL the paste can import it by. */
  const hasPictureBytes = (image: DrawingContent["image"]): boolean => {
    if (image.src !== undefined)
      return IMPORTABLE_PICTURE_SOURCE.test(image.src);
    const relationship =
      image.rId === undefined ? undefined : sourceRelationships?.get(image.rId);
    if (relationship === undefined) return false;
    if (relationship.targetMode === "External")
      return IMPORTABLE_PICTURE_SOURCE.test(relationship.target);
    const path = relationship.target.startsWith("/")
      ? relationship.target.slice(1)
      : `word/${relationship.target}`;
    return sourceDocument?.package.media?.has(path) === true;
  };
  /** Record what the paste did in place of keeping `xml`. */
  const warn = (code: CanonicalPasteWarning["code"], xml: string): void => {
    warnings.push({
      code,
      part: scope.part ?? "document",
      element: /^<(?<name>[^\s/>]+)/u.exec(xml)?.groups?.["name"] ?? "",
    });
  };
  /**
   * Kept markup (an opaque drawing, an embedded object, a kept block), with
   * the relationship ids it names renamed to the copies it pastes with. Its
   * parts (a chart, an embedding, a diagram) are carried by the package it
   * was read from, not by the model: in that package every relationship it
   * names is the one it named, and into another the paste copies them. In a
   * note, a comment, a header or a footer its ids are those of the part
   * holding it. Undefined when the parts cannot come with it: the clipboard
   * came from outside the editor, or its source package could not be read.
   */
  const keptMarkupPlaced = (
    xml: string,
  ): Result<
    { xml: string; ids: ReadonlyMap<string, string> } | undefined,
    CanonicalSessionError
  > => {
    const unchanged = Result.ok({ xml, ids: new Map<string, string>() });
    const referenced = relationshipIdsIn(xml);
    if (referenced.isErr()) return referenced;
    if (referenced.value.length === 0 || moved?.markup.has(xml) === true)
      return unchanged;
    if (resolvesInPlace(referenced.value)) return unchanged;
    if (scope.copier === undefined) return Result.ok(undefined);
    const copies = scope.copier.copy(referenced.value);
    if (copies.isErr())
      return copies.error.type === "unresolved"
        ? refuse(
            "Copied markup names a relationship its package does not hold.",
            INVALID_CLIPBOARD,
          )
        : Result.ok(undefined);
    const renamed = withRelationshipIds(xml, copies.value);
    if (renamed === undefined)
      return refuse(
        "The clipboard block markup is malformed.",
        INVALID_CLIPBOARD,
      );
    if (scope.part === undefined) importedResource = true;
    else changedAnnotationParts.add(scope.part);
    copiedPart = true;
    return Result.ok({ xml: renamed, ids: copies.value });
  };
  /**
   * An opaque drawing with the parts it names, under a fresh drawing id;
   * undefined when its parts cannot come with it.
   */
  const keptOpaqueDrawing = (
    drawing: DrawingContent,
  ): Result<DrawingContent | undefined, CanonicalSessionError> => {
    const placed = keptMarkupPlaced(drawing.rawXml ?? "");
    if (placed.isErr()) return placed;
    if (placed.value === undefined) return Result.ok(undefined);
    drawingIds ??= usedDrawingIds({
      documentBody: destination.package.document,
      headers: destination.package.headers,
      footers: destination.package.footers,
      footnotes: destination.package.footnotes,
      endnotes: destination.package.endnotes,
    });
    const id = takeDrawingId(drawingIds);
    const kept = structuredClone(drawing);
    const { ids } = placed.value;
    if (ids.size > 0) {
      const image = kept.image;
      const current =
        kept.rawXmlMode === DRAWING_RAW_XML_MODES.PREVIEW_ONLY &&
        kept.rawImageFingerprint === imageRawXmlFingerprint(image);
      kept.rawXml = placed.value.xml;
      const pictureId =
        image.rId === undefined ? undefined : ids.get(image.rId);
      const rId =
        pictureId === undefined ? undefined : relationshipIdOf(pictureId);
      if (rId !== undefined) image.rId = rId;
      const hlinkRId =
        image.hlinkRId === undefined ? undefined : ids.get(image.hlinkRId);
      if (hlinkRId !== undefined) image.hlinkRId = hlinkRId;
      if (image.hlinkClickSource !== undefined) {
        const source = image.hlinkClickSource;
        source.xml = withRelationshipIds(source.xml, ids) ?? source.xml;
        const sourceRId =
          source.rId === undefined ? undefined : ids.get(source.rId);
        if (sourceRId !== undefined) source.rId = sourceRId;
      }
      if (image.hlinkHoverXml !== undefined)
        image.hlinkHoverXml =
          withRelationshipIds(image.hlinkHoverXml, ids) ?? image.hlinkHoverXml;
      // A preview the renamed ids leave current stays current.
      if (current && kept.rawXmlMode === DRAWING_RAW_XML_MODES.PREVIEW_ONLY)
        kept.rawImageFingerprint = imageRawXmlFingerprint(image);
    }
    kept.image.id = id;
    const renumbered = withDrawingId(kept, id).rawXml;
    if (renumbered !== undefined) kept.rawXml = renumbered;
    return Result.ok(kept);
  };
  /**
   * Kept block markup with the parts it names, or, where they cannot come
   * with it, the paragraphs of text it holds; none when it holds no text.
   */
  const keptBlockPlaced = (
    xml: string,
    readerText: string | undefined,
  ): Result<{ xml: string } | Paragraph[], CanonicalSessionError> => {
    const placed = keptMarkupPlaced(xml);
    if (placed.isErr()) return placed;
    if (placed.value !== undefined) return Result.ok({ xml: placed.value.xml });
    const texts = keptBlockTexts(xml, readerText);
    warn(
      texts.length === 0
        ? CANONICAL_PASTE_WARNING_CODES.keptMarkupDropped
        : CANONICAL_PASTE_WARNING_CODES.keptMarkupAsText,
      xml,
    );
    return Result.ok(texts.map(textParagraph));
  };
  /** Every kept block among `container`, at any depth, placed or replaced. */
  const placeKeptBlocks = (
    container: BlockContent[],
  ): Result<void, CanonicalSessionError> => {
    for (let index = container.length - 1; index >= 0; index -= 1) {
      const block = container[index];
      if (block === undefined) continue;
      if (block.type === "preservedBlock") {
        const placed = keptBlockPlaced(block.xml, block.readerText);
        if (placed.isErr()) return placed;
        if (Array.isArray(placed.value))
          container.splice(index, 1, ...placed.value);
        else block.xml = placed.value.xml;
        continue;
      }
      const inner: BlockContent[][] = [];
      if (block.type === "table")
        for (const row of block.rows)
          for (const cell of row.cells) inner.push(cell.content);
      if (block.type === "blockSdt" || block.type === "blockCustomXml")
        inner.push(block.content);
      for (const content of inner) {
        const placedInner = placeKeptBlocks(content);
        if (placedInner.isErr()) return placedInner;
        // A cell or a container holds at least a paragraph.
        if (content.length === 0) content.push(textParagraph(""));
      }
    }
    return Result.ok(undefined);
  };
  for (const paragraph of paragraphs) {
    const imported = importResources(paragraph);
    if (imported.isErr()) return imported;
  }
  // Kept block markup replays as it was read: what it names must be there,
  // or it pastes as the text it holds.
  for (let index = blocks.length - 1; index >= 0; index -= 1) {
    const block = blocks[index];
    if (block === undefined) continue;
    if (block.type !== "preservedBlock") {
      const placedInner = placeKeptBlocks([block]);
      if (placedInner.isErr()) return placedInner;
      continue;
    }
    const placed = keptBlockPlaced(block.xml, block.readerText);
    if (placed.isErr()) return placed;
    if (Array.isArray(placed.value)) blocks.splice(index, 1, ...placed.value);
    else block.xml = placed.value.xml;
  }
  if (blocks.length === 0)
    return Result.err(
      new CanonicalSessionError({
        ground: INVALID_CLIPBOARD,
        message: "The clipboard contains no content the document can hold.",
        reason: "noChange",
      }),
    );
  // A note's or a comment's ids name its own part's relationships: its links
  // import by their targets and its images by their bytes, and the save
  // writes each into the part that holds it. Its kept markup resolves in, or
  // copies its parts into, that part's own relationships.
  sourceRelationships = undefined;
  for (const note of annotations?.notes ?? []) {
    scope = annotationScope(
      note.type === "footnote" ? "footnotes" : "endnotes",
    );
    for (const paragraph of nestedParagraphs(note.content)) {
      const imported = importResources(paragraph);
      if (imported.isErr()) return imported;
    }
    const placed = placeKeptBlocks(note.content);
    if (placed.isErr()) return placed;
    if (note.content.length === 0) note.content.push(textParagraph(""));
  }
  for (const comment of annotations?.comments ?? []) {
    scope = annotationScope("comments");
    for (const paragraph of comment.content) {
      const imported = importResources(paragraph);
      if (imported.isErr()) return imported;
    }
    // A comment's kept markup replays as it was read, or, where its parts
    // cannot come with it, as the paragraphs it holds in its place.
    const children = comment.preserved?.children ?? [];
    const keptChildren: typeof children = [];
    let inserted = 0;
    for (const child of children) {
      const placed = keptBlockPlaced(child.xml, undefined);
      if (placed.isErr()) return placed;
      const at = child.index + inserted;
      if (Array.isArray(placed.value)) {
        comment.content.splice(at, 0, ...placed.value);
        inserted += placed.value.length;
        continue;
      }
      keptChildren.push({ ...child, index: at, xml: placed.value.xml });
    }
    // A container keeps no empty record of markup it does not hold.
    if (keptChildren.length > 0 && comment.preserved !== undefined)
      comment.preserved.children = keptChildren;
    else delete comment.preserved;
  }
  const resourceOps: DocumentOp[] = [];
  if (
    importedResource ||
    copiedPart ||
    numbering !== destination.package.numbering ||
    styles !== destination.package.styles
  ) {
    // Only the parts a paste copied into change; the others stay as held.
    const changedTables = new Map(destination.package.annotationRelationships);
    for (const part of changedAnnotationParts) {
      const held =
        annotationRelationships.get(part) ??
        panic("A changed story part has no relationships.");
      changedTables.set(part, held);
    }
    const imported = {
      ...destination,
      package: {
        ...destination.package,
        ...(numbering === undefined ? {} : { numbering }),
        ...(styles === undefined ? {} : { styles }),
        ...(importedResource ? { relationships, media } : {}),
        ...(copiedPart ? { copiedParts } : {}),
        ...(changedAnnotationParts.size > 0
          ? { annotationRelationships: changedTables }
          : {}),
      },
    };
    resourceOps.push(
      packageResourcesOpOf({ before: destination, after: imported }),
    );
  }
  const openStart = importedSlice.openStart === 0 ? 0 : 1;
  const openEnd = importedSlice.openEnd === 0 ? 0 : 1;
  // Without a table the blocks are paragraphs, pasted as one fragment.
  const topParagraphs = blocks.filter(
    (block): block is Paragraph => block.type === "paragraph",
  );
  // Notes and comments, or a table, need the block paste.
  const pasted = (range: {
    from: TextPosition;
    to: TextPosition;
  }): EditorIntent => {
    if (annotations !== undefined)
      return {
        type: "pasteBlocks",
        ...range,
        blocks,
        openStart,
        openEnd,
        annotations,
      };
    if (topParagraphs.length === blocks.length)
      return {
        type: "replaceFragment",
        ...range,
        paragraphs: topParagraphs,
        openStart,
        openEnd,
      };
    return { type: "pasteBlocks", ...range, blocks, openStart, openEnd };
  };
  const withWarnings = (
    prepared: Result<CanonicalCommit, CanonicalSessionError>,
  ): Result<CanonicalPasteCommit, CanonicalSessionError> =>
    prepared.map((commit): CanonicalPasteCommit =>
      Object.assign(commit, { warnings }),
    );
  if (from.value === undefined || to.value === undefined)
    return withWarnings(
      session.prepareCellSelectionStructure(
        state,
        (at) => pasted({ from: at, to: at }),
        { resourceOps, semantic: "paste", story },
      ),
    );
  const range = { from: from.value, to: to.value };
  if (moveTarget === undefined)
    return withWarnings(
      session.prepareIntents(state, {
        intents: [pasted(range)],
        resourceOps,
        semantic: "paste",
      }),
    );
  const target = projection.addressAt(moveTarget);
  if (target.isErr()) return target;
  // Dragged paragraphs move as a fragment; with a table or a container, as
  // blocks.
  return withWarnings(
    session.prepareIntents(state, {
      intents: [
        topParagraphs.length === blocks.length
          ? {
              type: "moveFragment",
              target: target.value,
              ...range,
              paragraphs: topParagraphs,
              openStart,
              openEnd,
            }
          : {
              type: "moveBlocks",
              target: target.value,
              ...range,
              blocks,
              openStart,
              openEnd,
            },
      ],
      resourceOps,
      semantic: "paste",
    }),
  );
};
