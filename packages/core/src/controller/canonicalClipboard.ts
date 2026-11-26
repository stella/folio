import { panic, Result } from "better-result";
import { Fragment, type Node as ProseNode, Slice } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import { CellSelection, tableNodeTypes } from "prosemirror-tables";

import {
  NUMBER_FORMATS,
  paragraphNumberingReference,
  relationshipIdOf,
} from "@stll/docx-core/model";
import {
  type DocumentOp,
  type EditorIntent,
  packageResourcesOpOf,
  type PasteAnnotations,
  type TextPosition,
} from "@stll/docx-core/ops";

import { groundForDocumentOpRefusal } from "../canonical/canonicalRefusalGrounds";
import {
  type CanonicalCommit,
  type CanonicalSession,
  CanonicalSessionError,
} from "../canonical/canonicalSession";
import {
  createNumberingIdAllocator,
  numberingNumIdsTaken,
} from "../docx/numberingIds";
import { copyParagraphIndentationProjection } from "../docx/paragraphPropertySource";
import {
  visitInlineContentSlots,
  visitParagraphRuns,
} from "../docx/paragraphTraversal";
import { captureVerbatimXml } from "../docx/verbatimCapture";
import {
  findAttributeByNamespaceUri,
  getLocalName,
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
  CANONICAL_GAP,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";
import type {
  CounterFormat,
  Document,
  NumberFormat,
  NumberingDefinitions,
  Paragraph,
  Run,
  Table,
} from "../types/document";
import { readModel } from "../types/documentModel";
import { bytesToDataUrl } from "../utils/base64";
import { base64ToBytes } from "../utils/textEncoding";
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
 * The gap a clipboard block the paste cannot place falls under: a block
 * other than a paragraph or a table, at the top or in a cell, is story
 * content the session cannot project.
 */
const UNPLACED_BLOCK: CanonicalRefusalGround = {
  gap: CANONICAL_GAP.storyContentProjection,
};

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
type ClipboardParagraphPlace = {
  paragraph: Paragraph;
  replace: (next: Paragraph) => void;
};

/**
 * A pasted block's paragraphs in document order, a table's cell paragraphs
 * among them, each with the place it fills; undefined when a cell holds a
 * block other than a paragraph or a table.
 */
const clipboardParagraphs = (
  block: Paragraph | Table,
  replace: (next: Paragraph) => void = () => {},
): ClipboardParagraphPlace[] | undefined => {
  if (block.type === "paragraph") return [{ paragraph: block, replace }];
  const places: ClipboardParagraphPlace[] = [];
  for (const row of block.rows)
    for (const cell of row.cells)
      for (const [index, inner] of cell.content.entries()) {
        if (inner.type !== "paragraph" && inner.type !== "table")
          return undefined;
        const found = clipboardParagraphs(inner, (next) => {
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
/** Every table among the blocks, those nested in cells included. */
const clipboardTables = (blocks: readonly (Paragraph | Table)[]): Table[] =>
  blocks.flatMap((block) =>
    block.type === "table"
      ? [
          block,
          ...clipboardTables(
            block.rows.flatMap((row) =>
              row.cells.flatMap((cell) =>
                cell.content.filter(
                  (inner): inner is Paragraph | Table =>
                    inner.type === "paragraph" || inner.type === "table",
                ),
              ),
            ),
          ),
        ]
      : [],
  );

const withoutTableCaptures = (table: Table): void => {
  if (table.formatting !== undefined) {
    delete table.formatting.sourceXml;
    delete table.formatting.gridSourceXml;
  }
  for (const row of table.rows) {
    if (row.formatting !== undefined) delete row.formatting.sourceXml;
    for (const cell of row.cells) {
      if (cell.formatting !== undefined) delete cell.formatting.sourceXml;
      for (const inner of cell.content)
        if (inner.type === "table") withoutTableCaptures(inner);
    }
  }
};

/** A pasted block's own copy, each paragraph keeping its indentation provenance. */
const cloneClipboardBlock = (block: Paragraph | Table): Paragraph | Table => {
  const clone = structuredClone(block);
  if (clone.type === "table") withoutTableCaptures(clone);
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
 * The notes and comments the clipboard's references and markers name, read
 * from the document they were copied from; the paste copies each under a
 * fresh id. Without that document the ids could name the destination's own
 * notes and comments, so the paste refuses.
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
  const refusal = (message: string) =>
    Result.err(
      new CanonicalSessionError({
        ground: { gap: CANONICAL_GAP.storyContentProjection },
        reason: "refused",
        message,
      }),
    );
  if (source === undefined)
    return refusal(
      "Clipboard comments and note references require the document they were copied from.",
    );
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
    return refusal(
      "A clipboard comment or note reference names none in the document it was copied from.",
    );
  return Result.ok(annotations);
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
  moveSource?: { from: number; to: number };
};

export const prepareCanonicalPaste = ({
  session,
  state,
  slice: copied,
  plain = false,
  moveTarget,
  pasteTarget,
  sourceDocument,
  moveSource,
}: PrepareCanonicalPasteOptions): Result<
  CanonicalCommit,
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
  // cells it holds.
  const slice = withClosedTableEdges(copied);
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
  const tableType = state.schema.nodes["table"];
  let inline = true;
  let blocksOnly = true;
  slice.content.forEach((node) => {
    inline &&= node.isInline;
    blocksOnly &&= node.type === paragraphType || node.type === tableType;
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
  const converted = Result.try({
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
  if (converted.isErr()) return converted;
  const blocks: (Paragraph | Table)[] = [];
  for (const block of converted.value) {
    if (
      (block.type !== "paragraph" && block.type !== "table") ||
      clipboardParagraphs(block) === undefined
    )
      return refuse(
        "The clipboard contains an unsupported block.",
        UNPLACED_BLOCK,
      );
    blocks.push(cloneClipboardBlock(block));
  }
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
  const from = overCells
    ? Result.ok(undefined)
    : session.projection.addressAt(
        pasteTarget ?? moveSource?.from ?? selectionRange.from,
      );
  if (from.isErr()) return from;
  const to = overCells
    ? Result.ok(undefined)
    : session.projection.addressAt(
        pasteTarget ?? moveSource?.to ?? selectionRange.to,
      );
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
  // A note or comment part owns its relationships; a copied body naming an
  // image or an external link would name one its new part does not hold.
  for (const paragraph of annotationParagraphs) {
    let related = false;
    visitInlineContentSlots(paragraph, ({ item }) => {
      if (item.type === "hyperlink" && item.anchor === undefined)
        related = true;
    });
    visitParagraphRuns(paragraph, (run) => {
      if (run.content.some((child) => child.type === "drawing")) related = true;
    });
    if (related)
      return refuse(
        "Copied notes and comments holding images or external links require their part's relationships.",
        { gap: CANONICAL_GAP.storyContentProjection },
      );
  }
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
  let importedResource = false;
  let imageIndex = 1;
  let relationshipIndex = 1;
  const importExternalLink = (
    href: string | undefined,
    previousId: string | undefined,
  ) => {
    const source =
      previousId === undefined
        ? undefined
        : sourceDocument?.package.relationships?.get(previousId);
    let target = href;
    if (target === undefined || target.length === 0)
      target = source?.targetMode === "External" ? source.target : undefined;
    if (target === undefined || target.length === 0)
      return refuse(
        "The clipboard hyperlink has no resolved target.",
        INVALID_CLIPBOARD,
      );
    while (relationships.has(`rId${relationshipIndex}`)) relationshipIndex += 1;
    const id = `rId${relationshipIndex++}`;
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
  // New inline images have no opaque XML dependencies. Existing package images
  // retain their relationship only when its resolved source belongs to this package.
  for (const paragraph of paragraphs) {
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
    for (const item of runs) {
      if (plain)
        item.formatting = marksToTextFormatting(
          state.storedMarks ??
            state.doc
              .resolve(pasteTarget ?? moveSource?.from ?? selectionRange.from)
              .marks(),
        );
      for (const child of item.content) {
        if (child.type !== "drawing") continue;
        if (child.rawXmlMode !== undefined)
          return refuse(
            "Opaque drawings and embedded objects require their source package parts.",
            { gap: CANONICAL_GAP.storyContentProjection },
          );
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
        const relationship =
          image.rId === undefined ? undefined : relationships.get(image.rId);
        const ownedPath = relationship?.target.startsWith("/")
          ? relationship.target.slice(1)
          : `word/${relationship?.target ?? ""}`;
        const ownedMedia =
          relationship === undefined ? undefined : media.get(ownedPath);
        if (
          ownedMedia !== undefined &&
          (moveTarget !== undefined ||
            (image.src !== undefined && ownedMedia.dataUrl === image.src))
        )
          continue;
        const sourceRelationship =
          image.rId === undefined
            ? undefined
            : sourceDocument?.package.relationships?.get(image.rId);
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
        while (relationships.has(`rId${relationshipIndex}`))
          relationshipIndex += 1;
        const id = `rId${relationshipIndex++}`;
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
    }
  }
  const resourceOps: DocumentOp[] = [];
  if (
    importedResource ||
    numbering !== destination.package.numbering ||
    styles !== destination.package.styles
  ) {
    const imported = {
      ...destination,
      package: {
        ...destination.package,
        ...(numbering === undefined ? {} : { numbering }),
        ...(styles === undefined ? {} : { styles }),
        ...(importedResource ? { relationships, media } : {}),
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
  if (from.value === undefined || to.value === undefined)
    return session.prepareCellSelectionStructure(
      state,
      (at) => pasted({ from: at, to: at }),
      { resourceOps, semantic: "paste" },
    );
  const range = { from: from.value, to: to.value };
  if (moveTarget === undefined)
    return session.prepareIntents(state, {
      intents: [pasted(range)],
      resourceOps,
      semantic: "paste",
    });
  const target = session.projection.addressAt(moveTarget);
  if (target.isErr()) return target;
  // Dragged paragraphs move as a fragment; with a table, as blocks.
  return session.prepareIntents(state, {
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
  });
};
