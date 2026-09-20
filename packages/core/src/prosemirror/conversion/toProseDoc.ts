/**
 * Document to ProseMirror Conversion
 *
 * Converts our Document type (from DOCX parsing) to a ProseMirror document.
 * Preserves all formatting attributes for round-trip fidelity.
 *
 * Style Resolution:
 * When styles are provided, paragraph properties are resolved from the style chain:
 * - Document defaults (docDefaults)
 * - Normal style (if no explicit styleId)
 * - Style chain (basedOn inheritance)
 * - Inline properties (highest priority)
 */

import type { Node as PMNode } from "prosemirror-model";
import { panic } from "better-result";
import { PARSE_WARNING_CODES } from "@stll/docx-core/model";

import type { ParseContext } from "../../docx/parseContext";
import { createStyleEngine } from "../../style-engine";
import type { StyleEngine, TableCellParagraphSpacingOverlay } from "../../style-engine";
import type {
  InlineWrapper,
  BlockContent,
  BlockSdt,
  Document,
  Paragraph,
  ParagraphFormatting,
  PreservedAttribute,
  PreservedBlock,
  PreservedInline,
  Run,
  RunPropertyChange,
  TableCellBlock,
  TextFormatting,
  RunContent,
  Hyperlink,
  Image,
  TextBox,
  Shape,
  StyleDefinitions,
  Table,
  TableRow,
  TableRowFormatting,
  TableCell,
  TableCellFormatting,
  TableBorders,
  SimpleField,
  ComplexField,
  InlineSdt,
  Insertion,
  Deletion,
  DrawingAnchor,
  WrapDistanceSlots,
  DrawingContent,
  MoveFrom,
  MoveTo,
  MathEquation,
  ParagraphContent,
  ShapeTextBody,
  Theme,
} from "../../types/document";
import { resolveTableLook, type ResolvedTableLook } from "../../docx/tableLook";
import {
  mergeParagraphFormatting,
  mergeParagraphTabStops,
} from "../../utils/paragraphFormattingMerge";
import { rangedCommentIds } from "../../docx/commentAnchorIndex";
import { isInlineSdtContent, isTrackedChangeWrapperChild } from "../../docx/inlineWrapperContent";
import { resolveColorValueToHex } from "../../docx/drawingUtils";
import { copiedWrapPolygon } from "../../docx/wrapPolygon";
import { isNumberingReference, NO_NUMBERING_NUM_ID } from "../../docx/numberingReference";
import { isCellMergeContinuation } from "../../docx/tableParser";
import { isBaselineVertAlign } from "../../docx/runParser";
import {
  PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR,
  createProseParagraphWithPropertySource,
  getDocumentParagraphPropertySourceContract,
  recreateProseNodeWithParagraphPropertySource,
  transportTableCellsWithParagraphPropertySources,
} from "../../docx/paragraphPropertySource";
import {
  buildPageBreakRunSourceDescendantIndex,
  type PageBreakRunSourceDescendantIndex,
} from "../../internal/pageBreakRunSourceDescendantIndex";
import { DRAWING_RAW_XML_MODES } from "@stll/docx-core/model";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";
import { tableOfContentsStyleLevel } from "../../utils/tableOfContentsStyle";
import { emuToPixels, emuToStrokePixels } from "../../utils/units";
import { normalizeHorizontalScalePercent } from "../../utils/horizontalScale";
import { authoredTransformAttrs } from "../authoredTransformAttrs";
import { expectInlineWrapperMarkAttrs } from "../attrs";
import { setAutospacingBaseValue } from "../autospacingBase";
import {
  textFormattingToMarks,
  type AuthoredRunFormattingCarrier,
} from "../extensions/marks/markUtils";
import { inlineWrapperLayer } from "../inlineWrapperStack";
import { directionFromBidi } from "../paragraphDirection";
import { styleResolvedParagraphFormatting } from "../paragraphFormattingProvenance";
import { pageBreakRunParagraphProjectionDispositionForFeatures } from "../pageBreakRunProjection";
import { lineSpacingProvenanceFromSpacing } from "../paragraphSpacing";
import {
  getParagraphMarkSuppressionOverrides,
  hasDirectRunFormatting,
  resolveParagraphBodyRunFormatting,
  stripParagraphMarkFormattingForBodyRuns,
  stripParagraphMarkOnlyFormatting,
  suppressParagraphMarkFormatting,
} from "../runStyleFormatting";
import { schema } from "../schema";
import { RUN_FORMATTING_PROPERTY_SPECS, type InlineWrapperLayer } from "../schema/marks";
import { cascadeStyleTextFormatting } from "../styles/styleToggleCascade";
import { PRESERVED_XML_LEVELS } from "../schema/nodes";
import type {
  ImagePositionAttrs,
  ParagraphAttrs,
  TableAttrs,
  TableRowAttrs,
  TableCellAttrs,
  TextBoxAttrs,
} from "../schema/nodes";
import { assertValidProseMirrorDocument } from "../validation";
import { listRenderingAttrPatch } from "../listRenderingAttrs";
import { stampNumberedRefFieldBaselines } from "../numberedRefFields";
import { canCarryTrackedRunMark, trackedRunInlineAtomDisposition } from "../trackedRunInlineAtoms";
import {
  resolveEffectiveTableCellFormatting,
  type TableCellMarginsAttrs,
  type TableCellPosition,
} from "./effectiveTableCellFormatting";
import { sdtAttrsFromProperties } from "./sdtAttrs";

const DETACHED_WATERMARK_HOST = Symbol.for("stll.detachedWatermarkHost");

/**
 * Options for document conversion
 */
export type ToProseDocOptions = {
  /** Style definitions for resolving paragraph styles */
  styles?: StyleDefinitions;
  /** Theme used when converting themed table/cell values in nested content. */
  theme?: Theme | null;
  /**
   * Where a projection approximation is reported, on the channel the parsers
   * already use. Absent means nobody is listening, not that nothing happened.
   */
  warn?: ParseContext["warn"];
};

/** Records that a page break is projected less than exactly. */
type PageBreakProjectionWarn = (detail: string) => void;

const noPageBreakProjectionWarning: PageBreakProjectionWarn = () => {};

const pageBreakProjectionWarn = (
  warn: ParseContext["warn"] | undefined,
): PageBreakProjectionWarn =>
  warn === undefined
    ? noPageBreakProjectionWarning
    : (detail) => {
        warn({ code: PARSE_WARNING_CODES.pageBreakProjectionApproximated, detail });
      };

type ResolvedRunFormatting = {
  formatting: TextFormatting | undefined;
  implicitCharacterStyleApplied?: true;
  paragraphMarkOverrides?: TextFormatting;
  toggleCascade: ReturnType<typeof cascadeStyleTextFormatting>;
};

type RunFormattingResolver = (
  formatting: TextFormatting | undefined,
  fieldType?: string,
) => ResolvedRunFormatting;

type TrackedRunFormattingResolvers = {
  current: RunFormattingResolver;
  historical: RunFormattingResolver;
};

/**
 * Build a `nextTextBoxGroupId()` generator salted with a random per-load
 * nonce, so minted text-box anchor ids (`<salt>:<group>:<index>`) are unique
 * across separate conversions instead of purely sequential ("0:0", "0:1",
 * ...) starting fresh every load. `TextBoxAnchorExtension` parses
 * `data-docx-textbox-anchor` straight off pasted DOM and `fromProseDoc`
 * registers anchors first-seen-wins — without the salt, a pasted external
 * span carrying a guessed/copied id could collide with and hijack a real
 * text box anchored in a completely different document.
 */
const createTextBoxGroupIdFactory = (): (() => string) => {
  const salt = Math.random().toString(36).slice(2, 10);
  let index = 0;
  return () => `${salt}:${index++}`;
};

type HyperlinkInstanceIndexAllocator = () => number;
type PageBreakRunOwnerIdAllocator = () => number;

/** Keep imported hyperlink identity unique across every nested conversion scope. */
const createHyperlinkInstanceIndexAllocator = (): HyperlinkInstanceIndexAllocator => {
  let index = 0;
  return () => index++;
};

type BookmarkBoundaryCount = {
  starts: number;
  ends: number;
  firstStart?: number;
  firstEnd?: number;
};

/**
 * Find bookmark pairs across every editable inline wrapper in a document story.
 * Each endpoint is converted at its own structural position, so a range may
 * start outside a hyperlink and end inside it (or the inverse).
 */
const collectPairedBookmarkIds = (blocks: readonly BlockContent[]): ReadonlySet<number> => {
  const counts = new Map<number, BookmarkBoundaryCount>();
  let position = 0;
  const countBoundary = (id: number, type: "start" | "end"): void => {
    const count = counts.get(id) ?? { starts: 0, ends: 0 };
    if (type === "start") {
      count.starts += 1;
      count.firstStart ??= position;
    } else {
      count.ends += 1;
      count.firstEnd ??= position;
    }
    position += 1;
    counts.set(id, count);
  };

  const visitRun = (run: Run): void => {
    for (const content of run.content) {
      if (content.type === "shape" && content.shape.textBody) {
        visitBlocks(content.shape.textBody.content);
      }
    }
  };

  const visitHyperlink = (hyperlink: Hyperlink): void => {
    for (const child of hyperlink.children) {
      switch (child.type) {
        case "bookmarkStart":
          countBoundary(child.id, "start");
          break;
        case "bookmarkEnd":
          countBoundary(child.id, "end");
          break;
        case "run":
          visitRun(child);
          break;
        // Opaque markup: it anchors no bookmark and holds no run.
        case "preservedInline":
          break;
        default: {
          const unsupported: never = child;
          panic(`Unsupported hyperlink child: ${JSON.stringify(unsupported)}`);
        }
      }
    }
  };

  const visitParagraphContent = (content: Paragraph["content"][number]): void => {
    switch (content.type) {
      case "bookmarkStart":
        countBoundary(content.id, "start");
        return;
      case "bookmarkEnd":
        countBoundary(content.id, "end");
        return;
      case "run":
        visitRun(content);
        return;
      case "hyperlink":
        visitHyperlink(content);
        return;
      case "simpleField":
        for (const child of content.content) {
          if (child.type === "hyperlink") {
            visitHyperlink(child);
          } else if (child.type === "run") {
            visitRun(child);
          }
        }
        return;
      case "complexField":
        for (const run of [...(content.fieldCode ?? []), ...content.fieldResult]) {
          visitRun(run);
        }
        return;
      case "inlineSdt":
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
      // A transparent wrapper is transparent to bookmark pairing too: the
      // projection lifts it and its children become the paragraph's own
      // inline sequence, so a boundary inside one is converted at the
      // paragraph's level and has to be counted there.
      case "inlineWrapper":
        for (const child of content.content) {
          visitParagraphContent(child);
        }
        return;
      // Range markers and equations hold no bookmark boundary.
      case "commentRangeStart":
      case "commentRangeEnd":
      case "commentReference":
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
      case "mathEquation":
      case "preservedInline":
        return;
      default: {
        const unsupported: never = content;
        panic(`Unsupported paragraph content: ${JSON.stringify(unsupported)}`);
      }
    }
  };

  const visitBlocks = (nestedBlocks: readonly BlockContent[]): void => {
    for (const block of nestedBlocks) {
      switch (block.type) {
        case "paragraph":
          for (const content of block.content) {
            visitParagraphContent(content);
          }
          break;
        case "table":
          for (const row of block.rows) {
            for (const cell of row.cells) {
              visitBlocks(cell.content);
            }
          }
          break;
        case "blockSdt":
          visitBlocks(block.content);
          break;
        // Opaque markup: nothing inside it for a visitor to reach.
        case "preservedBlock":
          break;
        default: {
          const unsupported: never = block;
          panic(`Unsupported block content: ${JSON.stringify(unsupported)}`);
        }
      }
    }
  };

  visitBlocks(blocks);
  return new Set(
    [...counts].flatMap(([id, count]) =>
      count.starts === 1 &&
      count.ends === 1 &&
      count.firstStart !== undefined &&
      count.firstEnd !== undefined &&
      count.firstStart < count.firstEnd
        ? [id]
        : [],
    ),
  );
};

/**
 * Convert a Document to a ProseMirror document
 *
 * @param document - The Document to convert
 * @param options - Conversion options including style definitions
 */
export function toProseDoc(document: Document, options?: ToProseDocOptions): PMNode {
  const paragraphs = document.package.document.content;
  const nodes: PMNode[] = [];

  // Default to the document's own styles (symmetric with `theme` below) so a
  // run's character-style formatting is always flattened onto direct marks.
  // The save-side `w:rStyle` reconciliation relies on this: without it, an
  // unexpanded run would look like the user stripped the style's formatting
  // (eigenpal/docx-editor#833).
  const styleResolver = createStyleEngine(options?.styles ?? document.package.styles);
  const theme = options?.theme ?? document.package.theme ?? null;
  const nextTextBoxGroupId = createTextBoxGroupIdFactory();
  const nextHyperlinkInstanceIndex = createHyperlinkInstanceIndexAllocator();
  const pairedBookmarkIds = collectPairedBookmarkIds(paragraphs);
  const conversionContext = {
    theme,
    nextTextBoxGroupId,
    nextHyperlinkInstanceIndex,
    pairedBookmarkIds,
    pageBreakRunSourceDescendants: buildPageBreakRunSourceDescendantIndex(paragraphs),
    storyRangedCommentIds: rangedCommentIds(paragraphs),
    openCommentIds: new Set<number>(),
    warnPageBreakProjection: pageBreakProjectionWarn(options?.warn),
  };

  const convertBodyBlocks = (blocks: BlockContent[]): PMNode[] => {
    const out: PMNode[] = [];
    for (const block of blocks) {
      switch (block.type) {
        case "paragraph":
          out.push(
            ...convertParagraphWithTextBoxes(block, styleResolver, {
              textBoxGroupId: nextTextBoxGroupId(),
              context: conversionContext,
            }),
          );
          break;
        case "table":
          out.push(convertTable(block, styleResolver, conversionContext));
          break;
        case "blockSdt":
          out.push(convertBlockSdt(block, convertBodyBlocks));
          break;
        case "preservedBlock":
          out.push(convertPreservedBlock(block));
          break;
        default: {
          const unsupported: never = block;
          panic(`Unsupported block content: ${JSON.stringify(unsupported)}`);
        }
      }
    }
    return out;
  };

  nodes.push(...convertBodyBlocks(paragraphs));

  // Caret-after-final-SDT affordance is provided by `prosemirror-gapcursor`
  // at runtime; we previously injected a trailing empty paragraph here so
  // the caret was not trapped inside an isolating blockSdt, but the
  // synthetic paragraph survived `fromProseDoc` on save and silently
  // appended a `<w:p/>` to the DOCX on every round trip (which adds blank
  // space and shifts pagination in legal templates).

  // Ensure we have at least one paragraph
  if (nodes.length === 0) {
    nodes.push(schema.node("paragraph", {}, []));
  }

  const finalSectionStart =
    document.package.document.sections?.at(-1)?.properties.sectionStart ?? null;
  const adjustLineHeightInTable = document.package.settings?.adjustLineHeightInTable === true;
  const pmDoc = stampNumberedRefFieldBaselines(
    schema.node(
      "doc",
      {
        [PROSE_PARAGRAPH_SOURCE_CONTRACT_ATTR]:
          getDocumentParagraphPropertySourceContract(document) ?? null,
        _finalSectionStart: finalSectionStart,
        _adjustLineHeightInTable: adjustLineHeightInTable,
      },
      nodes,
    ),
  );
  assertValidProseMirrorDocument(
    pmDoc,
    "Document conversion produced an invalid ProseMirror document",
  );
  return pmDoc;
}

/**
 * Carry a block folio does not model into the editor as a zero-width node.
 *
 * The node's place in the document is the whole of its position, so it needs
 * no index and nothing has to keep one honest as the blocks around it change.
 */
function convertPreservedBlock(block: PreservedBlock): PMNode {
  return schema.node("preservedBlock", { xml: block.xml });
}

/**
 * Carry an inline child folio does not model into the editor as an opaque atom.
 *
 * The same atom the run level uses, tagged with the level it came from: the
 * markup is a paragraph child and the save path must not put it back inside a
 * `w:r`. It carries whatever marks surround it, so a capture inside a
 * `w:ins` keeps the insertion and is accepted or rejected with it.
 */
function preservedInlineNode(content: PreservedInline): PMNode {
  return schema.node("preservedXml", {
    xml: content.xml,
    text: content.text,
    level: PRESERVED_XML_LEVELS.inline,
  });
}

/**
 * Convert a `BlockSdt` model node into a `blockSdt` PM node, recursively
 * converting its children with the caller-supplied block converter. Pass
 * `rawPropertiesXml` / `rawEndPropertiesXml` through as attrs so the
 * serializer can replay them verbatim after a save.
 */
function convertBlockSdt(
  blockSdt: BlockSdt,
  convertBlocks: (blocks: BlockContent[]) => PMNode[],
): PMNode {
  const props = blockSdt.properties;
  const attrs: Record<string, unknown> = {
    sdtType: props.sdtType,
    alias: props.alias ?? null,
    tag: props.tag ?? null,
    id: props.id ?? null,
    lock: props.lock ?? null,
    placeholder: props.placeholder ?? null,
    showingPlaceholder: props.showingPlaceholder ?? false,
    dateFormat: props.dateFormat ?? null,
    dateValueISO: props.dateValueISO ?? null,
    listItems: props.listItems ? JSON.stringify(props.listItems) : null,
    dropdownLastValue: props.dropdownLastValue ?? null,
    checked: props.checked ?? null,
    // Mark explicitly when the source content was empty. fromProseDoc reads
    // this on save to drop the synthetic filler below — without an explicit
    // marker we couldn't distinguish source `<w:sdtContent/>` (filler
    // inserted here) from source `<w:sdtContent><w:p/></w:sdtContent>`
    // (a real authored empty paragraph the user wants preserved).
    _originallyEmpty: blockSdt.content.length === 0,
    rawPropertiesXml: props.rawPropertiesXml ?? null,
    rawEndPropertiesXml: props.rawEndPropertiesXml ?? null,
    rawSdtChildrenBeforeContent: props.rawSdtChildrenBeforeContent ?? null,
    rawSdtChildrenAfterContent: props.rawSdtChildrenAfterContent ?? null,
  };
  const children = convertBlocks(blockSdt.content);
  // ProseMirror `blockSdt` requires at least one block child; insert an empty
  // paragraph for a truly empty control rather than producing an invalid node.
  if (children.length === 0) {
    children.push(schema.node("paragraph", {}, []));
  }
  return schema.node("blockSdt", attrs, children);
}

/**
 * Convert a Paragraph to a ProseMirror paragraph node
 *
 * Resolves style-based text formatting and passes it to runs so that
 * paragraph styles (like Heading1) apply their font size, color, etc.
 */
function convertParagraph(
  paragraph: Paragraph,
  styleResolver: StyleEngine | null,
  context: TableConversionContext,
  extraRunFormatting?: TextFormatting,
  tableParagraphOverlay?: TableCellParagraphSpacingOverlay,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
): PMNode {
  const {
    nextHyperlinkInstanceIndex,
    pairedBookmarkIds,
    pageBreakRunSourceDescendants,
    storyRangedCommentIds,
    openCommentIds,
  } = context;
  let pageBreakRunOwnerId = 0;
  const nextPageBreakRunOwnerId = (): number => pageBreakRunOwnerId++;
  const { attrs, effectiveFrame } = paragraphFormattingToAttrs(
    paragraph,
    styleResolver,
    tableParagraphOverlay,
  );
  reportParagraphPageBreakProjection({
    paragraph,
    attrs,
    effectiveFrame,
    sourceDescendants: pageBreakRunSourceDescendants,
    warn: context.warnPageBreakProjection,
  });
  const isTocParagraph = attrs._tableOfContentsLevel !== undefined;
  const inlineNodes: PMNode[] = [];
  let inlineOffset = 0;
  let bookmarksArr: { id: number; name: string }[] | undefined;
  let emptyHyperlinks: NonNullable<ParagraphAttrs["_emptyHyperlinks"]> | undefined;

  // The wrappers the item now being converted sat inside. The loop over the
  // paragraph's content is the only writer, and it sets this before entering
  // the switch, so every node the switch emits gets the item's own stack.
  let wrapperStack: readonly InlineWrapperLayer[] = [];
  const emitInlineNodes = (nodes: PMNode[]): void => {
    if (nodes.length === 0) {
      return;
    }
    const markedNodes = applyCommentMarks(
      withInlineWrapperMark(nodes, wrapperStack),
      openCommentIds,
    );
    inlineNodes.push(...markedNodes);
    for (const node of markedNodes) {
      inlineOffset += node.nodeSize;
    }
  };
  const emitInlineNode = (node: PMNode | null): void => {
    if (!node) {
      return;
    }
    emitInlineNodes([node]);
  };

  const paragraphRunFormatting = resolveRunFormattingWithoutDefaults(
    paragraph.formatting?.runProperties,
    styleResolver,
  );
  // Paragraph-mark-only visual decorations (highlight, shading) paint the
  // paragraph glyph alone. Strip them from the body-run inheritance path.
  let inheritableParagraphRunFormatting: TextFormatting | undefined;
  if (paragraphRunFormatting && !isTocParagraph && paragraph.formatting?.styleId === undefined) {
    inheritableParagraphRunFormatting =
      stripParagraphMarkFormattingForBodyRuns(paragraphRunFormatting);
  }
  const {
    baseFormatting: baseRunFormatting,
    baseToggleCascade: orderedToggleFormatting,
    defaultFormatting: ordinaryBaseWithDefaultCharacter,
    defaultToggleCascade: defaultCharacterStyleCascade,
  } = resolveParagraphBodyRunFormatting({
    styleId: paragraph.formatting?.styleId,
    tableRunFormatting: extraRunFormatting,
    styleResolver,
  });
  const ordinaryDefaultRunFormatting = mergeTextFormatting(
    ordinaryBaseWithDefaultCharacter,
    inheritableParagraphRunFormatting,
  );
  const defaultToggleCascade = cascadeStyleTextFormatting(
    [
      { cascade: defaultCharacterStyleCascade, type: "carried" },
      { formatting: inheritableParagraphRunFormatting, type: "direct" },
    ],
    { ordinaryFormatting: ordinaryDefaultRunFormatting },
  );
  const defaultRunFormatting = defaultToggleCascade.formatting;
  if (extraRunFormatting !== undefined) {
    attrs._tableRunFormatting = extraRunFormatting;
    if (defaultRunFormatting) {
      attrs.defaultTextFormatting = defaultRunFormatting;
    } else {
      delete attrs.defaultTextFormatting;
    }
  }
  const getInheritedRunFormatting = (
    formatting: TextFormatting | undefined,
    fieldType?: string,
  ): ResolvedRunFormatting => {
    const hasCharacterStyle = formatting?.styleId !== undefined;
    const inheritedBaseFormatting = hasCharacterStyle
      ? baseRunFormatting
      : ordinaryBaseWithDefaultCharacter;
    const inheritedToggleCascade = hasCharacterStyle
      ? orderedToggleFormatting
      : defaultCharacterStyleCascade;
    if (fieldType === "TOC") {
      return {
        formatting: hasDirectRunFormatting(formatting)
          ? suppressParagraphMarkFormatting({
              baseFormatting: inheritedBaseFormatting,
              directFormatting: formatting,
              paragraphMarkFormatting: undefined,
            })
          : inheritedBaseFormatting,
        ...(!hasCharacterStyle ? { implicitCharacterStyleApplied: true } : {}),
        toggleCascade: inheritedToggleCascade,
      };
    }
    const hasExplicitRunFormatting =
      hasDirectRunFormatting(formatting) || formatting?.styleId !== undefined;
    if (!hasExplicitRunFormatting) {
      return {
        formatting: defaultRunFormatting,
        implicitCharacterStyleApplied: true,
        toggleCascade: defaultToggleCascade,
      };
    }
    const suppressedFormatting = suppressParagraphMarkFormatting({
      baseFormatting: inheritedBaseFormatting,
      directFormatting: formatting,
      paragraphMarkFormatting: inheritableParagraphRunFormatting,
    });
    const paragraphMarkOverrides =
      getParagraphMarkSuppressionOverrides({
        directFormatting: formatting,
        paragraphMarkFormatting: inheritableParagraphRunFormatting,
        suppressedFormatting,
      }) ?? (inheritableParagraphRunFormatting ? {} : undefined);
    return {
      formatting: suppressedFormatting,
      ...(!hasCharacterStyle ? { implicitCharacterStyleApplied: true } : {}),
      ...(paragraphMarkOverrides ? { paragraphMarkOverrides } : {}),
      toggleCascade: inheritedToggleCascade,
    };
  };
  const getHistoricalRunFormatting: RunFormattingResolver = (formatting) => {
    const hasCharacterStyle = formatting?.styleId !== undefined;
    const inheritedBaseFormatting = hasCharacterStyle
      ? baseRunFormatting
      : ordinaryBaseWithDefaultCharacter;
    const inheritedToggleCascade = hasCharacterStyle
      ? orderedToggleFormatting
      : defaultCharacterStyleCascade;
    return {
      formatting: inheritedBaseFormatting,
      ...(!hasCharacterStyle ? { implicitCharacterStyleApplied: true } : {}),
      toggleCascade: inheritedToggleCascade,
    };
  };
  const trackedRunFormattingResolvers: TrackedRunFormattingResolvers = {
    current: getInheritedRunFormatting,
    historical: getHistoricalRunFormatting,
  };
  const emitTrackedChange = (
    change: Insertion | Deletion | MoveFrom | MoveTo,
    markType: "insertion" | "deletion",
    moveKind: "moveFrom" | "moveTo" | null,
  ): void => {
    emitInlineNodes(
      convertTrackedChange(
        change,
        markType,
        nextHyperlinkInstanceIndex,
        nextPageBreakRunOwnerId,
        trackedRunFormattingResolvers,
        styleResolver,
        moveKind,
        textBoxAnchors,
        wrapperStack,
      ),
    );
  };

  for (const { content, stack } of withInlineWrapperStacks(paragraph.content)) {
    wrapperStack = stack;
    switch (content.type) {
      case "commentRangeStart":
        openCommentIds.add(content.id);
        break;
      case "commentRangeEnd":
        openCommentIds.delete(content.id);
        break;
      case "commentReference":
        // A reference whose own range exists already has its highlight;
        // anchoring it onto neighbouring text would stretch that range to
        // wherever the reference sits. Only a bare reference borrows a
        // neighbour's.
        if (!storyRangedCommentIds.has(content.id)) {
          anchorPointComment(inlineNodes, content.id);
        }
        // The reference is an inline atom of its own, so where it sits among
        // the range ends around it survives the edit. Emitting it through the
        // shared path gives it the marks of the ranges still open over it,
        // which is what tells `fromProseDoc` which ends precede it.
        emitInlineNode(schema.node("commentReference", { commentId: content.id }));
        break;
      case "run":
        emitInlineNodes(
          convertRun(
            content,
            getInheritedRunFormatting(content.formatting),
            nextPageBreakRunOwnerId,
            styleResolver,
            textBoxAnchors,
          ),
        );
        break;
      case "hyperlink": {
        const currentHyperlinkIndex = nextHyperlinkInstanceIndex();
        const linkNodes = convertHyperlink(content, {
          getInheritedRunFormatting,
          styleResolver,
          hyperlinkIndex: currentHyperlinkIndex,
          textBoxAnchors,
          nextPageBreakRunOwnerId,
        });
        if (linkNodes.length === 0) {
          emptyHyperlinks ??= [];
          emptyHyperlinks.push({
            offset: inlineOffset,
            ...(content.href !== undefined ? { href: content.href } : {}),
            ...(content.anchor !== undefined ? { anchor: content.anchor } : {}),
            ...(content.tooltip !== undefined ? { tooltip: content.tooltip } : {}),
            ...(content.rId !== undefined ? { rId: content.rId } : {}),
          });
          break;
        }
        emitInlineNodes(linkNodes);
        break;
      }
      case "simpleField":
      case "complexField":
        emitInlineNode(
          convertField(content, {
            getInheritedRunFormatting,
            styleResolver,
            nextHyperlinkInstanceIndex,
            nextPageBreakRunOwnerId,
            textBoxAnchors,
          }),
        );
        break;
      case "inlineSdt":
        emitInlineNode(
          convertInlineSdt(
            content,
            nextHyperlinkInstanceIndex,
            nextPageBreakRunOwnerId,
            getInheritedRunFormatting,
            trackedRunFormattingResolvers,
            styleResolver,
            textBoxAnchors,
            stack,
          ),
        );
        break;
      case "insertion":
      case "moveTo":
        emitTrackedChange(content, "insertion", content.type === "moveTo" ? "moveTo" : null);
        break;
      case "deletion":
      case "moveFrom":
        emitTrackedChange(content, "deletion", content.type === "moveFrom" ? "moveFrom" : null);
        break;
      case "mathEquation":
        emitInlineNode(convertMathEquation(content));
        break;
      case "bookmarkStart":
        if (pairedBookmarkIds.has(content.id)) {
          emitInlineNode(
            schema.node("bookmarkBoundary", {
              type: "start",
              id: content.id,
              name: content.name,
              colFirst: content.colFirst,
              colLast: content.colLast,
              displacedByCustomXml: content.displacedByCustomXml,
            }),
          );
          break;
        }
        // Legacy structural placement records only the start on a paragraph and
        // uses the paragraph attr to preserve its existing save behavior.
        if (!bookmarksArr) {
          bookmarksArr = [];
        }
        bookmarksArr.push({ id: content.id, name: content.name });
        break;
      case "bookmarkEnd":
        if (pairedBookmarkIds.has(content.id)) {
          emitInlineNode(
            schema.node("bookmarkBoundary", {
              type: "end",
              id: content.id,
              displacedByCustomXml: content.displacedByCustomXml,
            }),
          );
        }
        // An unpaired end has no node: the legacy paragraph attr records the
        // start alone, and the save path rebuilds the end from the source.
        break;
      // Move-range markers are block-level facts the paragraph's own capture
      // replays; the editor carries no inline node for them.
      case "moveFromRangeStart":
      case "moveFromRangeEnd":
      case "moveToRangeStart":
      case "moveToRangeEnd":
        break;
      case "preservedInline":
        emitInlineNode(preservedInlineNode(content));
        break;
      default: {
        const unsupported: never = content;
        panic(`Unsupported paragraph content: ${JSON.stringify(unsupported)}`);
      }
    }
  }

  if (bookmarksArr) {
    attrs.bookmarks = bookmarksArr;
  }
  if (emptyHyperlinks) {
    attrs._emptyHyperlinks = emptyHyperlinks;
  }

  return createProseParagraphWithPropertySource(schema.nodes["paragraph"], paragraph, {
    attrs,
    content: inlineNodes,
  });
}

/**
 * Apply comment marks to PM nodes within a comment range.
 */
function applyCommentMarks(nodes: PMNode[], commentIds: Set<number>): PMNode[] {
  if (commentIds.size === 0) {
    return nodes;
  }
  const commentMarkType = schema.marks["comment"];
  if (!commentMarkType) {
    return nodes;
  }
  const commentMarks = [...commentIds]
    .toSorted((left, right) => left - right)
    .map((commentId) => commentMarkType.create({ commentId }));

  return nodes.map((node) => {
    if (!node.isText && (!node.isInline || !node.type.allowsMarkType(commentMarkType))) {
      return node;
    }
    let marks = node.marks;
    for (const commentMark of commentMarks) {
      marks = commentMark.addToSet(marks);
    }
    return node.mark(marks);
  });
}

function anchorPointComment(nodes: PMNode[], commentId: number): void {
  const commentMark = schema.marks["comment"]?.create({ commentId });
  if (!commentMark) {
    return;
  }

  for (let index = nodes.length - 1; index >= 0; index--) {
    const node = nodes[index];
    if (!node?.isText) {
      continue;
    }
    nodes[index] = node.mark(commentMark.addToSet(node.marks));
    return;
  }
}

/**
 * Convert tracked change (insertion or deletion) content to PM nodes with
 * an insertion/deletion mark applied.
 */
function convertTrackedChange(
  change: Insertion | Deletion | MoveFrom | MoveTo,
  markType: "insertion" | "deletion",
  nextHyperlinkInstanceIndex: HyperlinkInstanceIndexAllocator,
  nextPageBreakRunOwnerId: PageBreakRunOwnerIdAllocator,
  runFormattingResolvers: TrackedRunFormattingResolvers,
  styleResolver?: StyleEngine | null,
  moveKind: "moveFrom" | "moveTo" | null = null,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
  wrappedBy: readonly InlineWrapperLayer[] = [],
): PMNode[] {
  const nodes: PMNode[] = [];
  const getTrackedRunFormatting =
    markType === "deletion" ? runFormattingResolvers.historical : runFormattingResolvers.current;
  // A wrapper the revision holds is lifted here rather than around the
  // revision: lifting it out would take its runs out of the revision with
  // them. What the wrapper holds that a revision may not — a comment or move
  // range boundary — has no place in this projection and the census records it
  // as lost in the editor projection.
  //
  // The accumulation starts empty even when the revision itself sits inside a
  // wrapper: the caller marks what this returns, and `wrappedBy` reaches only
  // the leaves no caller can still see, those inside a content control.
  for (const { content: item, stack } of withInlineWrapperStacks(change.content)) {
    if (!isTrackedChangeWrapperChild(item)) {
      continue;
    }
    const itemNodes: PMNode[] = [];
    if (item.type === "run") {
      itemNodes.push(
        ...convertRun(
          item,
          getTrackedRunFormatting(item.formatting),
          nextPageBreakRunOwnerId,
          styleResolver,
          textBoxAnchors,
        ),
      );
    } else if (item.type === "hyperlink") {
      const currentHyperlinkIndex = nextHyperlinkInstanceIndex();
      itemNodes.push(
        ...convertHyperlink(item, {
          getInheritedRunFormatting: getTrackedRunFormatting,
          styleResolver,
          hyperlinkIndex: currentHyperlinkIndex,
          textBoxAnchors,
          nextPageBreakRunOwnerId,
        }),
      );
    } else if (item.type === "simpleField" || item.type === "complexField") {
      const fieldNode = convertField(item, {
        getInheritedRunFormatting: getTrackedRunFormatting,
        styleResolver,
        nextHyperlinkInstanceIndex,
        nextPageBreakRunOwnerId,
        textBoxAnchors,
      });
      if (fieldNode) {
        itemNodes.push(fieldNode);
      }
    } else if (item.type === "mathEquation") {
      const mathNode = convertMathEquation(item);
      if (mathNode) {
        itemNodes.push(mathNode);
      }
    } else if (
      item.type === "insertion" ||
      item.type === "deletion" ||
      item.type === "moveFrom" ||
      item.type === "moveTo"
    ) {
      const nestedMarkType =
        item.type === "insertion" || item.type === "moveTo" ? "insertion" : "deletion";
      const nestedMoveKind = item.type === "moveFrom" || item.type === "moveTo" ? item.type : null;
      itemNodes.push(
        ...convertTrackedChange(
          item,
          nestedMarkType,
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          runFormattingResolvers,
          styleResolver,
          nestedMoveKind,
          textBoxAnchors,
        ),
      );
    } else if (item.type === "inlineSdt") {
      // The control keeps its node where the author put it, inside the
      // revision. The node is not an inline atom, so it carries no revision
      // mark of its own; the mark that says its text was inserted is the
      // editor carrier this projection still lacks.
      const sdtNode = convertInlineSdt(
        item,
        nextHyperlinkInstanceIndex,
        nextPageBreakRunOwnerId,
        getTrackedRunFormatting,
        runFormattingResolvers,
        styleResolver,
        textBoxAnchors,
        [...wrappedBy, ...stack],
      );
      if (sdtNode) {
        itemNodes.push(sdtNode);
      }
    } else if (item.type === "bookmarkStart") {
      itemNodes.push(
        schema.node("bookmarkBoundary", {
          type: "start",
          id: item.id,
          name: item.name,
          colFirst: item.colFirst,
          colLast: item.colLast,
          displacedByCustomXml: item.displacedByCustomXml,
        }),
      );
    } else if (item.type === "bookmarkEnd") {
      itemNodes.push(
        schema.node("bookmarkBoundary", {
          type: "end",
          id: item.id,
          displacedByCustomXml: item.displacedByCustomXml,
        }),
      );
    } else if (item.type === "preservedInline") {
      itemNodes.push(preservedInlineNode(item));
    } else {
      const unsupported: never = item;
      panic(`Unsupported tracked-run content: ${JSON.stringify(unsupported)}`);
    }
    nodes.push(...withInlineWrapperMark(itemNodes, stack));
  }

  // SAFETY: markType is "insertion" | "deletion", both registered in schema
  const mark = schema.marks[markType]!.create({
    revisionId: change.info.id,
    author: change.info.author,
    date: change.info.date ?? null,
    utcDate: change.info.utcDate?.value ?? null,
    initials: change.info.initials ?? null,
    moveKind,
    ...(markType === "deletion" ? { _historicalFormatting: true } : {}),
  });

  const applyTrackedMark = (node: PMNode): PMNode => {
    // ProseMirror marks cannot nest another mark of the same type. Keep the
    // inner revision intact rather than replacing its identity with the outer
    // wrapper; the surrounding nodes still retain the outer revision.
    if (node.marks.some(({ type }) => type.name === "insertion" || type.name === "deletion")) {
      return node;
    }
    if (trackedRunInlineAtomDisposition(node) === "outside-wrapper") {
      panic(`Inline atom ${JSON.stringify(node.type.name)} cannot occur in a tracked-run wrapper`);
    }
    if (canCarryTrackedRunMark(node)) {
      return node.mark(mark.addToSet(node.marks));
    }
    if (node.type.name === "sdt") {
      const children: PMNode[] = [];
      // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
      node.forEach((child) => children.push(applyTrackedMark(child)));
      return recreateProseNodeWithParagraphPropertySource(node, { content: children });
    }
    return node;
  };

  return nodes.map(applyTrackedMark);
}

/**
 * Convert ParagraphFormatting to ProseMirror paragraph attrs
 *
 * If a styleResolver is provided, resolves style-based formatting and merges
 * with inline formatting. Inline formatting takes precedence.
 */
type ParagraphFormattingProjection = {
  attrs: ParagraphAttrs;
  effectiveFrame: ParagraphFormatting["frame"];
};

function paragraphFormattingToAttrs(
  paragraph: Paragraph,
  styleResolver: StyleEngine | null,
  tableParagraphOverlay?: TableCellParagraphSpacingOverlay,
): ParagraphFormattingProjection {
  const formatting = paragraph.formatting;
  const styleId = formatting?.styleId;
  const styleName = styleId ? styleResolver?.getStyle(styleId)?.name : undefined;
  const tableOfContentsLevel = tableOfContentsStyleLevel({
    styleId,
    ...(styleName ? { styleName } : {}),
  });

  // Start with base attrs — only include defined values
  const attrs: ParagraphAttrs = {};

  if (paragraph.paraId) {
    attrs.paraId = paragraph.paraId;
  }
  if (paragraph.textId) {
    attrs.textId = paragraph.textId;
  }
  if (paragraph.reviewCarrier) {
    attrs.reviewCarrier = paragraph.reviewCarrier;
  }
  if (styleId) {
    attrs.styleId = styleId;
  }
  if (tableOfContentsLevel !== undefined) {
    attrs._tableOfContentsLevel = tableOfContentsLevel;
  }
  if (formatting?.numPr) {
    attrs.numPr = formatting.numPr;
  }
  if (formatting?.numPrFromStyle) {
    attrs.numPrFromStyle = formatting.numPrFromStyle;
  }
  // List rendering info from parsed numbering definitions
  if (paragraph.listRendering) {
    Object.assign(attrs, listRenderingAttrPatch(paragraph.listRendering));
  }
  // Store original inline formatting for lossless serialization round-trip
  if (formatting) {
    attrs._originalFormatting = formatting;
  }
  // Carry `w:pPrChange` (paragraph-property-change tracking) opaquely
  // through ProseMirror. Without this, every edit strips the entries
  // off the paragraph because nothing in PM's schema represents them.
  // Shallow-clone the array so the editor state owns its own
  // reference — mutating the Folio document later must not poke
  // through into PM attrs.
  if (paragraph.propertyChanges && paragraph.propertyChanges.length > 0) {
    attrs._propertyChanges = [...paragraph.propertyChanges];
  }
  // The attributes the authored `w:p` carried and the model has no field for.
  // The array instance is what `fromProseDoc` recognises on the way back: PM
  // copies attrs to both halves of a split, so the second half is holding this
  // very array and is told apart from a paragraph that authored its own.
  if (paragraph.preservedAttributes && paragraph.preservedAttributes.length > 0) {
    attrs._preservedAttributes = paragraph.preservedAttributes;
  }
  if (paragraph.pPrMark) {
    attrs.pPrMark = paragraph.pPrMark;
  }

  // Helper: assign a value only when defined
  const set = <K extends keyof ParagraphAttrs>(
    key: K,
    val: ParagraphAttrs[K] | undefined,
  ): void => {
    if (val !== undefined) {
      attrs[key] = val;
    }
  };

  // If we have a style resolver, resolve the style and get base properties.
  // Cell paragraphs (`tableParagraphOverlay` set) layer the enclosing table
  // style's modeled paragraph fields in between docDefaults and this
  // paragraph's own style chain — see resolveParagraphStyleInTable.
  let stylePpr: Paragraph["formatting"] | undefined;
  if (styleResolver) {
    const resolved = styleResolver.resolveParagraphStyleInTable(styleId, tableParagraphOverlay);
    stylePpr = resolved.paragraphFormatting;
    // What the paragraph would render as if it stated nothing of its own,
    // narrowed to the fields a save could otherwise materialise (see
    // ParagraphAttrs._resolvedFormatting).
    const resolvedFormatting = styleResolvedParagraphFormatting(stylePpr);
    if (resolvedFormatting) {
      attrs._resolvedFormatting = resolvedFormatting;
    }

    // Apply style-based values as defaults (inline overrides)
    set("alignment", formatting?.alignment ?? stylePpr?.alignment);
    set("alignmentFromStyle", stylePpr?.alignment);
    set("spaceBefore", formatting?.spaceBefore ?? stylePpr?.spaceBefore);
    set("spaceAfter", formatting?.spaceAfter ?? stylePpr?.spaceAfter);
    set("lineSpacing", formatting?.lineSpacing ?? stylePpr?.lineSpacing);
    set("lineSpacingRule", formatting?.lineSpacingRule ?? stylePpr?.lineSpacingRule);
    set("lineSpacingExplicit", lineSpacingProvenanceFromSpacing(formatting));
    set("snapToGrid", formatting?.snapToGrid ?? stylePpr?.snapToGrid);
    set("spacingExplicit", formatting?.spacingExplicit);
    const paragraphStyle = styleId
      ? (styleResolver.getStyle(styleId) ?? styleResolver.getDefaultParagraphStyle())
      : styleResolver.getDefaultParagraphStyle();
    const docDefaultSpacing = styleResolver.getDocDefaults()?.pPr;
    // This existing provenance attribute covers every resolved style layer:
    // default, named paragraph, and enclosing table styles. The direct
    // `formatting` object still wins per field.
    const spacingFromStyle: NonNullable<ParagraphAttrs["spacingFromImplicitDefaultStyle"]> = {};
    if (formatting?.spaceBefore === undefined && stylePpr?.spaceBefore !== undefined) {
      spacingFromStyle.before = true;
    }
    if (formatting?.spaceAfter === undefined && stylePpr?.spaceAfter !== undefined) {
      spacingFromStyle.after = true;
    }
    if (spacingFromStyle.before || spacingFromStyle.after) {
      attrs.spacingFromImplicitDefaultStyle = spacingFromStyle;
    }
    const spacingFromDocDefaults: NonNullable<ParagraphAttrs["spacingFromDocDefaults"]> = {};
    if (
      formatting?.spaceBefore === undefined &&
      tableParagraphOverlay?.spaceBefore === undefined &&
      paragraphStyle?.pPr?.spaceBefore === undefined &&
      docDefaultSpacing?.spaceBefore !== undefined
    ) {
      spacingFromDocDefaults.before = true;
    }
    if (
      formatting?.spaceAfter === undefined &&
      tableParagraphOverlay?.spaceAfter === undefined &&
      paragraphStyle?.pPr?.spaceAfter === undefined &&
      docDefaultSpacing?.spaceAfter !== undefined
    ) {
      spacingFromDocDefaults.after = true;
    }
    if (spacingFromDocDefaults.before || spacingFromDocDefaults.after) {
      attrs.spacingFromDocDefaults = spacingFromDocDefaults;
    }
    // When the paragraph explicitly removes the style's numbering (direct
    // numId=0 under a numbered style), the reference layout also drops the
    // style's marker-positioning indents. The paragraph keeps only the indents
    // it states itself (#765: a direct left=357 renders indented instead of
    // hanging the first line back to the margin). Outside that case w:ind
    // merges per attribute: a direct left-only indent keeps the style's
    // firstLine.
    const numberingRemoved =
      formatting?.numPr?.numId === NO_NUMBERING_NUM_ID &&
      isNumberingReference(stylePpr?.numPr?.numId);
    const numberingStyleIndent = numberingRemoved ? undefined : stylePpr;
    const effectiveIndent = mergeParagraphFormatting(numberingStyleIndent, formatting);
    set("indentLeft", effectiveIndent?.indentLeft);
    set("indentRight", formatting?.indentRight ?? stylePpr?.indentRight);
    set("indentFirstLine", effectiveIndent?.indentFirstLine);
    set("hangingIndent", effectiveIndent?.hangingIndent);
    set("borders", formatting?.borders ?? stylePpr?.borders);
    set("shading", formatting?.shading ?? stylePpr?.shading);
    set("tabs", mergeParagraphTabStops(stylePpr?.tabs, formatting?.tabs));
    set("kinsoku", formatting?.kinsoku ?? stylePpr?.kinsoku);
    set("overflowPunctuation", formatting?.overflowPunctuation ?? stylePpr?.overflowPunctuation);
    set("suppressAutoHyphens", formatting?.suppressAutoHyphens ?? stylePpr?.suppressAutoHyphens);

    // Page break control
    set("pageBreakBefore", formatting?.pageBreakBefore ?? stylePpr?.pageBreakBefore);
    set("keepNext", formatting?.keepNext ?? stylePpr?.keepNext);
    set("keepLines", formatting?.keepLines ?? stylePpr?.keepLines);
    set("widowControl", formatting?.widowControl ?? stylePpr?.widowControl);
    set("contextualSpacing", formatting?.contextualSpacing ?? stylePpr?.contextualSpacing);
    // Run-in heading (`<w:specVanish/>` on the paragraph mark) — see
    // ParagraphAttrs.runInWithNext.
    set("runInWithNext", formatting?.runInWithNext ?? stylePpr?.runInWithNext);

    // Outline level (for TOC)
    set("outlineLevel", formatting?.outlineLevel ?? stylePpr?.outlineLevel);

    // Text direction — a direct or style-sourced `w:bidi` is an authoritative
    // manual decision (auto-detection must not override it).
    set("direction", directionFromBidi(formatting?.bidi ?? stylePpr?.bidi));

    set(
      "defaultTextFormatting",
      resolveParagraphDefaultTextFormatting(styleId, formatting, styleResolver, {
        includeParagraphMarkRunProperties:
          tableOfContentsLevel === undefined &&
          (styleId === undefined || paragraph.content.length === 0),
      }),
    );

    // A direct numPr may carry only ilvl while the style supplies numId.
    // Merge the two fields so the effective list keeps the style's numbering
    // identity. A direct numId (including 0) is authoritative.
    if (
      stylePpr?.numPr &&
      formatting?.numPr?.numId === undefined &&
      isNumberingReference(stylePpr.numPr.numId)
    ) {
      attrs.numPr = { ...stylePpr.numPr, ...formatting?.numPr };
      attrs.numPrFromStyle = stylePpr.numPr;
    }
  } else {
    // No style resolver - use inline formatting only
    set("alignment", formatting?.alignment);
    set("spaceBefore", formatting?.spaceBefore);
    set("spaceAfter", formatting?.spaceAfter);
    set("lineSpacing", formatting?.lineSpacing);
    set("lineSpacingRule", formatting?.lineSpacingRule);
    set("lineSpacingExplicit", lineSpacingProvenanceFromSpacing(formatting));
    set("snapToGrid", formatting?.snapToGrid);
    set("spacingExplicit", formatting?.spacingExplicit);
    set("indentLeft", formatting?.indentLeft);
    set("indentRight", formatting?.indentRight);
    set("indentFirstLine", formatting?.indentFirstLine);
    set("hangingIndent", formatting?.hangingIndent);
    set("borders", formatting?.borders);
    set("shading", formatting?.shading);
    set("tabs", formatting?.tabs);
    set("kinsoku", formatting?.kinsoku);
    set("overflowPunctuation", formatting?.overflowPunctuation);
    set("suppressAutoHyphens", formatting?.suppressAutoHyphens);

    // Page break control
    set("pageBreakBefore", formatting?.pageBreakBefore);
    set("keepNext", formatting?.keepNext);
    set("keepLines", formatting?.keepLines);
    set("widowControl", formatting?.widowControl);
    set("runInWithNext", formatting?.runInWithNext);

    // Outline level
    set("outlineLevel", formatting?.outlineLevel);

    // Text direction — an imported `w:bidi` is an authoritative manual decision.
    set("direction", directionFromBidi(formatting?.bidi));

    // Default run properties (pPr/rPr)
    set(
      "defaultTextFormatting",
      stripParagraphMarkOnlyFormatting(
        resolveTextFormatting(formatting?.runProperties, styleResolver) ?? {},
      ),
    );
  }

  // Section break type and full section properties for layout + round-trip
  if (paragraph.sectionProperties) {
    attrs._sectionProperties = paragraph.sectionProperties;
    const st = paragraph.sectionProperties.sectionStart;
    if (st === "nextPage" || st === "continuous" || st === "oddPage" || st === "evenPage") {
      attrs.sectionBreakType = st;
    }
  }
  if (paragraph.renderedPageBreakBefore) {
    attrs.renderedPageBreakBefore = true;
  }

  const beforeAutospacing = formatting?.beforeAutospacing ?? stylePpr?.beforeAutospacing;
  const afterAutospacing = formatting?.afterAutospacing ?? stylePpr?.afterAutospacing;
  if (beforeAutospacing || afterAutospacing) {
    const base: NonNullable<ParagraphAttrs["_autospacingBase"]> = {};
    if (beforeAutospacing) {
      setAutospacingBaseValue(base, "before", attrs.spaceBefore);
    }
    if (afterAutospacing) {
      setAutospacingBaseValue(base, "after", attrs.spaceAfter);
    }
    attrs._autospacingBase = base;
  }

  return {
    attrs,
    effectiveFrame:
      formatting?.frame === undefined
        ? stylePpr?.frame
        : { ...stylePpr?.frame, ...formatting.frame },
  };
}

// ============================================================================
// TABLE CONVERSION
// ============================================================================

/**
 * A table style's (or one of its `w:tblStylePr` conditional regions')
 * contribution to cell formatting: cell properties, run defaults, and the
 * modeled paragraph overlay described on {@link TableCellParagraphSpacingOverlay}.
 */
type TableConditionalStyle = {
  tcPr?: TableCellFormatting;
  rPr?: TextFormatting;
  pPr?: TableCellParagraphSpacingOverlay;
};

/**
 * Pick the modeled paragraph fields out of a table style's (or conditional
 * region's) `w:pPr` for use as the cell-paragraph cascade overlay.
 */
function extractTableParagraphOverlay(
  pPr: ParagraphFormatting | undefined,
): TableCellParagraphSpacingOverlay | undefined {
  if (!pPr) {
    return undefined;
  }
  const overlay: TableCellParagraphSpacingOverlay = {};
  if (pPr.spaceBefore !== undefined) {
    overlay.spaceBefore = pPr.spaceBefore;
  }
  if (pPr.spaceAfter !== undefined) {
    overlay.spaceAfter = pPr.spaceAfter;
  }
  if (pPr.lineSpacing !== undefined) {
    overlay.lineSpacing = pPr.lineSpacing;
  }
  if (pPr.lineSpacingRule !== undefined) {
    overlay.lineSpacingRule = pPr.lineSpacingRule;
  }
  if (pPr.contextualSpacing !== undefined) {
    overlay.contextualSpacing = pPr.contextualSpacing;
  }
  if (pPr.frame !== undefined) {
    overlay.frame = pPr.frame;
  }
  return Object.keys(overlay).length > 0 ? overlay : undefined;
}

/**
 * Resolve table style conditional formatting
 */
function resolveTableStyleConditional(
  styleResolver: StyleEngine | null,
  tableStyleId: string | undefined,
  conditionType: string,
): TableConditionalStyle | undefined {
  if (!styleResolver || !tableStyleId) {
    return undefined;
  }

  const style = styleResolver.getStyle(tableStyleId);
  if (!style?.tblStylePr) {
    return undefined;
  }

  const conditional = style.tblStylePr.find((p) => p.type === conditionType);
  if (!conditional) {
    return undefined;
  }

  const runPropsFromPpr = conditional.pPr?.runProperties
    ? resolveRunFormattingWithoutDefaults(conditional.pPr.runProperties, styleResolver)
    : undefined;
  const resolvedRpr = conditional.rPr
    ? resolveRunFormattingWithoutDefaults(conditional.rPr, styleResolver)
    : undefined;
  const mergedRunProps = mergeTextFormatting(runPropsFromPpr, resolvedRpr);
  const paragraphOverlay = extractTableParagraphOverlay(conditional.pPr);

  const result: TableConditionalStyle = {};
  if (conditional.tcPr) {
    result.tcPr = conditional.tcPr;
  }
  if (mergedRunProps) {
    result.rPr = mergedRunProps;
  }
  if (paragraphOverlay) {
    result.pPr = paragraphOverlay;
  }
  return result;
}

function resolveTableBaseStyle(
  styleResolver: StyleEngine | null,
  tableStyleId: string | undefined,
): TableConditionalStyle | undefined {
  if (!styleResolver || !tableStyleId) {
    return undefined;
  }

  const style = styleResolver.getStyle(tableStyleId);
  if (!style) {
    return undefined;
  }

  const runPropsFromPpr = style.pPr?.runProperties
    ? resolveRunFormattingWithoutDefaults(style.pPr.runProperties, styleResolver)
    : undefined;
  const resolvedRpr = style.rPr
    ? resolveRunFormattingWithoutDefaults(style.rPr, styleResolver)
    : undefined;
  const mergedRunProps = mergeTextFormatting(runPropsFromPpr, resolvedRpr);
  const paragraphOverlay = extractTableParagraphOverlay(style.pPr);

  const result: TableConditionalStyle = {};
  if (style.tcPr) {
    result.tcPr = style.tcPr;
  }
  if (mergedRunProps) {
    result.rPr = mergedRunProps;
  }
  if (paragraphOverlay) {
    result.pPr = paragraphOverlay;
  }
  return result.tcPr || result.rPr || result.pPr ? result : undefined;
}

function mergeConditionalStyles(
  base?: TableConditionalStyle,
  override?: TableConditionalStyle,
): TableConditionalStyle | undefined {
  if (!base && !override) {
    return undefined;
  }
  if (!base) {
    return override;
  }
  if (!override) {
    return base;
  }

  const merged: TableConditionalStyle = {};

  const baseTcPr = base.tcPr;
  const overrideTcPr = override.tcPr;
  if (baseTcPr || overrideTcPr) {
    const tcPr: TableCellFormatting = {
      ...baseTcPr,
      ...overrideTcPr,
    };

    if (baseTcPr?.borders || overrideTcPr?.borders) {
      tcPr.borders = {
        ...baseTcPr?.borders,
        ...overrideTcPr?.borders,
      };
    }

    if (baseTcPr?.shading || overrideTcPr?.shading) {
      tcPr.shading = {
        ...baseTcPr?.shading,
        ...overrideTcPr?.shading,
      };
    }

    if (baseTcPr?.margins || overrideTcPr?.margins) {
      tcPr.margins = {
        ...baseTcPr?.margins,
        ...overrideTcPr?.margins,
      };
    }

    merged.tcPr = tcPr;
  }

  const mergedRPr = mergeTextFormatting(base.rPr, override.rPr);
  if (mergedRPr) {
    merged.rPr = mergedRPr;
  }

  // `override` (a more specific conditional region, e.g. firstRow) wins per
  // field over `base` (e.g. the table's wholeTable region or base style),
  // matching the tcPr/rPr merges above — see extractTableParagraphOverlay.
  const mergedPPr = mergeParagraphFormatting(base.pPr, override.pPr);
  if (mergedPPr) {
    merged.pPr = mergedPPr;
  }

  return merged;
}

function resolveTextFormatting(
  formatting: TextFormatting | undefined,
  styleResolver: StyleEngine | null,
): TextFormatting | undefined {
  if (!formatting) {
    return styleResolver?.resolveRunStyle(null);
  }
  if (!styleResolver) {
    return formatting;
  }

  const styleFormatting = styleResolver.resolveRunStyle(formatting.styleId);
  return mergeTextFormatting(styleFormatting, formatting);
}

/**
 * Resolve an embedded character-style reference without importing
 * `docDefaults`. The caller already has the paragraph cascade, including
 * document defaults, and will layer these own properties over it.
 */
type ParagraphDefaultFormattingResolver = Pick<
  StyleEngine,
  | "getStyle"
  | "getDocDefaults"
  | "getDefaultParagraphStyle"
  | "getDefaultCharacterStyle"
  | "getRunStyleOwnProperties"
>;

function resolveRunFormattingWithoutDefaults(
  formatting: TextFormatting | undefined,
  styleResolver: ParagraphDefaultFormattingResolver | null,
): TextFormatting | undefined {
  if (!formatting || !styleResolver) {
    return formatting;
  }

  const characterStyleFormatting = formatting.styleId
    ? styleResolver.getRunStyleOwnProperties(formatting.styleId)
    : undefined;
  return cascadeStyleTextFormatting([
    { formatting: characterStyleFormatting, type: "style" },
    { formatting, type: "direct" },
  ]).formatting;
}

/** @internal Recompute a paragraph's inherited run defaults from authored package state. */
export function resolveParagraphDefaultTextFormatting(
  styleId: string | undefined,
  formatting: Paragraph["formatting"] | undefined,
  styleResolver: ParagraphDefaultFormattingResolver,
  options: { includeParagraphMarkRunProperties?: boolean } = {},
): TextFormatting | undefined {
  const style = styleId
    ? (styleResolver.getStyle(styleId) ?? styleResolver.getDefaultParagraphStyle())
    : styleResolver.getDefaultParagraphStyle();
  const paragraphStyleRpr = style?.type === "paragraph" ? style.rPr : undefined;
  // The pPr/rPr block describes the paragraph mark only — see the comment on
  // `stripParagraphMarkOnlyFormatting`. We must NOT route this through
  // `resolveTextFormatting` here, because that folds docDefaults back into
  // the run properties and then overwrites the paragraph style's font
  // (e.g. FootnoteText's Times New Roman) with the docDefault Calibri when
  // merged into the cascade below.
  const rawParagraphMarkRpr =
    options.includeParagraphMarkRunProperties === false ? undefined : formatting?.runProperties;
  const paragraphRunProperties = rawParagraphMarkRpr
    ? stripParagraphMarkOnlyFormatting(
        resolveRunFormattingWithoutDefaults(rawParagraphMarkRpr, styleResolver) ?? {},
      )
    : undefined;

  const orderedBodyToggleFormatting = cascadeStyleTextFormatting(
    [
      { formatting: styleResolver.getDocDefaults()?.rPr, type: "defaults" },
      { formatting: paragraphStyleRpr, type: "style" },
      { formatting: styleResolver.getDefaultCharacterStyle()?.rPr, type: "style" },
    ],
    {
      ordinaryFormatting: mergeTextFormatting(
        mergeTextFormatting(
          styleResolver.getDocDefaults()?.rPr,
          styleResolver.getDefaultCharacterStyle()?.rPr,
        ),
        paragraphStyleRpr,
      ),
    },
  );
  const bodyRunDefaults = orderedBodyToggleFormatting.formatting;
  return cascadeStyleTextFormatting(
    [
      { cascade: orderedBodyToggleFormatting, type: "carried" },
      { formatting: paragraphRunProperties, type: "direct" },
    ],
    {
      ordinaryFormatting: mergeTextFormatting(bodyRunDefaults, paragraphRunProperties),
    },
  ).formatting;
}

/**
 * Convert a Table to a ProseMirror table node
 *
 * Handles column widths from w:tblGrid - if cell widths aren't specified,
 * we use the grid column widths to set cell widths. This ensures tables
 * preserve their layout when opened from DOCX files.
 */
/**
 * Calculate rowSpan values from vMerge attributes.
 * OOXML uses vMerge="restart" to start a vertical merge and vMerge="continue" for cells that should be merged.
 * This function converts that to rowSpan values and marks which cells should be skipped.
 */
type RowSpanInfo = {
  rowSpan: number;
  skip: boolean;
  preserveVMergeRestart?: boolean;
  continuationCells?: TableCell[];
};

function calculateRowSpans(table: Table): Map<string, RowSpanInfo> {
  const result = new Map<string, RowSpanInfo>();
  const numRows = table.rows.length;

  // Track active vertical merges per column (stores the row index where merge started)
  const activeMerges = new Map<number, number>();

  // Process each row
  for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
    // SAFETY: rowIndex < numRows <= table.rows.length
    const row = table.rows[rowIndex]!;
    if (row.cells.length === 0) {
      clearActiveVerticalMerges(activeMerges, result);
      continue;
    }
    let colIndex = row.formatting?.gridBefore ?? 0;
    const rowCells = row.cells.map((cell) => {
      const colspan = cell.formatting?.gridSpan ?? 1;
      const vMerge = cell.formatting?.vMerge;
      const isMergeContinuation = isCellMergeContinuation(cell);
      const startRow = isMergeContinuation ? activeMerges.get(colIndex) : undefined;
      const info = {
        cell,
        colIndex,
        colspan,
        vMerge,
        isMergeContinuation,
        startRow,
        hasMeaningfulContent: tableCellHasMeaningfulContent(cell),
        shouldSkip: isMergeContinuation && startRow !== undefined,
      };
      colIndex += colspan;
      return info;
    });
    const rowWouldBeEmpty = rowCells.length > 0 && rowCells.every((cell) => cell.shouldSkip);

    for (const cellInfo of rowCells) {
      const {
        colIndex: cellColIndex,
        vMerge,
        isMergeContinuation,
        startRow,
        hasMeaningfulContent,
      } = cellInfo;
      const key = `${rowIndex}-${cellColIndex}`;

      if (vMerge === "restart") {
        // Start of a new vertical merge. A restart directly under another one
        // ends that one, so close it before this row takes the column over.
        closeVerticalMerge(activeMerges, result, cellColIndex);
        activeMerges.set(cellColIndex, rowIndex);
        result.set(key, { rowSpan: 1, skip: false });
      } else if (isMergeContinuation) {
        // Continuation of a merge - only skip it when the parsed grid has a
        // matching restart in this exact column and the continuation is only a
        // structural placeholder. Real DOCX tables can be ragged, and some
        // continuation cells contain drawings or other payload that must not be
        // merged away.
        if (startRow === undefined || rowWouldBeEmpty || hasMeaningfulContent) {
          result.set(key, { rowSpan: 1, skip: false });
          if ((rowWouldBeEmpty || hasMeaningfulContent) && startRow !== undefined) {
            const restartCell = result.get(`${startRow}-${cellColIndex}`);
            if (restartCell) {
              restartCell.preserveVMergeRestart = true;
            }
            activeMerges.delete(cellColIndex);
          }
          continue;
        }

        // Increment rowSpan of the starting cell
        const startKey = `${startRow}-${cellColIndex}`;
        const startCell = result.get(startKey);
        if (startCell) {
          startCell.rowSpan++;
          startCell.continuationCells ??= [];
          startCell.continuationCells.push(cellInfo.cell);
        }
        result.set(key, { rowSpan: 1, skip: true });
      } else {
        // No vMerge - clear any active merge for this column
        closeVerticalMerge(activeMerges, result, cellColIndex);
        result.set(key, { rowSpan: 1, skip: false });
      }
    }
  }

  // A merge still open when the table ends never gained a continuation, so
  // nothing but the flag records that the cell said `w:vMerge="restart"`.
  clearActiveVerticalMerges(activeMerges, result);

  return result;
}

/**
 * End the vertical merge active in one column.
 *
 * A `w:vMerge="restart"` is carried through the editor by its cell's rowspan,
 * which only exists once a continuation joins it. A merge that closes with a
 * rowspan of 1 — a restart with no continuation, one interrupted by a plain
 * cell, one whose continuation is in a row a revision removed — has nothing
 * but this flag to say the cell was a merge origin, and dropping it changes
 * the table's visible structure.
 */
function closeVerticalMerge(
  activeMerges: Map<number, number>,
  result: Map<string, RowSpanInfo>,
  colIndex: number,
): void {
  const startRow = activeMerges.get(colIndex);
  if (startRow === undefined) {
    return;
  }
  const restartCell = result.get(`${startRow}-${colIndex}`);
  // A merge that absorbed a continuation is already spelled by the rowspan,
  // and flagging it would resurrect the restart when a revision later splits
  // the cell back apart. Only a merge closing at a rowspan of one needs it.
  if (restartCell && restartCell.rowSpan === 1) {
    restartCell.preserveVMergeRestart = true;
  }
  activeMerges.delete(colIndex);
}

function clearActiveVerticalMerges(
  activeMerges: Map<number, number>,
  result: Map<string, RowSpanInfo>,
): void {
  for (const colIndex of [...activeMerges.keys()]) {
    closeVerticalMerge(activeMerges, result, colIndex);
  }
}

function tableCellHasMeaningfulContent(cell: TableCell): boolean {
  return cell.content.some(blockHasMeaningfulContent);
}

function blockHasMeaningfulContent(block: TableCellBlock): boolean {
  // Markup the cell carries is content, even though folio cannot read it.
  if (block.type === "preservedBlock") {
    return true;
  }
  if (block.type === "table") {
    return block.rows.some((row) => row.cells.some((cell) => tableCellHasMeaningfulContent(cell)));
  }

  return block.content.some(paragraphContentHasMeaningfulContent);
}

function paragraphContentHasMeaningfulContent(content: Paragraph["content"][number]): boolean {
  if (content.type === "run") {
    return content.content.length > 0;
  }
  if (content.type === "hyperlink") {
    return content.children.some(paragraphContentHasMeaningfulContent);
  }
  if (
    content.type === "insertion" ||
    content.type === "deletion" ||
    content.type === "moveFrom" ||
    content.type === "moveTo"
  ) {
    return content.content.some(paragraphContentHasMeaningfulContent);
  }
  return true;
}

type TableConversionContext = {
  theme: Theme | null | undefined;
  nextTextBoxGroupId: () => string;
  nextHyperlinkInstanceIndex: HyperlinkInstanceIndexAllocator;
  pairedBookmarkIds: ReadonlySet<number>;
  pageBreakRunSourceDescendants: PageBreakRunSourceDescendantIndex;
  /** Comments the story opens a range for, for the point-comment question. */
  storyRangedCommentIds: ReadonlySet<number>;
  warnPageBreakProjection: PageBreakProjectionWarn;
  /**
   * Comment ranges open at the walk's current position. A range is a story
   * fact, not a paragraph one: it opens in one paragraph and closes in another,
   * and every inline node between the two carries its mark. The walk visits the
   * story in document order, so the set the block walker carries is exactly the
   * set of ranges covering the node it is converting — the same way
   * `pairedBookmarkIds` answers for a boundary whose partner is elsewhere.
   */
  openCommentIds: Set<number>;
};

function convertTable(
  table: Table,
  styleResolver: StyleEngine | null,
  context: TableConversionContext,
): PMNode {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      if (
        isCellMergeContinuation(cell) &&
        context.pageBreakRunSourceDescendants.containsPageBreakRun(cell.content)
      ) {
        context.warnPageBreakProjection(
          "A page break inside a vertically merged table cell paginates with the merged row",
        );
        continue;
      }
      reportSourceContainerPageBreakRun(
        cell.content,
        "table-cell",
        context.pageBreakRunSourceDescendants,
        context.warnPageBreakProjection,
      );
    }
  }

  // Calculate rowSpan values from vMerge
  const rowSpanMap = calculateRowSpans(table);

  // Get column widths from table grid
  const columnWidths = table.columnWidths;

  // Calculate total width from columnWidths if available (for percentage calculation)
  const totalWidth = columnWidths?.reduce((sum, w) => sum + w, 0) ?? 0;

  // Get the table style's conditional formatting
  const tableStyleId = table.formatting?.styleId;
  // `attrs.look` keeps what the author wrote; every read of a flag goes through
  // the resolver, so a table that states its look only as `w:val` — anything
  // older than the attribute form — gets its banding and its header row.
  const look = resolveTableLook(table.formatting?.look);

  // Resolve table borders through inline style, table style, then default table style.
  const tableStyle = tableStyleId ? styleResolver?.getStyle(tableStyleId) : undefined;
  const defaultTableStyle = styleResolver?.getDefaultTableStyle();
  const fallbackTableStyle = tableStyleId ? undefined : defaultTableStyle;
  const conditionalTableStyleId = tableStyle?.styleId ?? fallbackTableStyle?.styleId;
  const resolvedTableBorders =
    table.formatting?.borders ?? tableStyle?.tblPr?.borders ?? fallbackTableStyle?.tblPr?.borders;
  const fallbackTableIndent = fallbackTableStyle?.tblPr?.indent;
  // TableNormal's zero is Word's unindented default: keep it absent so a save
  // does not turn an inherited default into direct formatting. Direct and
  // explicit-style zero remain authored.
  const resolvedTableIndent =
    table.formatting?.indent ??
    tableStyle?.tblPr?.indent ??
    (fallbackTableIndent?.value === 0 ? undefined : fallbackTableIndent);
  const resolvedTableJustification =
    tableStyle?.tblPr?.justification ?? fallbackTableStyle?.tblPr?.justification;
  const resolvedTableBidi =
    table.formatting?.bidi ?? tableStyle?.tblPr?.bidi ?? fallbackTableStyle?.tblPr?.bidi;

  // Resolve default cell margins through the same table-style cascade.
  const tableCellMargins =
    table.formatting?.cellMargins ??
    tableStyle?.tblPr?.cellMargins ??
    fallbackTableStyle?.tblPr?.cellMargins;
  let cellMarginsAttr: TableCellMarginsAttrs | undefined;
  if (tableCellMargins) {
    const m: TableCellMarginsAttrs = {};
    if (tableCellMargins.top?.value !== undefined) {
      m.top = tableCellMargins.top.value;
    }
    if (tableCellMargins.bottom?.value !== undefined) {
      m.bottom = tableCellMargins.bottom.value;
    }
    if (tableCellMargins.left?.value !== undefined) {
      m.left = tableCellMargins.left.value;
    }
    if (tableCellMargins.right?.value !== undefined) {
      m.right = tableCellMargins.right.value;
    }
    cellMarginsAttr = m;
  }

  const attrs: TableAttrs = {};
  if (table.formatting?.styleId) {
    attrs.styleId = table.formatting.styleId;
  }
  if (table.formatting?.width?.value !== undefined) {
    attrs.width = table.formatting.width.value;
  }
  if (table.formatting?.width?.type) {
    attrs.widthType = table.formatting.width.type;
  }
  if (table.formatting?.justification) {
    attrs.justification = table.formatting.justification;
  }
  if (columnWidths) {
    attrs.columnWidths = columnWidths;
  }
  if (table.formatting?.floating) {
    attrs.floating = table.formatting.floating;
  }
  if (cellMarginsAttr) {
    attrs.cellMargins = cellMarginsAttr;
    attrs._resolvedCellMargins = cellMarginsAttr;
  }
  if (table.formatting?.look) {
    attrs.look = table.formatting.look;
  }
  if (table.formatting?.borders) {
    attrs.borders = table.formatting.borders;
  }
  if (resolvedTableIndent) {
    attrs._resolvedIndent = resolvedTableIndent;
  }
  if (resolvedTableJustification) {
    attrs._resolvedJustification = resolvedTableJustification;
  }
  if (resolvedTableBidi !== undefined) {
    attrs._resolvedBidi = resolvedTableBidi;
  }
  // Always set on import, even when the table had no `w:tblPr`: this attr is
  // also what tells `convertPMTable` the table came from a document, and a
  // table that arrived without table borders must not acquire them from a
  // bordered cell on the way back out.
  attrs._originalFormatting = table.formatting ?? {};
  // Carry `w:tblPrChange` opaquely through PM (same rationale as the
  // paragraph `_propertyChanges` attr) so edits don't strip the tracked
  // property-change history and accept/reject can resolve it.
  if (table.propertyChanges && table.propertyChanges.length > 0) {
    attrs.tblPrChange = [...table.propertyChanges];
  }

  const conditionalStyles: {
    wholeTable?: TableConditionalStyle;
    firstRow?: TableConditionalStyle;
    lastRow?: TableConditionalStyle;
    firstCol?: TableConditionalStyle;
    lastCol?: TableConditionalStyle;
    band1Horz?: TableConditionalStyle;
    band2Horz?: TableConditionalStyle;
    band1Vert?: TableConditionalStyle;
    band2Vert?: TableConditionalStyle;
    nwCell?: TableConditionalStyle;
    neCell?: TableConditionalStyle;
    swCell?: TableConditionalStyle;
    seCell?: TableConditionalStyle;
  } = {};
  const setCS = (key: keyof typeof conditionalStyles, type: string): void => {
    const val = resolveTableStyleConditional(styleResolver, conditionalTableStyleId, type);
    if (val) {
      conditionalStyles[key] = val;
    }
  };
  setCS("wholeTable", "wholeTable");
  const wholeTableStyle = mergeConditionalStyles(
    resolveTableBaseStyle(styleResolver, conditionalTableStyleId),
    conditionalStyles.wholeTable,
  );
  if (wholeTableStyle) {
    conditionalStyles.wholeTable = wholeTableStyle;
  }
  setCS("firstRow", "firstRow");
  setCS("lastRow", "lastRow");
  setCS("firstCol", "firstCol");
  setCS("lastCol", "lastCol");
  setCS("band1Horz", "band1Horz");
  setCS("band2Horz", "band2Horz");
  setCS("band1Vert", "band1Vert");
  setCS("band2Vert", "band2Vert");
  setCS("nwCell", "nwCell");
  setCS("neCell", "neCell");
  setCS("swCell", "swCell");
  setCS("seCell", "seCell");

  const bandingEnabledH = !look.noHBand;
  const bandingEnabledV = !look.noVBand;

  // Track data row index (excluding header rows) for banding
  let dataRowIndex = 0;
  const totalRows = table.rows.length;
  const gridColumnCount = columnWidths?.length ?? 0;
  const totalColumns = gridColumnCount > 0 ? gridColumnCount : countTableColumns(table.rows);
  const rows = table.rows.map((row, rowIndex) => {
    // Conditional formatting flag: firstRow in tblLook means "apply first-row styling"
    const isFirstRowStyled = rowIndex === 0 && look.firstRow;
    const isLastRow = rowIndex === totalRows - 1 && look.lastRow;

    const rowBandStyle = (() => {
      if (bandingEnabledH && !isFirstRowStyled && !isLastRow) {
        return (() => {
          if (dataRowIndex % 2 === 0) {
            return conditionalStyles.band1Horz;
          }
          return conditionalStyles.band2Horz;
        })();
      }
      return undefined;
    })();
    if (bandingEnabledH && !isFirstRowStyled && !isLastRow) {
      dataRowIndex++;
    }
    const resolvedRowJustification =
      tableStyle?.trPr?.justification ?? fallbackTableStyle?.trPr?.justification;

    return convertTableRow(
      row,
      styleResolver,
      context,
      isFirstRowStyled,
      columnWidths,
      totalWidth,
      conditionalStyles,
      rowBandStyle,
      bandingEnabledV,
      look,
      resolvedTableBorders, // Pass resolved table borders (own or from style)
      rowIndex,
      totalRows,
      totalColumns,
      rowSpanMap,
      cellMarginsAttr,
      resolvedRowJustification,
    );
  });

  return schema.node("table", attrs, rows);
}

function countTableColumns(rows: TableRow[]): number {
  let maxColumns = 0;
  for (const row of rows) {
    let rowColumns = row.formatting?.gridBefore ?? 0;
    for (const cell of row.cells) {
      rowColumns += cell.formatting?.gridSpan ?? 1;
    }
    rowColumns += row.formatting?.gridAfter ?? 0;
    maxColumns = Math.max(maxColumns, rowColumns);
  }
  return maxColumns;
}

type OmittedGridSlot = NonNullable<TableCellAttrs["_omittedGridSlot"]>;

function createOmittedGridSlotCell({
  colspan,
  slot,
  columnWidths,
  totalWidth,
  startColumn,
}: {
  colspan: number;
  slot: OmittedGridSlot;
  columnWidths: number[] | undefined;
  totalWidth: number | undefined;
  startColumn: number;
}): PMNode {
  const attrs: TableCellAttrs = { colspan, rowspan: 1, _omittedGridSlot: slot };
  if (columnWidths && totalWidth && totalWidth > 0) {
    const slotWidth = columnWidths
      .slice(startColumn, startColumn + colspan)
      .reduce((sum, width) => sum + width, 0);
    attrs.width = Math.round((slotWidth / totalWidth) * 100);
    attrs.widthType = "pct";
  }
  return schema.node("tableCell", attrs, [schema.node("paragraph")]);
}

/**
 * Convert a TableRow to a ProseMirror table row node
 */
function convertTableRow(
  row: TableRow,
  styleResolver: StyleEngine | null,
  context: TableConversionContext,
  isHeaderRow: boolean,
  columnWidths?: number[],
  totalWidth?: number,
  conditionalStyles?: {
    wholeTable?: TableConditionalStyle;
    firstRow?: TableConditionalStyle;
    lastRow?: TableConditionalStyle;
    firstCol?: TableConditionalStyle;
    lastCol?: TableConditionalStyle;
    band1Horz?: TableConditionalStyle;
    band2Horz?: TableConditionalStyle;
    band1Vert?: TableConditionalStyle;
    band2Vert?: TableConditionalStyle;
    nwCell?: TableConditionalStyle;
    neCell?: TableConditionalStyle;
    swCell?: TableConditionalStyle;
    seCell?: TableConditionalStyle;
  },
  rowBandStyle?: TableConditionalStyle,
  bandingEnabledV?: boolean,
  tableLook?: ResolvedTableLook,
  tableBorders?: TableBorders,
  rowIndex?: number,
  totalRows?: number,
  totalColumns?: number,
  rowSpanMap?: Map<string, RowSpanInfo>,
  defaultCellMargins?: TableCellMarginsAttrs,
  resolvedJustification?: NonNullable<TableRowFormatting["justification"]>,
): PMNode {
  const attrsWithoutStructuralChange: Omit<TableRowAttrs, "trIns" | "trDel"> = {
    // isHeader controls header row REPETITION on page breaks.
    // Only w:tblHeader (row.formatting.header) should trigger this — NOT tblLook/firstRow
    // which is purely a conditional formatting flag (ECMA-376 §17.7.6.1).
    isHeader: !!row.formatting?.header,
  };
  if (row.formatting?.height?.value !== undefined) {
    attrsWithoutStructuralChange.height = row.formatting.height.value;
  }
  if (row.formatting?.heightRule) {
    attrsWithoutStructuralChange.heightRule = row.formatting.heightRule;
  }
  if (row.formatting?.hidden) {
    attrsWithoutStructuralChange.hidden = true;
  }
  if (resolvedJustification) {
    attrsWithoutStructuralChange._resolvedJustification = resolvedJustification;
  }
  if (row.formatting) {
    attrsWithoutStructuralChange._originalFormatting = row.formatting;
  }
  // Carry `w:trPrChange` opaquely through PM for round-trip + accept/reject.
  if (row.propertyChanges && row.propertyChanges.length > 0) {
    attrsWithoutStructuralChange.trPrChange = [...row.propertyChanges];
  }
  // The row's attribute remainder, carried by identity for the reason the
  // paragraph's is.
  if (row.preservedAttributes && row.preservedAttributes.length > 0) {
    attrsWithoutStructuralChange._preservedAttributes = row.preservedAttributes;
  }
  let attrs: TableRowAttrs = attrsWithoutStructuralChange;
  const rowStructuralChange = row.structuralChange;
  if (rowStructuralChange) {
    const revision = {
      revisionId: rowStructuralChange.info.id,
      author: rowStructuralChange.info.author,
      date: rowStructuralChange.info.date ?? null,
      ...(rowStructuralChange.info.utcDate
        ? { utcDate: rowStructuralChange.info.utcDate.value }
        : {}),
      ...(rowStructuralChange.info.initials ? { initials: rowStructuralChange.info.initials } : {}),
    };
    switch (rowStructuralChange.type) {
      case "tableRowInsertion":
        attrs = { ...attrsWithoutStructuralChange, trIns: revision };
        break;
      case "tableRowDeletion":
        attrs = { ...attrsWithoutStructuralChange, trDel: revision };
        break;
      // A cell-level revision rides on the cell, not on the row that holds it.
      case "tableCellInsertion":
      case "tableCellDeletion":
      case "tableCellMerge":
        break;
      default: {
        const unsupported: never = rowStructuralChange;
        panic(`Unsupported table row structural change: ${JSON.stringify(unsupported)}`);
      }
    }
  }

  const numCells = row.cells.length;
  const gridBefore = row.formatting?.gridBefore ?? 0;
  const gridAfter = row.formatting?.gridAfter ?? 0;
  const isFirstRow = rowIndex === 0;
  const isLastRow = rowIndex === (totalRows ?? 1) - 1;
  const rowCnf = row.formatting?.conditionalFormat;
  const rowIsFirstRow = rowCnf?.firstRow ?? isFirstRow;
  const rowIsLastRow = rowCnf?.lastRow ?? isLastRow;
  const totalCols = totalColumns != null && totalColumns > 0 ? totalColumns : Math.max(numCells, 1);

  // A literal `<w:tr/>` from a non-Word producer parses with zero cells. PM's
  // tableRow content is `(tableCell | tableHeader)+`, so emit one placeholder
  // cell spanning the table's grid width to keep the row valid.
  let effectiveCells: TableCell[] = row.cells;
  const uncoveredColumns = Math.max(0, totalCols - gridBefore - gridAfter);
  if (effectiveCells.length === 0 && uncoveredColumns > 0) {
    const fallback: TableCell = {
      type: "tableCell",
      // convertTableCell supplies the PM-required empty paragraph. Keeping
      // this source-model placeholder empty avoids inventing authored nodes
      // that cannot belong to the entrypoint's source ownership index.
      content: [],
    };
    if (uncoveredColumns > 1) {
      fallback.formatting = { gridSpan: uncoveredColumns };
    }
    effectiveCells = [fallback];
  }

  // Track column index for mapping to columnWidths (accounting for colspan)
  let colIndex = gridBefore;
  const cells: PMNode[] = [];
  if (gridBefore > 0) {
    cells.push(
      createOmittedGridSlotCell({
        colspan: gridBefore,
        slot: "before",
        columnWidths,
        totalWidth,
        startColumn: 0,
      }),
    );
  }

  for (const cellIndex_item of effectiveCells) {
    const cell = cellIndex_item;
    const colspan = cell.formatting?.gridSpan ?? 1;

    // Check if this cell should be skipped (it's a vMerge continue cell)
    const rowSpanKey = `${rowIndex ?? 0}-${colIndex}`;
    const rowSpanInfo = rowSpanMap?.get(rowSpanKey);
    const shouldSkip = rowSpanInfo?.skip ?? false;
    const calculatedRowSpan = rowSpanInfo?.rowSpan ?? 1;
    const preserveVMergeRestart = rowSpanInfo?.preserveVMergeRestart ?? false;

    // Calculate the width for this cell from columnWidths if cell doesn't have own width
    let gridWidth: number | undefined;
    if (columnWidths && totalWidth && totalWidth > 0) {
      // Sum widths for all columns this cell spans
      let cellWidthTwips = 0;
      for (let i = 0; i < colspan && colIndex + i < columnWidths.length; i++) {
        cellWidthTwips += columnWidths[colIndex + i] ?? 0;
      }
      // Convert to percentage of total table width
      gridWidth = Math.round((cellWidthTwips / totalWidth) * 100);
    }
    colIndex += colspan;

    // Skip cells that are part of a vertical merge (vMerge="continue")
    if (shouldSkip) {
      continue;
    }

    // Determine cell position for table border application
    const isFirstCol = colIndex - colspan === 0;
    const isLastCol = colIndex === totalCols;
    const cellCnf = cell.formatting?.conditionalFormat;
    const cellIsFirstRow = cellCnf?.firstRow ?? rowIsFirstRow;
    const cellIsLastRow = cellCnf?.lastRow ?? rowIsLastRow;
    const cellIsFirstCol = cellCnf?.firstColumn ?? isFirstCol;
    const cellIsLastCol = cellCnf?.lastColumn ?? isLastCol;

    // Determine vertical banding style based on column index
    let vertBandStyle: TableConditionalStyle | undefined;
    if (bandingEnabledV) {
      const firstColOffset = tableLook?.firstColumn ? 1 : 0;
      const bandColIndex = colIndex - colspan - firstColOffset;
      const isEligible =
        bandColIndex >= 0 &&
        !(tableLook?.lastColumn && cellIsLastCol) &&
        !(tableLook?.firstColumn && cellIsFirstCol);
      if (isEligible) {
        vertBandStyle =
          bandColIndex % 2 === 0 ? conditionalStyles?.band1Vert : conditionalStyles?.band2Vert;
      }
    }

    if (cellCnf?.oddVBand) {
      vertBandStyle = conditionalStyles?.band1Vert;
    } else if (cellCnf?.evenVBand) {
      vertBandStyle = conditionalStyles?.band2Vert;
    }

    let effectiveRowBandStyle = rowBandStyle;
    if (rowCnf?.oddHBand) {
      effectiveRowBandStyle = conditionalStyles?.band1Horz;
    } else if (rowCnf?.evenHBand) {
      effectiveRowBandStyle = conditionalStyles?.band2Horz;
    }
    if (cellCnf?.oddHBand) {
      effectiveRowBandStyle = conditionalStyles?.band1Horz;
    } else if (cellCnf?.evenHBand) {
      effectiveRowBandStyle = conditionalStyles?.band2Horz;
    }

    // Build conditional style precedence (wholeTable -> banding -> columns -> rows -> corners)
    let cellConditionalStyle = conditionalStyles?.wholeTable;
    cellConditionalStyle = mergeConditionalStyles(cellConditionalStyle, effectiveRowBandStyle);
    cellConditionalStyle = mergeConditionalStyles(cellConditionalStyle, vertBandStyle);
    if (cellIsFirstCol && (tableLook?.firstColumn || rowCnf?.firstColumn || cellCnf?.firstColumn)) {
      cellConditionalStyle = mergeConditionalStyles(
        cellConditionalStyle,
        conditionalStyles?.firstCol,
      );
    }
    if (cellIsLastCol && (tableLook?.lastColumn || rowCnf?.lastColumn || cellCnf?.lastColumn)) {
      cellConditionalStyle = mergeConditionalStyles(
        cellConditionalStyle,
        conditionalStyles?.lastCol,
      );
    }
    if (cellIsFirstRow && (tableLook?.firstRow || rowCnf?.firstRow || cellCnf?.firstRow)) {
      cellConditionalStyle = mergeConditionalStyles(
        cellConditionalStyle,
        conditionalStyles?.firstRow,
      );
    }
    if (cellIsLastRow && (tableLook?.lastRow || rowCnf?.lastRow || cellCnf?.lastRow)) {
      cellConditionalStyle = mergeConditionalStyles(
        cellConditionalStyle,
        conditionalStyles?.lastRow,
      );
    }
    if (
      cellIsFirstRow &&
      cellIsFirstCol &&
      (tableLook?.firstRow || rowCnf?.firstRow || cellCnf?.firstRow) &&
      (tableLook?.firstColumn || rowCnf?.firstColumn || cellCnf?.firstColumn)
    ) {
      cellConditionalStyle = mergeConditionalStyles(
        cellConditionalStyle,
        conditionalStyles?.nwCell,
      );
    }
    if (
      cellIsFirstRow &&
      cellIsLastCol &&
      (tableLook?.firstRow || rowCnf?.firstRow || cellCnf?.firstRow) &&
      (tableLook?.lastColumn || rowCnf?.lastColumn || cellCnf?.lastColumn)
    ) {
      cellConditionalStyle = mergeConditionalStyles(
        cellConditionalStyle,
        conditionalStyles?.neCell,
      );
    }
    if (
      cellIsLastRow &&
      cellIsFirstCol &&
      (tableLook?.lastRow || rowCnf?.lastRow || cellCnf?.lastRow) &&
      (tableLook?.firstColumn || rowCnf?.firstColumn || cellCnf?.firstColumn)
    ) {
      cellConditionalStyle = mergeConditionalStyles(
        cellConditionalStyle,
        conditionalStyles?.swCell,
      );
    }
    if (
      cellIsLastRow &&
      cellIsLastCol &&
      (tableLook?.lastRow || rowCnf?.lastRow || cellCnf?.lastRow) &&
      (tableLook?.lastColumn || rowCnf?.lastColumn || cellCnf?.lastColumn)
    ) {
      cellConditionalStyle = mergeConditionalStyles(
        cellConditionalStyle,
        conditionalStyles?.seCell,
      );
    }

    cells.push(
      convertTableCell({
        cell,
        styleResolver,
        context,
        isHeader: isHeaderRow,
        gridWidthPercent: gridWidth,
        conditionalStyle: cellConditionalStyle,
        tableBorders,
        position: { isFirstRow, isLastRow, isFirstColumn: isFirstCol, isLastColumn: isLastCol },
        calculatedRowSpan,
        preserveVMergeRestart,
        vMergeContinuationCells: rowSpanInfo?.continuationCells,
        defaultCellMargins,
      }),
    );
  }

  if (gridAfter > 0) {
    cells.push(
      createOmittedGridSlotCell({
        colspan: gridAfter,
        slot: "after",
        columnWidths,
        totalWidth,
        startColumn: colIndex,
      }),
    );
  }

  return schema.node("tableRow", attrs, cells);
}

type ConvertTableCellOptions = {
  cell: TableCell;
  styleResolver: StyleEngine | null;
  context: TableConversionContext;
  isHeader: boolean;
  gridWidthPercent: number | undefined;
  conditionalStyle: TableConditionalStyle | undefined;
  tableBorders: TableBorders | undefined;
  position: TableCellPosition;
  calculatedRowSpan: number | undefined;
  preserveVMergeRestart: boolean | undefined;
  vMergeContinuationCells: TableCell[] | undefined;
  defaultCellMargins: TableCellMarginsAttrs | undefined;
};

/**
 * Convert a TableCell to a ProseMirror table cell node
 */
function convertTableCell({
  cell,
  styleResolver,
  context,
  isHeader,
  gridWidthPercent,
  conditionalStyle,
  tableBorders,
  position,
  calculatedRowSpan,
  preserveVMergeRestart,
  vMergeContinuationCells,
  defaultCellMargins,
}: ConvertTableCellOptions): PMNode {
  const formatting = cell.formatting;

  // Use the pre-calculated rowSpan from vMerge analysis
  const rowspan = calculatedRowSpan ?? 1;
  const effectiveFormatting = resolveEffectiveTableCellFormatting({
    directFormatting: formatting,
    styleFormatting: conditionalStyle?.tcPr,
    tableBorders,
    position,
    gridWidthPercent,
    defaultMargins: defaultCellMargins,
    theme: context.theme,
  });

  const attrs: TableCellAttrs = {
    colspan: formatting?.gridSpan ?? 1,
    rowspan,
  };
  if (effectiveFormatting.width.type === "value") {
    attrs.width = effectiveFormatting.width.value;
    if (effectiveFormatting.width.widthType) {
      attrs.widthType = effectiveFormatting.width.widthType;
    }
  }
  if (formatting?.verticalAlign) {
    attrs.verticalAlign = formatting.verticalAlign;
  }
  if (effectiveFormatting.background.type === "color") {
    attrs.backgroundColor = effectiveFormatting.background.rgb;
    attrs._resolvedBackgroundColor = effectiveFormatting.background.rgb;
  }
  if (formatting?.textDirection) {
    attrs.textDirection = formatting.textDirection;
  }
  if (formatting?.noWrap !== undefined) {
    attrs.noWrap = formatting.noWrap;
  }
  const hideMark = formatting?.hideMark ?? conditionalStyle?.tcPr?.hideMark;
  if (hideMark !== undefined) {
    attrs.hideMark = hideMark;
  }
  if (effectiveFormatting.borders) {
    attrs.borders = effectiveFormatting.borders;
    attrs._resolvedBorders = effectiveFormatting.borders;
  }
  if (effectiveFormatting.margins) {
    attrs.margins = effectiveFormatting.margins;
    attrs._resolvedMargins = effectiveFormatting.margins;
  }
  if (formatting) {
    attrs._originalFormatting = formatting;
  }
  // Carry `w:tcPrChange` opaquely through PM for round-trip + accept/reject.
  if (cell.propertyChanges && cell.propertyChanges.length > 0) {
    attrs.tcPrChange = [...cell.propertyChanges];
  }
  const cellStructuralChange = cell.structuralChange;
  if (cellStructuralChange) {
    const info = {
      revisionId: cellStructuralChange.info.id,
      author: cellStructuralChange.info.author,
      date: cellStructuralChange.info.date ?? null,
      ...(cellStructuralChange.info.utcDate
        ? { utcDate: cellStructuralChange.info.utcDate.value }
        : {}),
      ...(cellStructuralChange.info.initials
        ? { initials: cellStructuralChange.info.initials }
        : {}),
    };
    switch (cellStructuralChange.type) {
      case "tableCellInsertion":
        attrs.cellMarker = { kind: "ins", info };
        break;
      case "tableCellDeletion":
        attrs.cellMarker = { kind: "del", info };
        break;
      case "tableCellMerge":
        attrs.cellMarker = {
          kind: "merge",
          info,
          ...(cellStructuralChange.verticalMerge !== undefined
            ? { verticalMerge: cellStructuralChange.verticalMerge }
            : {}),
          ...(cellStructuralChange.verticalMergeOriginal !== undefined
            ? { verticalMergeOriginal: cellStructuralChange.verticalMergeOriginal }
            : {}),
        };
        break;
      // A row-level revision rides on the row, not on the cells inside it.
      case "tableRowInsertion":
      case "tableRowDeletion":
        break;
      default: {
        const unsupported: never = cellStructuralChange;
        panic(`Unsupported table cell structural change: ${JSON.stringify(unsupported)}`);
      }
    }
  }
  if (preserveVMergeRestart) {
    attrs._preserveVMergeRestart = true;
  }
  if (vMergeContinuationCells && vMergeContinuationCells.length > 0) {
    attrs._docxVMergeContinuationCells =
      transportTableCellsWithParagraphPropertySources(vMergeContinuationCells);
  }

  // Convert cell content (paragraphs and nested tables)
  const contentNodes: PMNode[] = [];
  for (const content of cell.content) {
    switch (content.type) {
      case "paragraph":
        contentNodes.push(
          ...convertParagraphWithTextBoxes(content, styleResolver, {
            textBoxGroupId: context.nextTextBoxGroupId(),
            context,
            ...(conditionalStyle?.rPr !== undefined
              ? { extraRunFormatting: conditionalStyle.rPr }
              : {}),
            ...(conditionalStyle?.pPr !== undefined
              ? { tableParagraphOverlay: conditionalStyle.pPr }
              : {}),
          }),
        );
        break;
      case "table":
        contentNodes.push(convertTable(content, styleResolver, context));
        break;
      case "preservedBlock":
        contentNodes.push(convertPreservedBlock(content));
        break;
      default: {
        const unsupported: never = content;
        panic(`Unsupported table cell content: ${JSON.stringify(unsupported)}`);
      }
    }
  }

  // Ensure cell has at least one paragraph
  if (contentNodes.length === 0) {
    contentNodes.push(schema.node("paragraph", {}, []));
  }

  // Use tableHeader for header cells, tableCell otherwise
  const nodeType = isHeader ? "tableHeader" : "tableCell";
  return schema.node(nodeType, attrs, contentNodes);
}

export function standaloneTableCellToProseMirror(
  cell: TableCell,
  nodeType: "tableCell" | "tableHeader",
): PMNode {
  const nextTextBoxGroupId = createTextBoxGroupIdFactory();
  const nextHyperlinkInstanceIndex = createHyperlinkInstanceIndexAllocator();
  const pageBreakRunSourceDescendants = buildPageBreakRunSourceDescendantIndex(cell.content);
  reportSourceContainerPageBreakRun(
    cell.content,
    "table-cell",
    pageBreakRunSourceDescendants,
    noPageBreakProjectionWarning,
  );
  return convertTableCell({
    cell,
    styleResolver: null,
    context: {
      theme: null,
      nextTextBoxGroupId,
      nextHyperlinkInstanceIndex,
      pairedBookmarkIds: collectPairedBookmarkIds(cell.content),
      pageBreakRunSourceDescendants,
      storyRangedCommentIds: rangedCommentIds(cell.content),
      openCommentIds: new Set<number>(),
      warnPageBreakProjection: noPageBreakProjectionWarning,
    },
    isHeader: nodeType === "tableHeader",
    gridWidthPercent: undefined,
    conditionalStyle: undefined,
    tableBorders: undefined,
    position: {},
    calculatedRowSpan: 1,
    preserveVMergeRestart: undefined,
    vMergeContinuationCells: undefined,
    defaultCellMargins: undefined,
  });
}

/**
 * Convert a SimpleField or ComplexField to a ProseMirror field node.
 * Preserves run formatting (bold, fontSize, color, etc.) as PM marks.
 * Accepts a run formatting resolver so fields inherit paragraph-level
 * formatting the same way regular text runs do.
 */
type ConvertFieldOptions = {
  getInheritedRunFormatting: RunFormattingResolver;
  styleResolver: StyleEngine | null | undefined;
  nextHyperlinkInstanceIndex: HyperlinkInstanceIndexAllocator;
  nextPageBreakRunOwnerId: PageBreakRunOwnerIdAllocator;
  textBoxAnchors: ReadonlyMap<Shape, string> | undefined;
};

function convertField(
  field: SimpleField | ComplexField,
  {
    getInheritedRunFormatting,
    styleResolver,
    nextHyperlinkInstanceIndex,
    nextPageBreakRunOwnerId,
    textBoxAnchors,
  }: ConvertFieldOptions,
): PMNode | null {
  // Extract display text and formatting from field content/result
  let displayText = "";
  let fieldFormatting: TextFormatting | undefined;
  let fieldPropertyChanges: readonly RunPropertyChange[] | undefined;
  const inlineNodes: PMNode[] = [];
  const hasPageBreakContent = fieldResultHasPageBreakContent(field);
  // A capture has no other carrier: a field collapsed to its display text
  // would drop the markup, so a field holding one keeps its children.
  const hasStructuredSourceContent =
    hasPageBreakContent ||
    (field.type === "simpleField" &&
      field.content.some(
        (content) => content.type === "hyperlink" || content.type === "preservedInline",
      ));
  const appendRun = (run: Run): void => {
    for (const content of run.content) {
      if (content.type === "text") {
        displayText += content.text;
      }
    }
    // Use formatting from the first run that has it.
    fieldFormatting ??= run.formatting;
    fieldPropertyChanges ??= run.propertyChanges;
    if (!hasStructuredSourceContent) {
      return;
    }
    inlineNodes.push(
      ...convertRun(
        run,
        getInheritedRunFormatting(run.formatting, field.fieldType),
        nextPageBreakRunOwnerId,
        styleResolver,
        textBoxAnchors,
      ),
    );
  };
  if (field.type === "simpleField") {
    for (const content of field.content) {
      if (content.type === "run") {
        appendRun(content);
        continue;
      }
      if (content.type === "preservedInline") {
        inlineNodes.push(preservedInlineNode(content));
        continue;
      }
      for (const child of content.children) {
        if (child.type === "run") {
          for (const runContent of child.content) {
            if (runContent.type === "text") {
              displayText += runContent.text;
            }
          }
          fieldFormatting ??= child.formatting;
          fieldPropertyChanges ??= child.propertyChanges;
        }
      }
      inlineNodes.push(
        ...convertHyperlink(content, {
          getInheritedRunFormatting: (formatting) =>
            getInheritedRunFormatting(formatting, field.fieldType),
          styleResolver,
          hyperlinkIndex: nextHyperlinkInstanceIndex(),
          textBoxAnchors,
          nextPageBreakRunOwnerId,
        }),
      );
    }
  } else {
    for (const run of field.fieldResult) {
      appendRun(run);
    }
  }

  // Collapsed complex field with no result run at all — fall back to the field's
  // captured run formatting so a footer PAGE number keeps its size/color instead
  // of rendering at the default (eigenpal/docx-editor#909). Only when the result
  // is genuinely empty: a present-but-unformatted result run intentionally
  // inherits paragraph defaults and must not adopt the field code's formatting.
  if (!fieldFormatting && field.type === "complexField" && field.fieldResult.length === 0) {
    fieldFormatting = field.formatting;
  }

  // Merge inherited paragraph formatting, the field run's own character style,
  // then its inline formatting (each later layer wins) — the same precedence as
  // `convertRun`. Flattening the character style here means a field result
  // styled only by a `w:rStyle` (e.g. a PAGE/REF field) keeps its style-derived
  // marks, so the save-side `w:rStyle` reconciliation does not treat the field
  // as diverging and strip the link (eigenpal/docx-editor#833).
  const inheritedFormatting = getInheritedRunFormatting(fieldFormatting, field.fieldType);
  const { marks } = buildRunMarks(fieldFormatting, inheritedFormatting, styleResolver);

  const hasConvertedHyperlinkContent = inlineNodes.some((node) =>
    node.marks.some((mark) => mark.type.name === "hyperlink"),
  );
  const hasConvertedPageBreakContent = inlineNodes.some(
    (node) => node.type.name === "pageBreakRun",
  );
  const hasConvertedPreservedContent = inlineNodes.some(
    (node) => node.type.name === "preservedXml",
  );
  const createStructuredField =
    hasConvertedPageBreakContent ||
    hasConvertedPreservedContent ||
    (hasStructuredSourceContent && hasConvertedHyperlinkContent);
  if (!createStructuredField && fieldPropertyChanges && fieldPropertyChanges.length > 0) {
    marks.push(schema.mark("runPropertyChange", { changes: [...fieldPropertyChanges] }));
  }
  return schema.node(
    createStructuredField ? "structuredField" : "field",
    {
      fieldType: field.fieldType,
      instruction: field.instruction,
      displayText,
      fieldKind: field.type === "simpleField" ? "simple" : "complex",
      fldLock: field.fldLock ?? null,
      dirty: field.dirty ?? null,
    },
    createStructuredField ? inlineNodes : undefined,
    marks,
  );
}

/** An inline item the wrappers have been lifted off, and which ones they were. */
type StackedContent = {
  content: Exclude<ParagraphContent, InlineWrapper>;
  /** Outermost first; empty when the item sat inside no wrapper. */
  stack: readonly InlineWrapperLayer[];
};

/**
 * Lift the transparent wrappers out of an inline content list, remembering for
 * each item which wrappers it sat inside.
 *
 * The tree becomes one sequence because every inline loop below narrows by a
 * chain of `else if` rather than by exhaustion, and a wrapper left in it would
 * reach whichever branch happens to be last. What the wrapper said is not lost
 * with the tree: it rides the `inlineWrapper` mark on the leaves it held, and
 * ProseMirror maintains that range across edits instead of an index.
 *
 * It is applied to the content of whatever holds the wrapper, never around it:
 * a wrapper inside a revision or a content control is lifted inside that
 * wrapper, so its runs keep the revision mark or the control they were
 * authored under.
 */
const withInlineWrapperStacks = (
  content: readonly ParagraphContent[],
  stack: readonly InlineWrapperLayer[] = [],
): StackedContent[] =>
  content.flatMap((item) =>
    item.type === "inlineWrapper"
      ? withInlineWrapperStacks(item.content, [...stack, inlineWrapperLayer(item)])
      : [{ content: item, stack }],
  );

/**
 * `nodes` with `stack` recorded outside whatever wrapper they already carry.
 *
 * A revision converts its own content first, so a node that arrives here
 * already marked sat inside the item this stack wraps: the two stacks
 * concatenate, outermost first. Concatenating is what the schema forces — the
 * mark excludes itself, so a second one replaces the first, and the inner
 * wrapper would be the one lost.
 *
 * A node that holds its own leaves, a content control, is marked here as the
 * one node it is; its leaves were marked with the whole enclosing stack when
 * it was built, because the painter reads the wrapper off the leaf.
 */
const withInlineWrapperMark = (nodes: PMNode[], stack: readonly InlineWrapperLayer[]): PMNode[] => {
  const markType = schema.marks["inlineWrapper"];
  if (stack.length === 0 || !markType) {
    return nodes;
  }
  return nodes.map((node) => {
    if (!node.isText && (!node.isInline || !node.type.allowsMarkType(markType))) {
      return node;
    }
    const inner = node.marks.find((mark) => mark.type === markType);
    const layers =
      inner === undefined ? stack : [...stack, ...expectInlineWrapperMarkAttrs(inner).stack];
    return node.mark(markType.create({ stack: layers }).addToSet(node.marks));
  });
};

/**
 * Convert a MathEquation to a ProseMirror math node.
 */
function convertMathEquation(math: MathEquation): PMNode | null {
  return schema.node("math", {
    display: math.display,
    ommlXml: math.ommlXml,
    plainText: math.plainText || "",
  });
}

/**
 * Convert an InlineSdt to a ProseMirror sdt node with inline content.
 *
 * `wrappedBy` is the stack the control itself sits inside. The control holds
 * its own leaves, so nothing outside can mark them afterwards: the caller's
 * stack has to reach them here, under the wrappers written inside the control.
 */
function convertInlineSdt(
  sdt: InlineSdt,
  nextHyperlinkInstanceIndex: HyperlinkInstanceIndexAllocator,
  nextPageBreakRunOwnerId: PageBreakRunOwnerIdAllocator,
  getInheritedRunFormatting: RunFormattingResolver,
  trackedRunFormattingResolvers: TrackedRunFormattingResolvers,
  styleResolver?: StyleEngine | null,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
  wrappedBy: readonly InlineWrapperLayer[] = [],
): PMNode | null {
  const props = sdt.properties;
  const inlineNodes: PMNode[] = [];

  // A wrapper inside the control is lifted here rather than out of it: a
  // wrapper lifted out of the control takes the control's content with it.
  for (const { content, stack } of withInlineWrapperStacks(sdt.content, wrappedBy)) {
    if (!isInlineSdtContent(content)) {
      continue;
    }
    const itemNodes: PMNode[] = [];
    switch (content.type) {
      case "run":
        itemNodes.push(
          ...convertRun(
            content,
            getInheritedRunFormatting(content.formatting),
            nextPageBreakRunOwnerId,
            styleResolver,
            textBoxAnchors,
          ),
        );
        break;
      case "hyperlink": {
        const currentHyperlinkIndex = nextHyperlinkInstanceIndex();
        itemNodes.push(
          ...convertHyperlink(content, {
            getInheritedRunFormatting,
            styleResolver,
            hyperlinkIndex: currentHyperlinkIndex,
            textBoxAnchors,
            nextPageBreakRunOwnerId,
          }),
        );
        break;
      }
      case "simpleField":
      case "complexField": {
        const fieldNode = convertField(content, {
          getInheritedRunFormatting,
          styleResolver,
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          textBoxAnchors,
        });
        if (fieldNode) {
          itemNodes.push(fieldNode);
        }
        break;
      }
      case "inlineSdt": {
        const nestedSdt = convertInlineSdt(
          content,
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          getInheritedRunFormatting,
          trackedRunFormattingResolvers,
          styleResolver,
          textBoxAnchors,
          stack,
        );
        if (nestedSdt) {
          itemNodes.push(nestedSdt);
        }
        break;
      }
      case "insertion":
      case "deletion":
      case "moveTo":
      case "moveFrom":
        itemNodes.push(
          ...convertTrackedChange(
            content,
            content.type === "insertion" || content.type === "moveTo" ? "insertion" : "deletion",
            nextHyperlinkInstanceIndex,
            nextPageBreakRunOwnerId,
            trackedRunFormattingResolvers,
            styleResolver,
            content.type === "moveTo" || content.type === "moveFrom" ? content.type : null,
            textBoxAnchors,
          ),
        );
        break;
      case "mathEquation": {
        const mathNode = convertMathEquation(content);
        if (mathNode) {
          itemNodes.push(mathNode);
        }
        break;
      }
      case "preservedInline":
        itemNodes.push(preservedInlineNode(content));
        break;
      default: {
        const unsupported: never = content;
        panic(`Unsupported inline SDT content: ${JSON.stringify(unsupported)}`);
      }
    }
    inlineNodes.push(...withInlineWrapperMark(itemNodes, stack));
  }

  return schema.node(
    "sdt",
    sdtAttrsFromProperties(props),
    inlineNodes.length > 0 ? inlineNodes : undefined,
  );
}

/**
 * Convert a Run to ProseMirror text nodes with marks
 *
 * @param run - The run to convert
 * @param resolvedStyleFormatting - Paragraph formatting and its explicit style-toggle state
 */
function convertRun(
  run: Run,
  resolvedStyleFormatting: ResolvedRunFormatting,
  nextPageBreakRunOwnerId: PageBreakRunOwnerIdAllocator,
  styleResolver?: StyleEngine | null,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
): PMNode[] {
  const nodes: PMNode[] = [];
  const { marks, mergedFormatting } = buildRunMarks(
    run.formatting,
    resolvedStyleFormatting,
    styleResolver,
  );
  if (run.propertyChanges && run.propertyChanges.length > 0) {
    marks.push(schema.mark("runPropertyChange", { changes: [...run.propertyChanges] }));
  }
  if (run.content.some((content) => content.type === "break" && content.breakType === "page")) {
    marks.push(schema.mark("pageBreakRunOwner", { id: nextPageBreakRunOwnerId() }));
  }

  for (const content of run.content) {
    const contentNodes = convertRunContent(
      content,
      marks,
      mergedFormatting,
      textBoxAnchors,
      run.formatting,
    );
    nodes.push(...contentNodes);
  }

  return nodes;
}

const runHasPageBreakContent = (run: Run): boolean =>
  run.content.some((content) => content.type === "break" && content.breakType === "page");

/**
 * Whether a field's result holds an explicit page break.
 *
 * `convertField` re-cuts the whole result into a structured field when it
 * does, so this answers both who gets re-cut and who gets reported. One
 * predicate for the two, because a reporter that disagreed with the converter
 * would name a loss nobody suffered.
 */
const fieldResultHasPageBreakContent = (field: SimpleField | ComplexField): boolean =>
  field.type === "simpleField"
    ? field.content.some((content) => {
        switch (content.type) {
          case "run":
            return runHasPageBreakContent(content);
          case "hyperlink":
            return content.children.some(
              (child) => child.type === "run" && runHasPageBreakContent(child),
            );
          case "preservedInline":
            return false;
          default: {
            const unsupported: never = content;
            panic(`Unsupported simple-field content: ${JSON.stringify(unsupported)}`);
          }
        }
      })
    : field.fieldResult.some(runHasPageBreakContent);

function reportPageBreakSourceRunContent(run: Run, warn: PageBreakProjectionWarn): void {
  if (!runHasPageBreakContent(run)) {
    return;
  }

  reportRunContentBesidePageBreak(run, "page-break-bearing-run", warn);
}

const PAGE_BREAK_CONTAINER_DESCRIPTIONS = {
  "table-cell":
    "An interior page break in a table cell does not paginate the cell; its row moves as a unit",
  "text-box": "A page break inside a text box does not paginate the text box",
} as const satisfies Record<"table-cell" | "text-box", string>;

type PageBreakContainerOwner = keyof typeof PAGE_BREAK_CONTAINER_DESCRIPTIONS;

function reportSourceContainerPageBreakRun(
  content: BlockContent[],
  owner: PageBreakContainerOwner,
  sourceDescendants: PageBreakRunSourceDescendantIndex,
  warn: PageBreakProjectionWarn,
): void {
  if (!sourceDescendants.containsPageBreakRun(content)) {
    return;
  }
  if (owner === "table-cell" && hasSingleLeadingTableCellPageBreak(content, sourceDescendants)) {
    return;
  }
  warn(PAGE_BREAK_CONTAINER_DESCRIPTIONS[owner]);
}

/**
 * A page break opening a cell advances the whole row during pagination.
 *
 * Word writes this shape routinely and the row's own `breakBefore` projects it
 * losslessly, so what follows the break inside the cell rides along: the row
 * moves as a unit. The break has to be the cell's only one and has to open its
 * first paragraph; an interior break still requires table-fragment ownership,
 * which cell-local flow cannot model. `convertTableCell` in the layout bridge
 * decides the same question over the projected runs.
 */
function hasSingleLeadingTableCellPageBreak(
  content: BlockContent[],
  sourceDescendants: PageBreakRunSourceDescendantIndex,
): boolean {
  const first = content.at(0);
  if (first?.type !== "paragraph" || !hasSingleLeadingParagraphPageBreak(first)) {
    return false;
  }
  return !content.slice(1).some((block) => sourceDescendants.blockContainsPageBreakRun(block));
}

type LeadingPageBreakScan = {
  contentBeforeBreak: boolean;
  pageBreaks: number;
  textBoxShapeAfterBreak: boolean;
};

const scanLeadingPageBreakRun = (run: Run, scan: LeadingPageBreakScan): void => {
  for (const content of run.content) {
    if (content.type === "renderedPageBreak") {
      continue;
    }
    if (content.type === "break" && content.breakType === "page") {
      scan.pageBreaks += 1;
      continue;
    }
    // Only a shape with a text body anchors content: a plain shape after the
    // break loses no host, so it must not read as a text-box anchor.
    if (content.type === "shape" && content.shape.textBody && scan.pageBreaks > 0) {
      scan.textBoxShapeAfterBreak = true;
    }
    if (scan.pageBreaks === 0) {
      scan.contentBeforeBreak = true;
    }
  }
};

const scanLeadingPageBreakContent = (
  content: Paragraph["content"][number],
  scan: LeadingPageBreakScan,
): void => {
  switch (content.type) {
    case "run":
      scanLeadingPageBreakRun(content, scan);
      return;
    case "hyperlink":
      for (const child of content.children) scanLeadingPageBreakContent(child, scan);
      return;
    case "insertion":
    case "deletion":
    case "moveFrom":
    case "moveTo":
    case "inlineSdt":
    // Transparent: the scan is looking for a page break, and a wrapper says
    // how its content is laid out rather than where a page ends.
    case "inlineWrapper":
      for (const child of content.content) scanLeadingPageBreakContent(child, scan);
      return;
    case "simpleField":
      for (const child of content.content) scanLeadingPageBreakContent(child, scan);
      return;
    case "complexField":
      for (const child of content.fieldResult) scanLeadingPageBreakRun(child, scan);
      return;
    case "bookmarkStart":
    case "bookmarkEnd":
    case "commentRangeStart":
    case "commentRangeEnd":
    case "moveFromRangeStart":
    case "moveFromRangeEnd":
    case "moveToRangeStart":
    case "moveToRangeEnd":
      return;
    case "commentReference":
    case "mathEquation":
    case "preservedInline":
      if (scan.pageBreaks === 0) scan.contentBeforeBreak = true;
      return;
    default: {
      const unsupported: never = content;
      return unsupported;
    }
  }
};

const scanParagraphPageBreaks = (paragraph: Paragraph): LeadingPageBreakScan => {
  const scan: LeadingPageBreakScan = {
    contentBeforeBreak: false,
    pageBreaks: 0,
    textBoxShapeAfterBreak: false,
  };
  for (const content of paragraph.content) scanLeadingPageBreakContent(content, scan);
  return scan;
};

const hasSingleLeadingParagraphPageBreak = (paragraph: Paragraph): boolean => {
  const scan = scanParagraphPageBreaks(paragraph);
  return !scan.contentBeforeBreak && scan.pageBreaks === 1;
};

type ParagraphPageBreakProjectionOptions = {
  paragraph: Paragraph;
  attrs: ParagraphAttrs;
  effectiveFrame: ParagraphFormatting["frame"];
  sourceDescendants: PageBreakRunSourceDescendantIndex;
  warn: PageBreakProjectionWarn;
};

function reportParagraphPageBreakProjection({
  paragraph,
  attrs,
  effectiveFrame,
  sourceDescendants,
  warn,
}: ParagraphPageBreakProjectionOptions): void {
  const sourceFeatures = sourceDescendants.paragraphFeatures(paragraph);
  if (!sourceFeatures.hasPageBreakRun) {
    return;
  }
  reportParagraphPageBreakRunContent(paragraph, warn);

  const disposition = pageBreakRunParagraphProjectionDispositionForFeatures({
    attrs,
    effectiveFrame,
    textBoxAnchorAfterPageBreak:
      sourceFeatures.hasTextBoxShape &&
      !sourceFeatures.pageBreakSharesTextBoxShape &&
      scanParagraphPageBreaks(paragraph).textBoxShapeAfterBreak,
  });
  if (disposition.status === "exact") {
    return;
  }
  if (disposition.reason !== "textBoxAnchor" && hasSingleLeadingParagraphPageBreak(paragraph)) {
    return;
  }
  warn(disposition.message);
}

/**
 * Report every run-level loss a paragraph's page breaks bring with them.
 *
 * One walk of the paragraph rather than a check at each `convertRun` and
 * `convertField`: the reporter is a conversion-wide fact, and a per-call-site
 * check would have to be threaded through every nested converter to reach it.
 */
function reportParagraphPageBreakRunContent(
  paragraph: Paragraph,
  warn: PageBreakProjectionWarn,
): void {
  const visitField = (field: SimpleField | ComplexField): void => {
    if (field.type === "complexField") {
      if (field.fieldCode.some(runHasPageBreakContent)) {
        warn("A complex-field instruction holds an explicit page break");
      }
      if (!fieldResultHasPageBreakContent(field)) {
        return;
      }
      for (const run of field.fieldResult) {
        reportRunContentBesidePageBreak(run, "field-result", warn);
      }
      return;
    }
    if (!fieldResultHasPageBreakContent(field)) {
      return;
    }
    for (const content of field.content) {
      if (content.type === "run") {
        reportRunContentBesidePageBreak(content, "field-result", warn);
        continue;
      }
      if (content.type !== "hyperlink") {
        continue;
      }
      for (const child of content.children) {
        if (child.type === "run") {
          reportRunContentBesidePageBreak(child, "field-result", warn);
        }
      }
    }
  };

  const visitContent = (content: Paragraph["content"][number]): void => {
    switch (content.type) {
      case "run":
        reportPageBreakSourceRunContent(content, warn);
        return;
      case "hyperlink":
        for (const child of content.children) {
          if (child.type === "run") {
            reportPageBreakSourceRunContent(child, warn);
          }
        }
        return;
      case "simpleField":
      case "complexField":
        visitField(content);
        return;
      case "insertion":
      case "deletion":
      case "moveFrom":
      case "moveTo":
      case "inlineSdt":
      case "inlineWrapper":
        for (const child of content.content) visitContent(child);
        return;
      default:
        return;
    }
  };

  for (const content of paragraph.content) visitContent(content);
}

const PAGE_BREAK_OWNER_DESCRIPTIONS = {
  "field-result": "A field result with an explicit page break",
  "page-break-bearing-run": "A page-break-bearing run",
} as const satisfies Record<"field-result" | "page-break-bearing-run", string>;

type PageBreakContentOwner = keyof typeof PAGE_BREAK_OWNER_DESCRIPTIONS;

/**
 * Record what a page-break-bearing run loses when the editor re-cuts it.
 *
 * The run is split into one node per inline atom and rebuilt from the
 * `pageBreakRunOwner` mark, and the kinds below have no node of their own:
 * a hyphen returns as its character, a field character or instruction leaves
 * no trace, and a text-box shape is hoisted to its own block. Every one of
 * those losses is what the same run suffers with no page break in it, so
 * refusing the document here declined to open a file folio otherwise reads.
 */
function reportRunContentBesidePageBreak(
  run: Run,
  owner: PageBreakContentOwner,
  warn: PageBreakProjectionWarn,
): void {
  for (const content of run.content) {
    switch (content.type) {
      case "break":
      case "drawing":
      case "endnoteRef":
      case "footnoteRef":
      // An opaque atom rebuilds from its own attributes, exactly as a symbol
      // does, so the page-break owner can re-cut the run around it. Refusing
      // would cost the whole document the editor, which is the worse loss.
      case "preservedXml":
      case "renderedPageBreak":
      case "symbol":
      case "tab":
      case "text":
        continue;
      case "shape":
        if (content.shape.textBody) {
          warn(`${PAGE_BREAK_OWNER_DESCRIPTIONS[owner]} also holds a text-box shape`);
        }
        continue;
      case "fieldChar":
      case "instrText":
      case "noBreakHyphen":
      case "softHyphen":
        warn(`${PAGE_BREAK_OWNER_DESCRIPTIONS[owner]} also holds ${content.type}`);
        continue;
      default: {
        const unsupported: never = content;
        panic(`Unsupported page-break-bearing run content: ${JSON.stringify(unsupported)}`);
      }
    }
  }
}

/**
 * Build the marks for a run, layering per the OOXML cascade: inherited
 * (docDefaults + paragraph style) formatting, then the run's character style
 * (w:rStyle), then direct run formatting on top.
 *
 * Use getRunStyleOwnProperties (not resolveRunStyle) to avoid docDefaults
 * from the character style overriding paragraph style properties.
 * The inherited formatting already includes docDefaults from paragraph
 * style resolution, so we only need the character style's own properties.
 *
 * When the run references a character style, a compact `characterStyle` mark
 * carries its styleId so document-level style context can resolve it without
 * copying the style onto every run. Unknown styleIds resolve to nothing: no
 * formatting is flattened, and the reference round-trips verbatim.
 */
type BuiltRunMarks = {
  marks: ReturnType<typeof schema.mark>[];
  // The fully merged run formatting (inherited < character style < direct),
  // returned so callers can forward it to `convertRunContent` for attributes
  // that are read off the formatting rather than the marks (e.g. footnote
  // anchor vertAlign superscript).
  mergedFormatting: TextFormatting | undefined;
};

const RUN_FORMATTING_INFERENCE = {
  bold: "ordinary-toggle",
  boldCs: "structural",
  italic: "ordinary-toggle",
  italicCs: "structural",
  underline: "underline",
  strike: "visible-boolean",
  doubleStrike: "double-strike",
  vertAlign: "visible-value",
  smallCaps: "visible-boolean",
  allCaps: "visible-boolean",
  hidden: "visible-boolean",
  noProof: "structural",
  color: "color",
  highlight: "visible-value",
  shading: "preserve",
  fontSize: "font-size",
  fontSizeCs: "structural",
  fontFamily: "font-family",
  language: "nested-visible-value",
  spacing: "visible-value",
  position: "visible-value",
  scale: "scale",
  kerning: "visible-value",
  effect: "visible-value",
  emphasisMark: "visible-value",
  emboss: "visible-boolean",
  imprint: "visible-boolean",
  outline: "visible-boolean",
  shadow: "visible-boolean",
  rtl: "visible-boolean",
  cs: "structural",
  styleId: "structural",
} as const satisfies Record<
  keyof typeof RUN_FORMATTING_PROPERTY_SPECS,
  | "color"
  | "double-strike"
  | "font-family"
  | "font-size"
  | "nested-visible-value"
  | "ordinary-toggle"
  | "preserve"
  | "scale"
  | "structural"
  | "underline"
  | "visible-boolean"
  | "visible-value"
>;

const sameFormattingValue = (left: unknown, right: unknown): boolean => {
  if (left === right) {
    return true;
  }
  if (typeof left !== "object" || left === null || typeof right !== "object" || right === null) {
    return false;
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
      return false;
    }
    return left.every((value, index) => sameFormattingValue(value, right.at(index)));
  }
  const leftEntries = Object.entries(left).filter(([, value]) => value !== undefined);
  const rightEntries = Object.entries(right).filter(([, value]) => value !== undefined);
  if (leftEntries.length !== rightEntries.length) {
    return false;
  }
  return leftEntries.every(
    ([property, value]) =>
      Object.hasOwn(right, property) && sameFormattingValue(value, Reflect.get(right, property)),
  );
};

const isReconstructibleFontFamily = (directValue: unknown, effectiveValue: unknown): boolean => {
  if (
    typeof directValue !== "object" ||
    directValue === null ||
    !sameFormattingValue(directValue, effectiveValue)
  ) {
    return false;
  }
  const properties = Object.keys(directValue).filter(
    (property) => Reflect.get(directValue, property) !== undefined,
  );
  return (
    properties.length === 2 &&
    properties.includes("ascii") &&
    properties.includes("hAnsi") &&
    typeof Reflect.get(directValue, "ascii") === "string" &&
    typeof Reflect.get(directValue, "hAnsi") === "string"
  );
};

type CanReconstructAuthoredRunFormattingOptions = {
  directFormatting: TextFormatting | undefined;
  effectiveFormatting: TextFormatting | undefined;
  inheritedFormatting: TextFormatting | undefined;
  paragraphMarkOverrides: TextFormatting | undefined;
};

const canReconstructAuthoredRunFormatting = ({
  directFormatting,
  effectiveFormatting,
  inheritedFormatting,
  paragraphMarkOverrides,
}: CanReconstructAuthoredRunFormattingOptions): boolean => {
  if (paragraphMarkOverrides !== undefined) {
    return false;
  }
  for (const property of Object.keys(directFormatting ?? {}) as (keyof TextFormatting)[]) {
    const directValue = directFormatting?.[property];
    if (directValue === undefined || property === "styleId") {
      continue;
    }
    const inheritedValue = inheritedFormatting?.[property];
    const effectiveValue = effectiveFormatting?.[property];
    switch (RUN_FORMATTING_INFERENCE[property]) {
      case "structural":
        continue;
      case "ordinary-toggle":
        if (directValue !== true || inheritedValue !== true) {
          continue;
        }
        return false;
      case "visible-boolean":
        if (directValue === false || (directValue === true && inheritedValue !== true)) {
          continue;
        }
        return false;
      case "font-size":
        if (typeof directValue === "number" && Number.isFinite(directValue) && directValue > 0) {
          continue;
        }
        return false;
      case "font-family":
        if (isReconstructibleFontFamily(directValue, effectiveValue)) {
          continue;
        }
        return false;
      case "color":
        if (
          typeof directValue === "object" &&
          directValue !== null &&
          Reflect.get(directValue, "auto") !== true &&
          Reflect.get(directValue, "themeColor") === undefined &&
          sameFormattingValue(directValue, effectiveValue)
        ) {
          continue;
        }
        return false;
      case "underline":
        if (
          typeof directValue === "object" &&
          directValue !== null &&
          Reflect.get(directValue, "style") !== "none" &&
          Reflect.get(directValue, "color") === undefined &&
          sameFormattingValue(directValue, effectiveValue) &&
          !sameFormattingValue(directValue, inheritedValue)
        ) {
          continue;
        }
        return false;
      case "double-strike":
        if (directValue === false) {
          continue;
        }
        if (
          directValue === true &&
          inheritedFormatting?.strike !== true &&
          inheritedFormatting?.doubleStrike !== true
        ) {
          continue;
        }
        return false;
      case "scale":
        if (
          normalizeHorizontalScalePercent(
            typeof directValue === "number" ? directValue : undefined,
          ) === directValue &&
          directValue !== inheritedValue
        ) {
          continue;
        }
        return false;
      case "nested-visible-value":
        if (
          sameFormattingValue(directValue, effectiveValue) &&
          !sameFormattingValue(directValue, inheritedValue)
        ) {
          continue;
        }
        return false;
      case "visible-value":
        if (
          directValue !== "none" &&
          directValue !== "baseline" &&
          sameFormattingValue(directValue, effectiveValue) &&
          !sameFormattingValue(directValue, inheritedValue)
        ) {
          continue;
        }
        return false;
      case "preserve":
        return false;
    }
  }
  return true;
};

function buildRunMarks(
  runFormatting: TextFormatting | undefined,
  inherited: ResolvedRunFormatting,
  styleResolver: StyleEngine | null | undefined,
): BuiltRunMarks {
  if (
    runFormatting === undefined &&
    inherited.implicitCharacterStyleApplied === true &&
    inherited.paragraphMarkOverrides === undefined
  ) {
    return {
      marks: textFormattingToMarks(inherited.formatting, schema, {
        overrideFormatting: undefined,
        directFormatting: undefined,
        authoredCarrier: "reconstruct",
      }),
      mergedFormatting: inherited.formatting,
    };
  }
  const styleId = runFormatting?.styleId;
  let characterStyleFormatting: TextFormatting | undefined;
  if (styleId) {
    characterStyleFormatting = styleResolver?.getRunStyleOwnProperties(styleId);
  } else if (!inherited.implicitCharacterStyleApplied) {
    characterStyleFormatting = styleResolver?.getDefaultCharacterStyle()?.rPr;
  }
  let ordinaryRunStyleFormatting = inherited.formatting ? { ...inherited.formatting } : {};
  if (styleId) {
    ordinaryRunStyleFormatting =
      mergeTextFormatting(inherited.formatting, characterStyleFormatting) ?? {};
  }
  const cascadedStyleFormatting = cascadeStyleTextFormatting(
    [
      { cascade: inherited.toggleCascade, type: "carried" },
      { formatting: characterStyleFormatting, type: "style" },
    ],
    { ordinaryFormatting: ordinaryRunStyleFormatting },
  );
  const runStyleFormatting = cascadedStyleFormatting.formatting;
  const finalToggleFormatting = cascadeStyleTextFormatting(
    [
      { cascade: cascadedStyleFormatting, type: "carried" },
      { formatting: runFormatting, type: "direct" },
    ],
    { ordinaryFormatting: mergeTextFormatting(runStyleFormatting, runFormatting) },
  );
  const mergedFormatting = finalToggleFormatting.formatting;
  const authoredCarrier: AuthoredRunFormattingCarrier = canReconstructAuthoredRunFormatting({
    directFormatting: runFormatting,
    effectiveFormatting: mergedFormatting,
    inheritedFormatting: runStyleFormatting,
    paragraphMarkOverrides: inherited.paragraphMarkOverrides,
  })
    ? "reconstruct"
    : "preserve";
  const overrideFormatting = getRunFormattingOverrides({
    directFormatting: runFormatting,
    effectiveStyleFormatting: runStyleFormatting,
    hasCharacterStyle: styleId !== undefined,
    paragraphMarkOverrides: inherited.paragraphMarkOverrides,
  });
  if (authoredCarrier === "reconstruct") {
    if (overrideFormatting?.bold === true) {
      delete overrideFormatting.bold;
    }
    if (overrideFormatting?.italic === true) {
      delete overrideFormatting.italic;
    }
  }
  const marks = textFormattingToMarks(mergedFormatting, schema, {
    overrideFormatting,
    directFormatting: runFormatting,
    authoredCarrier,
  });

  if (styleId) {
    marks.push(schema.mark("characterStyle", { styleId }));
  }

  return { marks, mergedFormatting };
}

const ORDINARY_STYLE_TOGGLE_KEYS = [
  "strike",
  "allCaps",
  "smallCaps",
  "hidden",
  "emboss",
  "imprint",
  "shadow",
  "outline",
] as const satisfies readonly (keyof TextFormatting)[];

type GetRunFormattingOverridesOptions = {
  directFormatting: TextFormatting | undefined;
  effectiveStyleFormatting: TextFormatting | undefined;
  hasCharacterStyle: boolean;
  paragraphMarkOverrides: TextFormatting | undefined;
};

function getRunFormattingOverrides({
  directFormatting,
  effectiveStyleFormatting,
  hasCharacterStyle,
  paragraphMarkOverrides,
}: GetRunFormattingOverridesOptions): TextFormatting | undefined {
  const overrides = mergeTextFormatting(paragraphMarkOverrides, directFormatting);
  if (!overrides || !directFormatting) {
    return overrides;
  }

  // A positive PM mark already preserves direct formatting unless character-style
  // subtraction would mistake it for an inherited visual. Keep only that ambiguous
  // positive state in the structural override; negative state remains explicit.
  // Bold and italic stay in the override too: their presence distinguishes a
  // direct ordinary toggle from an inherited mark before deciding whether its
  // complex-script partner was authored.
  for (const key of ORDINARY_STYLE_TOGGLE_KEYS) {
    if (
      directFormatting[key] === true &&
      (!hasCharacterStyle || effectiveStyleFormatting?.[key] !== true)
    ) {
      Reflect.deleteProperty(overrides, key);
    }
  }

  return overrides;
}

/**
 * The vertical alignment a note reference mark carries.
 *
 * The mark models the two alignments a reference can sit at; `subscript` is not
 * one of them and drops to `null`. `baseline` is read through its owner, since
 * it is the reserved "no offset" value rather than an offset of its own.
 */
const noteReferenceVertAlign = (
  vertAlign: TextFormatting["vertAlign"],
): "baseline" | "superscript" | null => {
  if (isBaselineVertAlign(vertAlign)) {
    return "baseline";
  }
  return vertAlign === "superscript" ? "superscript" : null;
};

/**
 * The run properties an inline atom has to carry itself.
 *
 * `withRunBoundaryMarks` keeps the run's formatting marks off an image or
 * shape node, so the node is the only place left to record the `w:rPr` the
 * run was authored with.
 */
const carriedRunFormatting = (
  formatting: TextFormatting | undefined,
): TextFormatting | undefined =>
  formatting && Object.keys(formatting).length > 0 ? formatting : undefined;

/**
 * Convert RunContent to ProseMirror nodes
 */
function convertRunContent(
  content: RunContent,
  marks: ReturnType<typeof schema.mark>[],
  formatting?: TextFormatting,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
  /**
   * The run's own `w:rPr`, NOT the style-resolved `formatting` above. An
   * inline atom carries it verbatim so the save can rebuild the run; carrying
   * the resolved value instead would write the style's run properties into
   * the run as direct formatting.
   */
  authoredFormatting?: TextFormatting,
): PMNode[] {
  switch (content.type) {
    case "text":
      if (content.text) {
        return [schema.text(content.text, marks)];
      }
      return [];

    case "break":
      if (content.breakType === "textWrapping" || !content.breakType) {
        const attrs = {
          ...(content.breakType !== undefined ? { breakType: content.breakType } : {}),
          ...(content.clear !== undefined ? { clear: content.clear } : {}),
        };
        return [schema.node("hardBreak", attrs).mark(marks)];
      }
      if (content.breakType === "column") {
        return [
          schema
            .node("hardBreak", {
              breakType: "column",
              ...(content.clear !== undefined ? { clear: content.clear } : {}),
            })
            .mark(marks),
        ];
      }
      return [
        schema
          .node("pageBreakRun", content.clear === undefined ? undefined : { clear: content.clear })
          .mark(marks),
      ];

    case "renderedPageBreak":
      return [schema.node("renderedPageBreak").mark(marks)];

    case "tab":
      // Convert to tab node for proper rendering. Keep the run marks because
      // Word commonly represents signature blanks as underlined tab runs.
      return [
        schema
          .node("tab", content.positional ? { positional: content.positional } : undefined)
          .mark(marks),
      ];

    case "drawing":
      return [
        withRunBoundaryMarks(
          convertImage({
            image: content.image,
            rawXml: content.rawXml,
            rawXmlMode: content.rawXmlMode,
            rawImageFingerprint:
              content.rawXmlMode === DRAWING_RAW_XML_MODES.PRESERVE_ONLY
                ? undefined
                : content.rawImageFingerprint,
            runFormatting: carriedRunFormatting(authoredFormatting),
          }),
          marks,
        ),
      ];

    case "shape": {
      // Shapes with text body are handled as text boxes at block level
      // Other shapes render as inline SVG
      const shp = content.shape;
      if (shp.textBody) {
        const anchorId = textBoxAnchors?.get(shp);
        return anchorId ? [schema.node("textBoxAnchor", { anchorId }).mark(marks)] : [];
      }
      return [
        withRunBoundaryMarks(convertShape(shp, carriedRunFormatting(authoredFormatting)), marks),
      ];
    }

    case "footnoteRef": {
      // Footnote reference - render as superscript number with footnoteRef mark
      const footnoteMark = schema.mark("footnoteRef", {
        id: content.id.toString(),
        noteType: "footnote",
        vertAlign: noteReferenceVertAlign(formatting?.vertAlign),
      });
      return [schema.text(content.id.toString(), [...marks, footnoteMark])];
    }

    case "endnoteRef": {
      // Endnote reference - render as superscript number with footnoteRef mark
      const endnoteMark = schema.mark("footnoteRef", {
        id: content.id.toString(),
        noteType: "endnote",
        vertAlign: noteReferenceVertAlign(formatting?.vertAlign),
      });
      return [schema.text(content.id.toString(), [...marks, endnoteMark])];
    }

    case "fieldChar":
    case "instrText":
      // Complex field structure markers — handled at the run/paragraph
      // level via `convertField`, not as standalone inline content.
      return [];

    // Opaque: the editor cannot edit markup it has no model for, and only has
    // to carry it. The visible text rides along so a `w:ruby` base still reads.
    case "preservedXml":
      return [
        schema
          .node("preservedXml", {
            xml: content.xml,
            text: content.text,
            level: PRESERVED_XML_LEVELS.run,
          })
          .mark(marks),
      ];

    case "noBreakHyphen":
      return [schema.text("‑", marks)];

    case "softHyphen":
      return [schema.text("­", marks)];

    case "symbol":
      return [schema.node("symbol", { font: content.font, char: content.char }).mark(marks)];
  }
}

function withRunBoundaryMarks(node: PMNode, marks: ReturnType<typeof schema.mark>[]): PMNode {
  const ownsWrapperOrSourceRun = marks.some(
    ({ type }) => type.name === "hyperlink" || type.name === "pageBreakRunOwner",
  );
  if (!ownsWrapperOrSourceRun) {
    return node;
  }

  return node.mark(marks);
}

/**
 * Convert an Image to a ProseMirror image node
 *
 * DOCX images have size in EMUs (English Metric Units), which must be
 * converted to pixels for proper HTML rendering.
 * 914400 EMU = 1 inch = 96 CSS pixels
 *
 * Image types in DOCX:
 * 1. Inline (wp:inline) - flows with text like a character
 * 2. Floating/Anchored (wp:anchor) with wrap types:
 *    - Square/Tight/Through: text wraps around image
 *      - wrapText='left' → text on LEFT, image floats RIGHT
 *      - wrapText='right' → text on RIGHT, image floats LEFT
 *      - wrapText='bothSides' → depends on horizontal alignment
 *    - TopAndBottom: image on its own line, text above/below only
 *    - None/Behind/InFront: positioned image, no text wrap
 */
type PartialImagePosition = Partial<NonNullable<Image["position"]>>;
type PartialImageSize = Partial<Image["size"]>;

type ConvertImageOptions = {
  image: Image;
  rawXml: DrawingContent["rawXml"];
  rawXmlMode: DrawingContent["rawXmlMode"];
  rawImageFingerprint: string | undefined;
  /** The `w:rPr` of the run the drawing came from; see `carriedRunFormatting`. */
  runFormatting: TextFormatting | undefined;
};

/**
 * The authored-EMU carrier for a drawing, or nothing when the source authored
 * none of these values and there is accordingly nothing to carry.
 */
const authoredEmuAttrs = <Values extends Record<string, number | undefined>>(
  values: Values,
): Values | undefined =>
  Object.values(values).some((value) => value !== undefined) ? values : undefined;

/**
 * A copy of the anchor record: ProseMirror keeps object-valued attrs by
 * reference, so sharing it with the source would let a mutation of either reach
 * the other outside a transaction.
 */
const copiedDrawingAnchor = (anchor: DrawingAnchor | undefined): DrawingAnchor | undefined =>
  anchor === undefined
    ? undefined
    : {
        ...anchor,
        ...(anchor.simplePosition === undefined
          ? {}
          : { simplePosition: { ...anchor.simplePosition } }),
      };

/**
 * A copy of the inset-slot record, for the reason {@link copiedDrawingAnchor}
 * is copied: ProseMirror keeps object-valued attrs by reference.
 */
const copiedWrapDistanceSlots = (
  slots: WrapDistanceSlots | undefined,
): WrapDistanceSlots | undefined =>
  slots === undefined
    ? undefined
    : {
        ...(slots.drawing === undefined ? {} : { drawing: { ...slots.drawing } }),
        ...(slots.wrapChild === undefined ? {} : { wrapChild: { ...slots.wrapChild } }),
      };

/** The wrap insets of a drawing, keyed by the pixel attribute each becomes. */
const wrapDistanceEmu = (
  wrap: Image["wrap"] | undefined,
): Record<"distTop" | "distBottom" | "distLeft" | "distRight", number | undefined> => ({
  distTop: wrap?.distT,
  distBottom: wrap?.distB,
  distLeft: wrap?.distL,
  distRight: wrap?.distR,
});

function convertImage({
  image,
  rawXml,
  rawXmlMode,
  rawImageFingerprint,
  runFormatting,
}: ConvertImageOptions): PMNode {
  // Convert EMU to pixels for proper sizing
  const imageData: { size?: PartialImageSize } = image;
  const imageSize = imageData.size;
  const widthPx = imageSize?.width ? emuToPixels(imageSize.width) : undefined;
  const heightPx = imageSize?.height ? emuToPixels(imageSize.height) : undefined;

  // Determine wrap type and float direction
  const wrapType = image.wrap.type;
  const wrapText = image.wrap.wrapText;
  const imagePosition: PartialImagePosition | undefined = image.position;
  const hAlign = imagePosition?.horizontal?.alignment;

  // Determine CSS float based on wrap settings
  // In DOCX: wrapText='left' means "text flows on the left" → image is on right → float: right
  //          wrapText='right' means "text flows on the right" → image is on left → float: left
  let cssFloat: "left" | "right" | "none" | undefined;

  if (wrapType === "inline") {
    cssFloat = "none"; // Inline images don't float
  } else if (wrapType === "topAndBottom") {
    cssFloat = "none"; // Block images don't float
  } else if (wrapType === "square" || wrapType === "tight" || wrapType === "through") {
    // These wrap types support text wrapping around the image
    if (wrapText === "left") {
      cssFloat = "right"; // Text on left → image floats right
    } else if (wrapText === "right") {
      cssFloat = "left"; // Text on right → image floats left
    } else {
      // bothSides, largest, or any other wrapText value:
      // use horizontal alignment to determine float
      if (hAlign === "left") {
        cssFloat = "left";
      } else if (hAlign === "right") {
        cssFloat = "right";
      } else {
        cssFloat = "none"; // Center or no alignment → block
      }
    }
  } else {
    // Behind, inFront, etc. - positioned images, no float
    cssFloat = "none";
  }

  // Determine display mode for CSS.
  //
  // - inline           → inline run, participates in flow
  // - topAndBottom     → block image, takes its own line
  // - behind / inFront → float (anchored at absolute coords; the page-level
  //   layer paints them, so they must be lifted out of the paragraph flow
  //   even though they don't carve a text-wrap exclusion zone)
  // - square / tight / through with cssFloat → float
  // - everything else (centered etc.) → block
  let displayMode: "inline" | "block" | "float";
  if (wrapType === "inline") {
    displayMode = "inline";
  } else if (wrapType === "topAndBottom") {
    displayMode = "block";
  } else if (wrapType === "behind" || wrapType === "inFront") {
    displayMode = "float";
  } else if (cssFloat !== "none") {
    displayMode = "float";
  } else {
    displayMode = "block";
  }

  const transformAttrs = authoredTransformAttrs(image.transform);

  // Convert wrap distances from EMU to pixels for margins
  const distTop = image.wrap.distT != null ? emuToPixels(image.wrap.distT) : undefined;
  const distBottom = image.wrap.distB != null ? emuToPixels(image.wrap.distB) : undefined;
  const distLeft = image.wrap.distL != null ? emuToPixels(image.wrap.distL) : undefined;
  const distRight = image.wrap.distR != null ? emuToPixels(image.wrap.distR) : undefined;

  // Build position data for floating images
  let position:
    | {
        horizontal?: {
          relativeTo?: string;
          posOffset?: number;
          align?: string;
        };
        vertical?: { relativeTo?: string; posOffset?: number; align?: string };
      }
    | undefined;
  if (imagePosition) {
    position = {};
    if (imagePosition.horizontal) {
      const h: { relativeTo?: string; posOffset?: number; align?: string } = {
        relativeTo: imagePosition.horizontal.relativeTo,
      };
      if (imagePosition.horizontal.posOffset !== undefined) {
        h.posOffset = imagePosition.horizontal.posOffset;
      }
      if (imagePosition.horizontal.alignment) {
        h.align = imagePosition.horizontal.alignment;
      }
      position.horizontal = h;
    }

    if (imagePosition.vertical) {
      const v: { relativeTo?: string; posOffset?: number; align?: string } = {
        relativeTo: imagePosition.vertical.relativeTo,
      };
      if (imagePosition.vertical.posOffset !== undefined) {
        v.posOffset = imagePosition.vertical.posOffset;
      }
      if (imagePosition.vertical.alignment) {
        v.align = imagePosition.vertical.alignment;
      }
      position.vertical = v;
    }
  }

  // Convert outline to border attrs
  let borderWidth: number | undefined;
  let borderColor: string | undefined;
  let borderStyle: string | undefined;
  if (image.outline && image.outline.width) {
    borderWidth = emuToStrokePixels(image.outline.width);
    if (image.outline.color?.rgb) {
      borderColor = `#${image.outline.color.rgb}`;
    }
    // Map OOXML dash styles to CSS border styles
    const styleMap: Record<string, string> = {
      solid: "solid",
      dot: "dotted",
      dash: "dashed",
      lgDash: "dashed",
      dashDot: "dashed",
      lgDashDot: "dashed",
      lgDashDotDot: "dashed",
      sysDot: "dotted",
      sysDash: "dashed",
      sysDashDot: "dashed",
      sysDashDotDot: "dashed",
    };
    borderStyle = image.outline.style ? styleMap[image.outline.style] || "solid" : "solid";
  }

  return schema.node("image", {
    src: image.src || "",
    // The descriptor travels, the raster does not: reading this attr never
    // builds one, so projecting a document into the editor costs nothing per
    // diagram beyond the shapes the parse already read.
    preview: image.preview,
    docPrName: image.docPrName,
    alt: image.alt,
    title: image.title,
    width: widthPx,
    height: heightPx,
    rId: image.rId,
    wrapType,
    displayMode,
    cssFloat,
    ...transformAttrs,
    // eigenpal #424 (opacity render pipeline). PR #513 added Image.opacity
    // on the model; thread it onto the PM node so the layout-bridge and
    // painter can honor it.
    opacity: image.opacity,
    brightness: image.effects?.brightness,
    contrast: image.effects?.contrast,
    distTop,
    distBottom,
    distLeft,
    distRight,
    // eigenpal #424: thread wp:srcRect crop fractions through PM attrs.
    cropTop: image.crop?.top,
    cropRight: image.crop?.right,
    cropBottom: image.crop?.bottom,
    cropLeft: image.crop?.left,
    // wp:effectExtent stays in EMU: nothing renders it and emuToPixels rounds.
    paddingTop: image.padding?.top,
    paddingRight: image.padding?.right,
    paddingBottom: image.padding?.bottom,
    paddingLeft: image.padding?.left,
    position,
    anchor: copiedDrawingAnchor(image.anchor),
    wrapDistanceSlots: copiedWrapDistanceSlots(image.wrap.distanceSlots),
    wrapPolygon: copiedWrapPolygon(image.wrap.polygon),
    // Two facts, carried separately: a decorative image is displayed and
    // skipped by assistive technology, a hidden one is not displayed.
    decorative: image.decorative,
    hidden: image.hidden,
    // Copy: ProseMirror keeps array-valued attrs by reference.
    docPrExtensions: image.docPrExtensions ? [...image.docPrExtensions] : undefined,
    // Copy: ProseMirror keeps object-valued attrs by reference, so sharing this
    // with the source Image would let a mutation of either reach the other
    // outside a transaction. `position` below is already built fresh.
    frameLocks: image.frameLocks ? { ...image.frameLocks } : undefined,
    borderWidth,
    // The pixel values above are lossy; `fromProseDoc` writes these back while
    // the pixels still project from them.
    _docxAuthoredEmu: authoredEmuAttrs({
      width: imageSize?.width,
      height: imageSize?.height,
      borderWidth: image.outline?.width,
      ...wrapDistanceEmu(image.wrap),
    }),
    borderColor,
    borderStyle,
    wrapText,
    hlinkHref: image.hlinkHref,
    hlinkRId: image.hlinkRId,
    // Copy: ProseMirror keeps object-valued attrs by reference.
    hlinkClickSource: image.hlinkClickSource ? { ...image.hlinkClickSource } : undefined,
    hlinkHoverXml: image.hlinkHoverXml,
    docPrId: image.id,
    _docxRawXml: rawXml,
    _docxRawXmlMode: rawXmlMode,
    _docxRawImageFingerprint: rawImageFingerprint,
    _docxObjectPreview:
      rawXml !== undefined && /<(?:[A-Za-z_][\w.-]*:)?object(?:\s|>)/u.test(rawXml),
    _docxRunFormatting: runFormatting,
  });
}

/**
 * Convert a Hyperlink to ProseMirror nodes with link mark
 *
 * @param hyperlink - The hyperlink to convert
 * @param options - Formatting, source identity, and extracted text-box anchors
 */
type ConvertHyperlinkOptions = {
  getInheritedRunFormatting: RunFormattingResolver;
  styleResolver: StyleEngine | null | undefined;
  hyperlinkIndex: number;
  textBoxAnchors: ReadonlyMap<Shape, string> | undefined;
  nextPageBreakRunOwnerId: PageBreakRunOwnerIdAllocator;
};

function convertHyperlink(
  hyperlink: Hyperlink,
  {
    getInheritedRunFormatting,
    styleResolver,
    hyperlinkIndex,
    textBoxAnchors,
    nextPageBreakRunOwnerId,
  }: ConvertHyperlinkOptions,
): PMNode[] {
  const nodes: PMNode[] = [];

  // Create link mark — internal anchors use #bookmarkName format
  const href = hyperlink.href || (hyperlink.anchor ? `#${hyperlink.anchor}` : "");
  const linkMark = schema.mark("hyperlink", {
    href,
    tooltip: hyperlink.tooltip,
    rId: hyperlink.rId,
    _docxHyperlinkIndex: hyperlinkIndex,
  });

  for (const child of hyperlink.children) {
    if (child.type === "bookmarkStart") {
      nodes.push(
        schema.node(
          "bookmarkBoundary",
          {
            type: "start",
            id: child.id,
            name: child.name,
            colFirst: child.colFirst,
            colLast: child.colLast,
            displacedByCustomXml: child.displacedByCustomXml,
          },
          undefined,
          [linkMark],
        ),
      );
      continue;
    }
    if (child.type === "bookmarkEnd") {
      nodes.push(
        schema.node(
          "bookmarkBoundary",
          {
            type: "end",
            id: child.id,
            displacedByCustomXml: child.displacedByCustomXml,
          },
          undefined,
          [linkMark],
        ),
      );
      continue;
    }
    if (child.type === "preservedInline") {
      // The same opaque atom the paragraph level uses, carrying the link
      // mark: markup authored inside a `w:hyperlink` is accepted, rejected
      // and moved with the link rather than beside it.
      nodes.push(preservedInlineNode(child).mark([linkMark]));
      continue;
    }
    if (child.type === "run") {
      // Merge style formatting with run's inline formatting
      const inheritedFormatting = getInheritedRunFormatting(child.formatting);
      const { marks: runMarks, mergedFormatting } = buildRunMarks(
        child.formatting,
        inheritedFormatting,
        styleResolver,
      );
      if (child.propertyChanges && child.propertyChanges.length > 0) {
        runMarks.push(schema.mark("runPropertyChange", { changes: [...child.propertyChanges] }));
      }
      if (
        child.content.some((content) => content.type === "break" && content.breakType === "page")
      ) {
        runMarks.push(schema.mark("pageBreakRunOwner", { id: nextPageBreakRunOwnerId() }));
      }
      // Add link mark to run marks
      const allMarks = [...runMarks, linkMark];

      // Delegate to convertRunContent so tabs/breaks/fields/symbols inside
      // a hyperlink round-trip (eigenpal #566). The earlier text-only loop
      // silently dropped TOC entries' tab between title and page number,
      // collapsing the right-aligned page number flush against the title.
      for (const content of child.content) {
        nodes.push(
          ...convertRunContent(
            content,
            allMarks,
            mergedFormatting,
            textBoxAnchors,
            child.formatting,
          ),
        );
      }
    }
  }

  return nodes;
}

// ============================================================================
// SHAPE CONVERSION
// ============================================================================

/**
 * Convert a Shape to a ProseMirror shape node (inline SVG)
 */
function convertShape(shape: Shape, runFormatting?: TextFormatting): PMNode {
  const shapeData: { size?: Partial<Shape["size"]> } = shape;
  const shapeSize = shapeData.size;
  const widthPx = shapeSize?.width ? emuToPixels(shapeSize.width) : 100;
  const heightPx = shapeSize?.height ? emuToPixels(shapeSize.height) : 80;
  const shapeAttrs: { shapeType?: Shape["shapeType"] } = shape;

  let fillColor: string | undefined;
  let fillColorValue: NonNullable<Shape["fill"]>["color"] | undefined;
  let fillType: string = "solid";
  let gradientType: string | undefined;
  let gradientAngle: number | undefined;
  let gradientStops: string | undefined;
  if (shape.fill) {
    fillType = shape.fill.type;
    if (shape.fill.color) {
      fillColorValue = shape.fill.color;
      fillColor = resolveColorValueToHex(shape.fill.color);
    }
    // Extract gradient data
    if (shape.fill.type === "gradient" && shape.fill.gradient) {
      const g = shape.fill.gradient;
      gradientType = g.type;
      gradientAngle = g.angle;
      // Convert stops to serializable format with CSS colors
      gradientStops = JSON.stringify(
        g.stops.map((s) => ({
          position: s.position,
          color: s.color.rgb ? `#${s.color.rgb}` : "#000000",
        })),
      );
    }
  }

  let outlineWidth: number | undefined;
  let outlineColor: string | undefined;
  let outlineColorValue: NonNullable<Shape["outline"]>["color"] | undefined;
  let outlineStyle: string | undefined = "none";
  let outlineCap: NonNullable<Shape["outline"]>["cap"] | undefined;
  let outlineJoin: NonNullable<Shape["outline"]>["join"] | undefined;
  let outlineHeadEnd: NonNullable<Shape["outline"]>["headEnd"] | undefined;
  let outlineTailEnd: NonNullable<Shape["outline"]>["tailEnd"] | undefined;
  if (shape.outline) {
    if (shape.outline.width) {
      outlineWidth = emuToStrokePixels(shape.outline.width);
    }
    if (shape.outline.color) {
      outlineColorValue = shape.outline.color;
      outlineColor = resolveColorValueToHex(shape.outline.color);
    }
    outlineStyle = shape.outline.style;
    outlineCap = shape.outline.cap;
    outlineJoin = shape.outline.join;
    outlineHeadEnd = shape.outline.headEnd;
    outlineTailEnd = shape.outline.tailEnd;
  } else {
    outlineWidth = 0;
  }

  const transformAttrs = authoredTransformAttrs(shape.transform);

  const wrapType = shape.wrap?.type ?? "inline";
  const displayMode = wrapType === "inline" ? "inline" : "float";
  let cssFloat: "left" | "right" | "none" = "none";
  if (shape.wrap?.wrapText === "left") {
    cssFloat = "right";
  } else if (shape.wrap?.wrapText === "right") {
    cssFloat = "left";
  }

  let position: ImagePositionAttrs | undefined;
  if (shape.position) {
    position = {
      horizontal: {
        relativeTo: shape.position.horizontal.relativeTo,
        ...(shape.position.horizontal.posOffset !== undefined
          ? { posOffset: shape.position.horizontal.posOffset }
          : {}),
        ...(shape.position.horizontal.alignment
          ? { align: shape.position.horizontal.alignment }
          : {}),
      },
      vertical: {
        relativeTo: shape.position.vertical.relativeTo,
        ...(shape.position.vertical.posOffset !== undefined
          ? { posOffset: shape.position.vertical.posOffset }
          : {}),
        ...(shape.position.vertical.alignment ? { align: shape.position.vertical.alignment } : {}),
      },
    };
  }

  return schema.node("shape", {
    _docxRunFormatting: runFormatting,
    shapeType: shapeAttrs.shapeType ?? "rect",
    geometryAdjustments:
      shape.geometryAdjustments === undefined
        ? undefined
        : JSON.stringify(shape.geometryAdjustments),
    shapeId: shape.id,
    shapeName: shape.name,
    anchor: copiedDrawingAnchor(shape.anchor),
    wrapDistanceSlots: copiedWrapDistanceSlots(shape.wrap?.distanceSlots),
    wrapPolygon: copiedWrapPolygon(shape.wrap?.polygon),
    alt: shape.alt,
    title: shape.title,
    width: widthPx,
    height: heightPx,
    // The pixels above and the wrap insets below are lossy; `fromProseDoc`
    // writes these back while the pixels still project from them.
    _docxAuthoredEmu: authoredEmuAttrs({
      width: shapeSize?.width,
      height: shapeSize?.height,
      outlineWidth: shape.outline?.width,
      ...wrapDistanceEmu(shape.wrap),
    }),
    fillColor,
    fillColorValue,
    fillType,
    gradientType,
    gradientAngle,
    gradientStops,
    outlineWidth,
    outlineColor,
    outlineColorValue,
    outlineStyle,
    outlineCap,
    outlineJoin,
    outlineHeadEnd,
    outlineTailEnd,
    ...transformAttrs,
    displayMode,
    cssFloat,
    wrapType,
    wrapText: shape.wrap?.wrapText,
    distTop: shape.wrap?.distT !== undefined ? emuToPixels(shape.wrap.distT) : undefined,
    distBottom: shape.wrap?.distB !== undefined ? emuToPixels(shape.wrap.distB) : undefined,
    distLeft: shape.wrap?.distL !== undefined ? emuToPixels(shape.wrap.distL) : undefined,
    distRight: shape.wrap?.distR !== undefined ? emuToPixels(shape.wrap.distR) : undefined,
    position,
  });
}

// ============================================================================
// TEXT BOX CONVERSION
// ============================================================================

/**
 * Convert a paragraph block to PM nodes, extracting text boxes as sibling nodes.
 * Skips ghost empty paragraphs that only contained text box drawings.
 */
type ConvertParagraphWithTextBoxesOptions = {
  textBoxGroupId: string;
  context: TableConversionContext;
  extraRunFormatting?: TextFormatting;
  preserveEmptyWrapper?: boolean;
  tableParagraphOverlay?: TableCellParagraphSpacingOverlay;
};

function convertParagraphWithTextBoxes(
  block: Paragraph,
  styleResolver: StyleEngine | null,
  {
    textBoxGroupId,
    context,
    extraRunFormatting,
    preserveEmptyWrapper,
    tableParagraphOverlay,
  }: ConvertParagraphWithTextBoxesOptions,
): PMNode[] {
  const { textBoxes, textBoxAnchors } = extractTextBoxesFromParagraph(block, textBoxGroupId);
  const pmParagraph = convertParagraph(
    block,
    styleResolver,
    context,
    extraRunFormatting,
    tableParagraphOverlay,
    textBoxAnchors,
  );
  const nodes: PMNode[] = [];
  const isEmptyAfterExtraction =
    textBoxes.length > 0 && !hasContentBesidesTextBoxAnchors(pmParagraph);
  const keepWrapperParagraph =
    isEmptyAfterExtraction &&
    (preserveEmptyWrapper === true || hasParagraphBoundaryPayload(block, pmParagraph));
  if (!isEmptyAfterExtraction || keepWrapperParagraph) {
    nodes.push(pmParagraph);
  }
  const standalone = isEmptyAfterExtraction && !keepWrapperParagraph;
  for (const [index, { textBox, anchorId, trackedChange, inlineSdts }] of textBoxes.entries()) {
    nodes.push(
      convertTextBox(textBox, styleResolver, {
        placement: standalone ? "standalone" : "inlineWithPrevious",
        groupId: textBoxGroupId,
        anchorId,
        context,
        trackedChange,
        inlineSdts,
        // The host paragraph is gone from the projection, so the first node of
        // the group speaks for it: `fromProseDoc` rebuilds one paragraph for a
        // group and puts the remainder back on it. Only the first, or a group
        // of three boxes would claim the same authored attributes three times.
        ...(standalone && index === 0 && block.preservedAttributes
          ? { hostPreservedAttributes: block.preservedAttributes }
          : {}),
      }),
    );
  }
  return nodes;
}

function hasContentBesidesTextBoxAnchors(paragraph: PMNode): boolean {
  let hasContent = false;
  paragraph.descendants((node) => {
    if (node.type.name === "textBoxAnchor") {
      return false;
    }
    if (node.type.name === "sdt") {
      if (node.childCount === 0) {
        hasContent = true;
      }
      return !hasContent;
    }
    hasContent = true;
    return false;
  });
  return hasContent;
}

function hasParagraphBoundaryPayload(block: Paragraph, pmParagraph: PMNode): boolean {
  const bookmarks = pmParagraph.attrs["bookmarks"];
  const emptyHyperlinks = pmParagraph.attrs["_emptyHyperlinks"];
  return Boolean(
    block.sectionProperties ||
    block.propertyChanges?.length ||
    (Array.isArray(bookmarks) && bookmarks.length > 0) ||
    (Array.isArray(emptyHyperlinks) && emptyHyperlinks.length > 0),
  );
}

/**
 * Extract text boxes from paragraph runs.
 * Text boxes appear as ShapeContent where the shape has textBody.
 */
type ExtractedTextBox = {
  textBox: TextBox;
  anchorId: string;
  trackedChange: NonNullable<TextBoxAttrs["_docxTrackedChange"]> | undefined;
  inlineSdts: NonNullable<TextBoxAttrs["_docxInlineSdts"]>;
};

type TextBoxExtractionContext = {
  trackedChange?: NonNullable<TextBoxAttrs["_docxTrackedChange"]>;
  inlineSdts: NonNullable<TextBoxAttrs["_docxInlineSdts"]>;
};

type ExtractTextBoxesResult = {
  textBoxes: readonly ExtractedTextBox[];
  textBoxAnchors: ReadonlyMap<Shape, string>;
};

const NO_EXTRACTED_TEXT_BOXES: ExtractTextBoxesResult = {
  textBoxes: [],
  textBoxAnchors: new Map(),
};

function extractTextBoxesFromParagraph(
  paragraph: Paragraph,
  textBoxGroupId: string,
): ExtractTextBoxesResult {
  if (!mayContainTextBox(paragraph)) {
    return NO_EXTRACTED_TEXT_BOXES;
  }

  return extractTextBoxes(paragraph, textBoxGroupId);
}

function mayContainTextBox(paragraph: Paragraph): boolean {
  for (const item of paragraph.content) {
    if (item.type !== "run") {
      if (
        item.type === "inlineSdt" ||
        item.type === "hyperlink" ||
        item.type === "insertion" ||
        item.type === "deletion" ||
        item.type === "moveFrom" ||
        item.type === "moveTo"
      ) {
        return true;
      }
      continue;
    }

    for (const content of item.content) {
      if (content.type === "shape" && content.shape.textBody) {
        return true;
      }
    }
  }
  return false;
}

function extractTextBoxes(paragraph: Paragraph, textBoxGroupId: string): ExtractTextBoxesResult {
  const textBoxes: ExtractedTextBox[] = [];
  const textBoxAnchors = new Map<Shape, string>();
  const visitRun = (run: Run, context: TextBoxExtractionContext): void => {
    for (const runContent of run.content) {
      if (runContent.type !== "shape" || !runContent.shape.textBody) {
        continue;
      }
      const anchorId = `${textBoxGroupId}:${textBoxes.length}`;
      textBoxAnchors.set(runContent.shape, anchorId);
      textBoxes.push({
        textBox: textBoxFromShape(runContent.shape, runContent.shape.textBody),
        anchorId,
        trackedChange: context.trackedChange,
        inlineSdts: context.inlineSdts,
      });
    }
  };

  const visitHyperlink = (hyperlink: Hyperlink, context: TextBoxExtractionContext): void => {
    for (const child of hyperlink.children) {
      if (child.type === "run") {
        visitRun(child, context);
      }
    }
  };

  const visitTrackedChange = (
    change: Insertion | Deletion | MoveFrom | MoveTo,
    context: TextBoxExtractionContext,
  ): void => {
    const trackedChange = { type: change.type, info: change.info } as const satisfies NonNullable<
      TextBoxAttrs["_docxTrackedChange"]
    >;
    for (const item of change.content) {
      if (item.type === "run") {
        visitRun(item, { ...context, trackedChange });
        continue;
      }
      if (item.type === "hyperlink") {
        visitHyperlink(item, { ...context, trackedChange });
      }
    }
  };

  const visitInlineSdt = (sdt: InlineSdt, context: TextBoxExtractionContext): void => {
    const inlineSdts = [...context.inlineSdts, sdtAttrsFromProperties(sdt.properties)];
    const nestedContext = { ...context, inlineSdts };

    for (const item of sdt.content) {
      if (item.type === "run") {
        visitRun(item, nestedContext);
        continue;
      }
      if (item.type === "inlineSdt") {
        visitInlineSdt(item, nestedContext);
        continue;
      }
      if (item.type === "hyperlink") {
        visitHyperlink(item, nestedContext);
        continue;
      }
      if (
        item.type === "insertion" ||
        item.type === "deletion" ||
        item.type === "moveFrom" ||
        item.type === "moveTo"
      ) {
        visitTrackedChange(item, nestedContext);
      }
    }
  };

  const rootContext: TextBoxExtractionContext = { inlineSdts: [] };
  for (const item of paragraph.content) {
    if (item.type === "run") {
      visitRun(item, rootContext);
      continue;
    }
    if (item.type === "inlineSdt") {
      visitInlineSdt(item, rootContext);
      continue;
    }
    if (item.type === "hyperlink") {
      visitHyperlink(item, rootContext);
      continue;
    }
    if (
      item.type === "insertion" ||
      item.type === "deletion" ||
      item.type === "moveFrom" ||
      item.type === "moveTo"
    ) {
      visitTrackedChange(item, rootContext);
    }
  }

  return {
    textBoxes,
    textBoxAnchors,
  };
}

function textBoxFromShape(shape: Shape, textBody: ShapeTextBody): TextBox {
  const textBox: TextBox = {
    type: "textBox",
    size: shape.size,
    content: textBody.content,
  };
  if (shape.id) {
    textBox.id = shape.id;
  }
  // `""` is a name someone wrote, so presence is the test, not truthiness.
  if (shape.name !== undefined) {
    textBox.name = shape.name;
  }
  if (shape.alt !== undefined) {
    textBox.alt = shape.alt;
  }
  if (shape.title !== undefined) {
    textBox.title = shape.title;
  }
  if (shape.position) {
    textBox.position = shape.position;
  }
  if (shape.wrap) {
    textBox.wrap = shape.wrap;
  }
  if (shape.anchor) {
    textBox.anchor = shape.anchor;
  }
  if (shape.fill) {
    textBox.fill = shape.fill;
  }
  if (shape.outline) {
    textBox.outline = shape.outline;
  }
  if (shape.transform) {
    textBox.transform = shape.transform;
  }
  if (textBody.margins) {
    textBox.margins = textBody.margins;
  }
  if (textBody.autoFit) {
    textBox.autoFit = textBody.autoFit;
  }
  if (textBody.textWrap) {
    textBox.textWrap = textBody.textWrap;
  }
  if (textBody.anchor) {
    textBox.verticalAlign = textBody.anchor;
  }
  if (textBody.wordArt) {
    textBox.wordArt = textBody.wordArt;
  }
  return textBox;
}

/**
 * Convert a TextBox to a ProseMirror textBox node
 */
function convertTextBox(
  textBox: TextBox,
  styleResolver: StyleEngine | null,
  options: {
    placement?: "standalone" | "inlineWithPrevious";
    groupId?: string;
    anchorId: string;
    context: TableConversionContext;
    trackedChange: NonNullable<TextBoxAttrs["_docxTrackedChange"]> | undefined;
    inlineSdts: NonNullable<TextBoxAttrs["_docxInlineSdts"]>;
    /** The host `w:p`'s attribute remainder, when this node stands in for it. */
    hostPreservedAttributes?: PreservedAttribute[];
  },
): PMNode {
  reportSourceContainerPageBreakRun(
    textBox.content,
    "text-box",
    options.context.pageBreakRunSourceDescendants,
    options.context.warnPageBreakProjection,
  );

  const textBoxData: { size?: Partial<TextBox["size"]> } = textBox;
  const textBoxSize = textBoxData.size;
  const widthPx = textBoxSize?.width ? emuToPixels(textBoxSize.width) : 200;
  const heightPx = textBoxSize?.height ? emuToPixels(textBoxSize.height) : undefined;

  // Convert fill color
  let fillColor: string | undefined;
  if (textBox.fill?.color?.rgb) {
    fillColor = `#${textBox.fill.color.rgb}`;
  }

  // Convert outline
  let outlineWidth: number | undefined;
  let outlineColor: string | undefined;
  let outlineStyle: string | undefined;
  if (textBox.outline && textBox.outline.width) {
    outlineWidth = emuToStrokePixels(textBox.outline.width);
    if (textBox.outline.color?.rgb) {
      outlineColor = `#${textBox.outline.color.rgb}`;
    }
    outlineStyle = textBox.outline.style || "solid";
  }

  const transformAttrs = authoredTransformAttrs(textBox.transform);

  // Convert margins from EMU to pixels. A margin the source did not author
  // stays absent: minting the default here wrote it back as an authored inset
  // (`shape.textBody.margins.top: absent became N`), and every consumer of the
  // attr already resolves absence against `DEFAULT_TEXTBOX_MARGINS`.
  const marginTop =
    textBox.margins?.top !== undefined ? emuToPixels(textBox.margins.top) : undefined;
  const marginBottom =
    textBox.margins?.bottom !== undefined ? emuToPixels(textBox.margins.bottom) : undefined;
  const marginLeft =
    textBox.margins?.left !== undefined ? emuToPixels(textBox.margins.left) : undefined;
  const marginRight =
    textBox.margins?.right !== undefined ? emuToPixels(textBox.margins.right) : undefined;

  // Convert text box content to PM nodes.
  //
  // `w:txbxContent` is a story of its own: a comment range open in the body
  // does not cover what the box holds, and a marker inside the box does not
  // close it. Carrying the body's open set in would mark the box's text, and
  // the save would then write a second range for that comment id inside the
  // box; carrying the box's markers out would close a body range at the box.
  const textBoxContext = { ...options.context, openCommentIds: new Set<number>() };
  const contentNodes: PMNode[] = [];
  for (const block of textBox.content) {
    if (block.type === "paragraph") {
      contentNodes.push(
        ...convertParagraphWithTextBoxes(block, styleResolver, {
          textBoxGroupId: textBoxContext.nextTextBoxGroupId(),
          context: textBoxContext,
        }),
      );
      continue;
    }
    contentNodes.push(convertTable(block, styleResolver, textBoxContext));
  }

  // Ensure at least one paragraph
  if (contentNodes.length === 0) {
    contentNodes.push(schema.node("paragraph", {}, []));
  }

  // Map wrap settings into the PM textBox attrs so the page renderer can
  // build floating exclusion rects for body text wrapping (eigenpal #474).
  // Mirrors the float/displayMode derivation used by `convertImage` above.
  const wrapType = textBox.wrap?.type;
  const wrapText = textBox.wrap?.wrapText;
  const hAlign = textBox.position?.horizontal.alignment;
  let position: ImagePositionAttrs | undefined;
  if (textBox.position) {
    position = {
      horizontal: {
        relativeTo: textBox.position.horizontal.relativeTo,
        ...(textBox.position.horizontal.posOffset !== undefined
          ? { posOffset: textBox.position.horizontal.posOffset }
          : {}),
        ...(textBox.position.horizontal.alignment
          ? { align: textBox.position.horizontal.alignment }
          : {}),
      },
      vertical: {
        relativeTo: textBox.position.vertical.relativeTo,
        ...(textBox.position.vertical.posOffset !== undefined
          ? { posOffset: textBox.position.vertical.posOffset }
          : {}),
        ...(textBox.position.vertical.alignment
          ? { align: textBox.position.vertical.alignment }
          : {}),
      },
    };
  }

  let cssFloat: "left" | "right" | "none" | undefined;
  if (wrapType === undefined || wrapType === "inline") {
    cssFloat = "none";
  } else if (wrapType === "topAndBottom") {
    cssFloat = "none";
  } else if (wrapType === "square" || wrapType === "tight" || wrapType === "through") {
    if (wrapText === "left") {
      cssFloat = "right";
    } else if (wrapText === "right") {
      cssFloat = "left";
    } else if (hAlign === "left") {
      cssFloat = "left";
    } else if (hAlign === "right") {
      cssFloat = "right";
    } else {
      cssFloat = "none";
    }
  } else {
    cssFloat = "none";
  }

  let displayMode: "inline" | "block" | "float";
  if (wrapType === undefined || wrapType === "inline") {
    displayMode = "inline";
  } else if (wrapType === "topAndBottom") {
    displayMode = "block";
  } else if (wrapType === "behind" || wrapType === "inFront") {
    displayMode = "float";
  } else if (cssFloat !== "none") {
    displayMode = "float";
  } else {
    displayMode = "block";
  }

  // Presence, not truthiness: an authored inset of 0 is a value, and testing
  // it for truth dropped it on save (`shape.wrap.distB: N became absent`).
  const distTop = textBox.wrap?.distT !== undefined ? emuToPixels(textBox.wrap.distT) : undefined;
  const distBottom =
    textBox.wrap?.distB !== undefined ? emuToPixels(textBox.wrap.distB) : undefined;
  const distLeft = textBox.wrap?.distL !== undefined ? emuToPixels(textBox.wrap.distL) : undefined;
  const distRight = textBox.wrap?.distR !== undefined ? emuToPixels(textBox.wrap.distR) : undefined;

  return schema.node(
    "textBox",
    {
      width: widthPx,
      height: heightPx,
      // The pixel values here are lossy; `fromProseDoc` writes these back
      // while the pixels still project from them.
      _docxAuthoredEmu: authoredEmuAttrs({
        width: textBoxSize?.width,
        height: textBoxSize?.height,
        outlineWidth: textBox.outline?.width,
        marginTop: textBox.margins?.top,
        marginBottom: textBox.margins?.bottom,
        marginLeft: textBox.margins?.left,
        marginRight: textBox.margins?.right,
        ...wrapDistanceEmu(textBox.wrap),
      }),
      anchor: copiedDrawingAnchor(textBox.anchor),
      wrapDistanceSlots: copiedWrapDistanceSlots(textBox.wrap?.distanceSlots),
      wrapPolygon: copiedWrapPolygon(textBox.wrap?.polygon),
      autoFit: textBox.autoFit,
      wordArt: textBox.wordArt,
      textWrap: textBox.textWrap,
      verticalAlign: textBox.verticalAlign,
      textBoxId: textBox.id,
      textBoxName: textBox.name,
      alt: textBox.alt,
      title: textBox.title,
      fillColor,
      outlineWidth,
      outlineColor,
      outlineStyle,
      ...transformAttrs,
      marginTop,
      marginBottom,
      marginLeft,
      marginRight,
      displayMode,
      cssFloat,
      wrapType: wrapType ?? "inline",
      wrapText,
      distTop,
      distBottom,
      distLeft,
      distRight,
      position,
      _docxPlacement: options.placement,
      _docxGroupId: options.groupId,
      _docxAnchorId: options.anchorId,
      _docxTextBodyContentState:
        textBox.content.length === 0 ? { type: "source-empty" } : { type: "authored" },
      _docxTrackedChange: options.trackedChange,
      _docxInlineSdts: options.inlineSdts.length > 0 ? options.inlineSdts : undefined,
      _preservedAttributes: options.hostPreservedAttributes,
    },
    contentNodes,
  );
}

/**
 * Convert HeaderFooter content (array of Paragraph/Table blocks) to a ProseMirror document.
 * Used for editing headers/footers in their own ProseMirror editor and for
 * the unified header/footer render pipeline (see
 * `core/layout-bridge/headerFooterLayout.ts`). `theme` lives in
 * `ToProseDocOptions` for future themeColor cell-shading resolution; folio's
 * `convertTable` does not yet thread it (orthogonal upstream divergence).
 */
export function headerFooterToProseDoc(
  content: BlockContent[],
  options?: ToProseDocOptions,
): PMNode {
  const nodes: PMNode[] = [];
  const styleResolver = options?.styles ? createStyleEngine(options.styles) : null;
  const theme = options?.theme ?? null;
  const nextTextBoxGroupId = createTextBoxGroupIdFactory();
  const nextHyperlinkInstanceIndex = createHyperlinkInstanceIndexAllocator();
  const pairedBookmarkIds = collectPairedBookmarkIds(content);
  const conversionContext = {
    theme,
    nextTextBoxGroupId,
    nextHyperlinkInstanceIndex,
    pairedBookmarkIds,
    pageBreakRunSourceDescendants: buildPageBreakRunSourceDescendantIndex(content),
    storyRangedCommentIds: rangedCommentIds(content),
    openCommentIds: new Set<number>(),
    warnPageBreakProjection: pageBreakProjectionWarn(options?.warn),
  };

  const convertBlocks = (blocks: BlockContent[]): PMNode[] => {
    const out: PMNode[] = [];
    for (const block of blocks) {
      switch (block.type) {
        case "paragraph": {
          const isDetachedWatermarkHost = Reflect.get(block, DETACHED_WATERMARK_HOST) === true;
          const paragraphNodes = convertParagraphWithTextBoxes(block, styleResolver, {
            textBoxGroupId: nextTextBoxGroupId(),
            context: conversionContext,
            preserveEmptyWrapper: isDetachedWatermarkHost,
          });
          if (isDetachedWatermarkHost) {
            const paragraphNodeIndex = paragraphNodes.findIndex(
              ({ type }) => type.name === "paragraph",
            );
            const paragraphNode = paragraphNodes[paragraphNodeIndex];
            if (paragraphNode) {
              paragraphNodes[paragraphNodeIndex] = recreateProseNodeWithParagraphPropertySource(
                paragraphNode,
                { attrs: { ...paragraphNode.attrs, _detachedWatermarkHost: true } },
              );
            }
          }
          out.push(...paragraphNodes);
          break;
        }
        case "table":
          out.push(convertTable(block, styleResolver, conversionContext));
          break;
        case "blockSdt":
          out.push(convertBlockSdt(block, convertBlocks));
          break;
        case "preservedBlock":
          out.push(convertPreservedBlock(block));
          break;
        default: {
          const unsupported: never = block;
          panic(`Unsupported block content: ${JSON.stringify(unsupported)}`);
        }
      }
    }
    return out;
  };

  nodes.push(...convertBlocks(content));
  // Caret affordance after a final isolating blockSdt is handled by
  // prosemirror-gapcursor at runtime; we no longer pad the converted doc
  // with a synthetic trailing paragraph because that paragraph survives
  // the reverse pass and pollutes both round-trip saves and
  // `setContentControlContent(filter, blocks)` callers that pass blocks
  // ending in a nested blockSdt.

  if (nodes.length === 0) {
    nodes.push(schema.node("paragraph", {}, []));
  }

  const pmDoc = stampNumberedRefFieldBaselines(schema.node("doc", null, nodes));
  assertValidProseMirrorDocument(
    pmDoc,
    "Header/footer conversion produced an invalid ProseMirror document",
  );
  return pmDoc;
}

export function footnoteToProseDoc(content: BlockContent[], options?: ToProseDocOptions): PMNode {
  return headerFooterToProseDoc(content, options);
}

/**
 * Create an empty ProseMirror document
 */
export function createEmptyDoc(): PMNode {
  return schema.node("doc", null, [schema.node("paragraph", {}, [])]);
}
