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

import { createStyleEngine } from "../../style-engine";
import type { StyleEngine, TableCellParagraphSpacingOverlay } from "../../style-engine";
import type {
  BlockContent,
  BlockSdt,
  Document,
  Paragraph,
  ParagraphFormatting,
  Run,
  RunPropertyChange,
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
  TableLook,
  SimpleField,
  ComplexField,
  InlineSdt,
  Insertion,
  Deletion,
  DrawingContent,
  MoveFrom,
  MoveTo,
  MathEquation,
  ShapeTextBody,
  Theme,
} from "../../types/document";
import {
  mergeParagraphFormatting,
  mergeParagraphTabStops,
} from "../../utils/paragraphFormattingMerge";
import { resolveColorValueToHex } from "../../docx/drawingUtils";
import {
  linkProseParagraphPropertySource,
  recreateProseNodeWithParagraphPropertySource,
} from "../../docx/paragraphPropertySource";
import { mergeTextFormatting } from "../../utils/textFormattingMerge";
import { tableOfContentsStyleLevel } from "../../utils/tableOfContentsStyle";
import { emuToPixels } from "../../utils/units";
import { normalizeHorizontalScalePercent } from "../../utils/horizontalScale";
import { setAutospacingBaseValue } from "../autospacingBase";
import { buildRunFormattingOverrideAttrs } from "../extensions/marks/RunFormattingOverrideExtension";
import { directionFromBidi } from "../paragraphDirection";
import { lineSpacingProvenanceFromSpacing } from "../paragraphSpacing";
import {
  getParagraphMarkSuppressionOverrides,
  hasDirectRunFormatting,
  stripParagraphMarkFormattingForBodyRuns,
  stripParagraphMarkOnlyFormatting,
  suppressParagraphMarkFormatting,
} from "../runStyleFormatting";
import { schema } from "../schema";
import {
  COMPLEX_SCRIPT_RUN_PROPERTY_KEYS,
  RUN_FORMATTING_PROPERTY_SPECS,
  type ComplexScriptRunPropertyKey,
} from "../schema/marks";
import { cascadeStyleTextFormatting } from "../styles/styleToggleCascade";
import type {
  ImagePositionAttrs,
  ParagraphAttrs,
  TableAttrs,
  TableRowAttrs,
  TableCellAttrs,
  TextBoxAttrs,
} from "../schema/nodes";
import { assertValidProseMirrorDocument } from "../validation";
import { stampNumberedRefFieldBaselines } from "../numberedRefFields";
import { canCarryTrackedRunMark, trackedRunInlineAtomDisposition } from "../trackedRunInlineAtoms";
import {
  resolveEffectiveTableCellFormatting,
  type TableCellMarginsAttrs,
  type TableCellPosition,
} from "./effectiveTableCellFormatting";
import { shadingToRunShadingAttrs } from "./runShadingMark";
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
      if (child.type === "bookmarkStart") {
        countBoundary(child.id, "start");
      } else if (child.type === "bookmarkEnd") {
        countBoundary(child.id, "end");
      } else {
        visitRun(child);
      }
    }
  };

  const visitParagraphContent = (content: Paragraph["content"][number]): void => {
    if (content.type === "bookmarkStart") {
      countBoundary(content.id, "start");
    } else if (content.type === "bookmarkEnd") {
      countBoundary(content.id, "end");
    } else if (content.type === "run") {
      visitRun(content);
    } else if (content.type === "hyperlink") {
      visitHyperlink(content);
    } else if (content.type === "simpleField") {
      for (const child of content.content) {
        if (child.type === "hyperlink") {
          visitHyperlink(child);
        } else {
          visitRun(child);
        }
      }
    } else if (content.type === "complexField") {
      for (const run of [...(content.fieldCode ?? []), ...content.fieldResult]) {
        visitRun(run);
      }
    } else if (content.type === "inlineSdt") {
      for (const child of content.content) {
        visitParagraphContent(child);
      }
    } else if (
      content.type === "insertion" ||
      content.type === "deletion" ||
      content.type === "moveFrom" ||
      content.type === "moveTo"
    ) {
      for (const child of content.content) {
        visitParagraphContent(child);
      }
    }
  };

  const visitBlocks = (nestedBlocks: readonly BlockContent[]): void => {
    for (const block of nestedBlocks) {
      if (block.type === "paragraph") {
        for (const content of block.content) {
          visitParagraphContent(content);
        }
      } else if (block.type === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            visitBlocks(cell.content);
          }
        }
      } else {
        visitBlocks(block.content);
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
  };

  const convertBodyBlocks = (blocks: BlockContent[]): PMNode[] => {
    const out: PMNode[] = [];
    for (const block of blocks) {
      if (block.type === "paragraph") {
        out.push(
          ...convertParagraphWithTextBoxes(block, styleResolver, {
            textBoxGroupId: nextTextBoxGroupId(),
            context: conversionContext,
          }),
        );
      } else if (block.type === "table") {
        out.push(convertTable(block, styleResolver, conversionContext));
      } else {
        out.push(convertBlockSdt(block, convertBodyBlocks));
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
  nextHyperlinkInstanceIndex: HyperlinkInstanceIndexAllocator,
  pairedBookmarkIds: ReadonlySet<number>,
  activeCommentIds?: Set<number>,
  extraRunFormatting?: TextFormatting,
  tableParagraphOverlay?: TableCellParagraphSpacingOverlay,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
): PMNode {
  let pageBreakRunOwnerId = 0;
  const nextPageBreakRunOwnerId = (): number => pageBreakRunOwnerId++;
  const attrs = paragraphFormattingToAttrs(paragraph, styleResolver, tableParagraphOverlay);
  const isTocParagraph = attrs._tableOfContentsLevel !== undefined;
  const inlineNodes: PMNode[] = [];
  let inlineOffset = 0;
  let bookmarksArr: { id: number; name: string }[] | undefined;
  let emptyHyperlinks: NonNullable<ParagraphAttrs["_emptyHyperlinks"]> | undefined;

  // Track active comment ranges for this paragraph
  const commentIds = activeCommentIds ?? new Set<number>();
  const emitInlineNodes = (nodes: PMNode[]): void => {
    if (nodes.length === 0) {
      return;
    }
    const markedNodes = applyCommentMarks(nodes, commentIds);
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

  // Get style-based text formatting (font size, bold, color, etc.)
  let styleRunFormatting: TextFormatting | undefined;
  let paragraphStyleRunFormatting: TextFormatting | undefined;
  let paragraphStyleFontFamily: TextFormatting["fontFamily"] | undefined;
  if (styleResolver) {
    const resolved = styleResolver.resolveParagraphStyle(paragraph.formatting?.styleId);
    styleRunFormatting = resolved.runFormatting;
    const paragraphStyle = paragraph.formatting?.styleId
      ? (styleResolver.getStyle(paragraph.formatting.styleId) ??
        styleResolver.getDefaultParagraphStyle())
      : styleResolver.getDefaultParagraphStyle();
    paragraphStyleRunFormatting =
      paragraphStyle?.type === "paragraph" ? paragraphStyle.rPr : undefined;
    paragraphStyleFontFamily = resolveParagraphStyleFontFamily(
      paragraph.formatting?.styleId,
      styleResolver,
    );
  }

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
  const ordinaryStyleFormatting =
    paragraph.formatting?.styleId === undefined
      ? mergeTextFormatting(styleRunFormatting, extraRunFormatting)
      : mergeTextFormatting(extraRunFormatting, styleRunFormatting);
  const orderedToggleFormatting = cascadeStyleTextFormatting(
    [
      { formatting: styleResolver?.getDocDefaults()?.rPr, type: "defaults" },
      { formatting: extraRunFormatting, type: "style" },
      { formatting: paragraphStyleRunFormatting, type: "style" },
    ],
    {
      ordinaryFormatting: ordinaryStyleFormatting,
    },
  );
  let baseRunFormatting = orderedToggleFormatting.formatting;
  // A table style can carry legacy theme/fallback fonts from the template
  // that created it. Preserve the paragraph style's authored font slots over
  // the table contribution; direct run formatting still wins later.
  if (paragraphStyleFontFamily) {
    baseRunFormatting = mergeTextFormatting(baseRunFormatting, {
      fontFamily: paragraphStyleFontFamily,
    });
  }
  // w:pPr/w:rPr formats the paragraph mark, not the visible runs of a named
  // paragraph style. Style-less generated documents historically use it as
  // their highest-precedence run default.
  const defaultCharacterFormatting = styleResolver?.getDefaultCharacterStyle()?.rPr;
  const ordinaryBaseWithDefaultCharacter = mergeTextFormatting(
    defaultCharacterFormatting,
    baseRunFormatting,
  );
  const defaultCharacterStyleCascade = cascadeStyleTextFormatting(
    [
      { cascade: orderedToggleFormatting, type: "carried" },
      { formatting: defaultCharacterFormatting, type: "style" },
    ],
    { ordinaryFormatting: ordinaryBaseWithDefaultCharacter },
  );
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
        getInheritedRunFormatting,
        styleResolver,
        moveKind,
        textBoxAnchors,
      ),
    );
  };

  for (const content of paragraph.content) {
    if (content.type === "commentRangeStart") {
      commentIds.add(content.id);
    } else if (content.type === "commentRangeEnd") {
      commentIds.delete(content.id);
    } else if (content.type === "commentReference") {
      anchorPointComment(inlineNodes, content.id);
    } else if (content.type === "run") {
      emitInlineNodes(
        convertRun(
          content,
          getInheritedRunFormatting(content.formatting),
          nextPageBreakRunOwnerId,
          styleResolver,
          textBoxAnchors,
        ),
      );
    } else if (content.type === "hyperlink") {
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
        continue;
      }
      emitInlineNodes(linkNodes);
    } else if (content.type === "simpleField" || content.type === "complexField") {
      emitInlineNode(
        convertField(content, {
          getInheritedRunFormatting,
          styleResolver,
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          textBoxAnchors,
        }),
      );
    } else if (content.type === "inlineSdt") {
      emitInlineNode(
        convertInlineSdt(
          content,
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          getInheritedRunFormatting,
          styleResolver,
          textBoxAnchors,
        ),
      );
    } else if (content.type === "insertion" || content.type === "moveTo") {
      emitTrackedChange(content, "insertion", content.type === "moveTo" ? "moveTo" : null);
    } else if (content.type === "deletion" || content.type === "moveFrom") {
      emitTrackedChange(content, "deletion", content.type === "moveFrom" ? "moveFrom" : null);
    } else if (content.type === "mathEquation") {
      emitInlineNode(convertMathEquation(content));
    } else if (content.type === "bookmarkStart" && pairedBookmarkIds.has(content.id)) {
      emitInlineNode(
        schema.node("bookmarkBoundary", {
          type: "start",
          id: content.id,
          name: content.name,
          colFirst: content.colFirst,
          colLast: content.colLast,
        }),
      );
    } else if (content.type === "bookmarkEnd" && pairedBookmarkIds.has(content.id)) {
      emitInlineNode(schema.node("bookmarkBoundary", { type: "end", id: content.id }));
    } else if (content.type === "bookmarkStart") {
      // Legacy structural placement records only the start on a paragraph and
      // uses the paragraph attr to preserve its existing save behavior.
      if (!bookmarksArr) {
        bookmarksArr = [];
      }
      bookmarksArr.push({ id: content.id, name: content.name });
    }
  }

  if (bookmarksArr) {
    attrs.bookmarks = bookmarksArr;
  }
  if (emptyHyperlinks) {
    attrs._emptyHyperlinks = emptyHyperlinks;
  }

  const proseParagraph = schema.node("paragraph", attrs, inlineNodes);
  linkProseParagraphPropertySource(proseParagraph, paragraph);
  return proseParagraph;
}

const resolveParagraphStyleFontFamily = (
  styleId: string | undefined,
  styleResolver: StyleEngine,
): TextFormatting["fontFamily"] | undefined => {
  let style = styleId ? styleResolver.getStyle(styleId) : styleResolver.getDefaultParagraphStyle();
  const visited = new Set<string>();
  while (style?.type === "paragraph" && !visited.has(style.styleId)) {
    visited.add(style.styleId);
    if (style.rPr?.fontFamily) {
      return style.rPr.fontFamily;
    }
    style = style.basedOn ? styleResolver.getStyle(style.basedOn) : undefined;
  }
  return undefined;
};

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
  getInheritedRunFormatting: RunFormattingResolver,
  styleResolver?: StyleEngine | null,
  moveKind: "moveFrom" | "moveTo" | null = null,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
): PMNode[] {
  const nodes: PMNode[] = [];
  for (const item of change.content) {
    if (item.type === "run") {
      nodes.push(
        ...convertRun(
          item,
          getInheritedRunFormatting(item.formatting),
          nextPageBreakRunOwnerId,
          styleResolver,
          textBoxAnchors,
        ),
      );
    } else if (item.type === "hyperlink") {
      const currentHyperlinkIndex = nextHyperlinkInstanceIndex();
      nodes.push(
        ...convertHyperlink(item, {
          getInheritedRunFormatting,
          styleResolver,
          hyperlinkIndex: currentHyperlinkIndex,
          textBoxAnchors,
          nextPageBreakRunOwnerId,
        }),
      );
    } else if (item.type === "simpleField" || item.type === "complexField") {
      const fieldNode = convertField(item, {
        getInheritedRunFormatting,
        styleResolver,
        nextHyperlinkInstanceIndex,
        nextPageBreakRunOwnerId,
        textBoxAnchors,
      });
      if (fieldNode) {
        nodes.push(fieldNode);
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
      nodes.push(
        ...convertTrackedChange(
          item,
          nestedMarkType,
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          getInheritedRunFormatting,
          styleResolver,
          nestedMoveKind,
          textBoxAnchors,
        ),
      );
    } else if (item.type === "bookmarkStart") {
      nodes.push(
        schema.node("bookmarkBoundary", {
          type: "start",
          id: item.id,
          name: item.name,
          colFirst: item.colFirst,
          colLast: item.colLast,
        }),
      );
    } else if (item.type === "bookmarkEnd") {
      nodes.push(schema.node("bookmarkBoundary", { type: "end", id: item.id }));
    } else {
      const unsupported: never = item;
      panic(`Unsupported tracked-run content: ${JSON.stringify(unsupported)}`);
    }
  }

  // SAFETY: markType is "insertion" | "deletion", both registered in schema
  const mark = schema.marks[markType]!.create({
    revisionId: change.info.id,
    author: change.info.author,
    date: change.info.date ?? null,
    utcDate: change.info.utcDate?.value ?? null,
    initials: change.info.initials ?? null,
    moveKind,
  });

  return nodes.map((node) => {
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
    return node;
  });
}

/**
 * Convert ParagraphFormatting to ProseMirror paragraph attrs
 *
 * If a styleResolver is provided, resolves style-based formatting and merges
 * with inline formatting. Inline formatting takes precedence.
 */
function paragraphFormattingToAttrs(
  paragraph: Paragraph,
  styleResolver: StyleEngine | null,
  tableParagraphOverlay?: TableCellParagraphSpacingOverlay,
): ParagraphAttrs {
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
  if (paragraph.listRendering?.numFmt) {
    attrs.listNumFmt = paragraph.listRendering.numFmt;
  }
  if (paragraph.listRendering?.isBullet) {
    attrs.listIsBullet = paragraph.listRendering.isBullet;
  }
  if (paragraph.listRendering?.isLegal) {
    attrs.listIsLegal = paragraph.listRendering.isLegal;
  }
  if (paragraph.listRendering?.marker) {
    attrs.listMarker = paragraph.listRendering.marker;
  }
  if (paragraph.listRendering?.markerTemplate) {
    attrs.listMarkerTemplate = paragraph.listRendering.markerTemplate;
  }
  if (paragraph.listRendering?.markerHidden) {
    attrs.listMarkerHidden = paragraph.listRendering.markerHidden;
  }
  if (paragraph.listRendering?.markerFormatting) {
    attrs.listMarkerFormatting = paragraph.listRendering.markerFormatting;
  }
  if (paragraph.listRendering?.markerAlignment) {
    attrs.listMarkerAlignment = paragraph.listRendering.markerAlignment;
  }
  if (paragraph.listRendering?.markerSuffix) {
    attrs.listMarkerSuffix = paragraph.listRendering.markerSuffix;
  }
  if (paragraph.listRendering?.markerAllCaps) {
    attrs.listMarkerAllCaps = paragraph.listRendering.markerAllCaps;
  }
  if (paragraph.listRendering?.implicitChildLevelAdvances !== undefined) {
    attrs.listImplicitChildLevelAdvances = paragraph.listRendering.implicitChildLevelAdvances;
  }
  if (paragraph.listRendering?.markerSecondSlotOffsetTwips !== undefined) {
    attrs.listMarkerSecondSlotOffsetTwips = paragraph.listRendering.markerSecondSlotOffsetTwips;
  }
  if (paragraph.listRendering?.levelNumFmts) {
    attrs.listLevelNumFmts = paragraph.listRendering.levelNumFmts;
  }
  if (paragraph.listRendering && "levelStarts" in paragraph.listRendering) {
    const { levelStarts } = paragraph.listRendering;
    if (Array.isArray(levelStarts) && levelStarts.every((value) => typeof value === "number")) {
      attrs.listLevelStarts = levelStarts;
    }
  }
  if (paragraph.listRendering?.abstractNumId !== undefined) {
    attrs.listAbstractNumId = paragraph.listRendering.abstractNumId;
  }
  if (paragraph.listRendering?.startOverride !== undefined) {
    attrs.listStartOverride = paragraph.listRendering.startOverride;
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
  // style's paragraph-spacing fields in between docDefaults and this
  // paragraph's own style chain — see resolveParagraphStyleInTable.
  let stylePpr: Paragraph["formatting"] | undefined;
  if (styleResolver) {
    const resolved = styleResolver.resolveParagraphStyleInTable(styleId, tableParagraphOverlay);
    stylePpr = resolved.paragraphFormatting;

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
      formatting?.numPr?.numId === 0 && stylePpr?.numPr !== undefined && stylePpr.numPr.numId !== 0;
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
    if (stylePpr?.numPr && formatting?.numPr?.numId === undefined && stylePpr.numPr.numId !== 0) {
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

  return attrs;
}

// ============================================================================
// TABLE CONVERSION
// ============================================================================

/**
 * A table style's (or one of its `w:tblStylePr` conditional regions')
 * contribution to cell formatting: cell properties, run defaults, and — for
 * the table-row-height fix — the paragraph-spacing overlay described on
 * {@link TableCellParagraphSpacingOverlay}.
 */
type TableConditionalStyle = {
  tcPr?: TableCellFormatting;
  rPr?: TextFormatting;
  pPr?: TableCellParagraphSpacingOverlay;
};

/**
 * Pick the paragraph-spacing fields out of a table style's (or conditional
 * region's) `w:pPr` for use as the cell-paragraph cascade overlay. Narrower
 * than the full `ParagraphFormatting` bag — see
 * {@link TableCellParagraphSpacingOverlay}.
 */
function extractTableParagraphSpacingOverlay(
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
  const paragraphSpacingOverlay = extractTableParagraphSpacingOverlay(conditional.pPr);

  const result: TableConditionalStyle = {};
  if (conditional.tcPr) {
    result.tcPr = conditional.tcPr;
  }
  if (mergedRunProps) {
    result.rPr = mergedRunProps;
  }
  if (paragraphSpacingOverlay) {
    result.pPr = paragraphSpacingOverlay;
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
  const paragraphSpacingOverlay = extractTableParagraphSpacingOverlay(style.pPr);

  const result: TableConditionalStyle = {};
  if (style.tcPr) {
    result.tcPr = style.tcPr;
  }
  if (mergedRunProps) {
    result.rPr = mergedRunProps;
  }
  if (paragraphSpacingOverlay) {
    result.pPr = paragraphSpacingOverlay;
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
  // matching the tcPr/rPr merges above — see extractTableParagraphSpacingOverlay.
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
      const startRow = vMerge === "continue" ? activeMerges.get(colIndex) : undefined;
      const info = {
        cell,
        colIndex,
        colspan,
        vMerge,
        startRow,
        hasMeaningfulContent: tableCellHasMeaningfulContent(cell),
        shouldSkip: vMerge === "continue" && startRow !== undefined,
      };
      colIndex += colspan;
      return info;
    });
    const rowWouldBeEmpty = rowCells.length > 0 && rowCells.every((cell) => cell.shouldSkip);

    for (const cellInfo of rowCells) {
      const { colIndex: cellColIndex, vMerge, startRow, hasMeaningfulContent } = cellInfo;
      const key = `${rowIndex}-${cellColIndex}`;

      if (vMerge === "restart") {
        // Start of a new vertical merge
        activeMerges.set(cellColIndex, rowIndex);
        result.set(key, { rowSpan: 1, skip: false });
      } else if (vMerge === "continue") {
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
        activeMerges.delete(cellColIndex);
        result.set(key, { rowSpan: 1, skip: false });
      }
    }
  }

  return result;
}

function clearActiveVerticalMerges(
  activeMerges: Map<number, number>,
  result: Map<string, RowSpanInfo>,
): void {
  for (const [colIndex, startRow] of activeMerges) {
    const restartCell = result.get(`${startRow}-${colIndex}`);
    if (restartCell) {
      restartCell.preserveVMergeRestart = true;
    }
  }
  activeMerges.clear();
}

function tableCellHasMeaningfulContent(cell: TableCell): boolean {
  return cell.content.some(blockHasMeaningfulContent);
}

function blockHasMeaningfulContent(block: Paragraph | Table): boolean {
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
};

function convertTable(
  table: Table,
  styleResolver: StyleEngine | null,
  context: TableConversionContext,
): PMNode {
  // Calculate rowSpan values from vMerge
  const rowSpanMap = calculateRowSpans(table);

  // Get column widths from table grid
  const columnWidths = table.columnWidths;

  // Calculate total width from columnWidths if available (for percentage calculation)
  const totalWidth = columnWidths?.reduce((sum, w) => sum + w, 0) ?? 0;

  // Get the table style's conditional formatting
  const tableStyleId = table.formatting?.styleId;
  const look = table.formatting?.look;

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
  if (table.formatting) {
    attrs._originalFormatting = table.formatting;
  }
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

  const bandingEnabledH = look?.noHBand !== true;
  const bandingEnabledV = look?.noVBand !== true;

  // Track data row index (excluding header rows) for banding
  let dataRowIndex = 0;
  const totalRows = table.rows.length;
  const gridColumnCount = columnWidths?.length ?? 0;
  const totalColumns = gridColumnCount > 0 ? gridColumnCount : countTableColumns(table.rows);
  const rows = table.rows.map((row, rowIndex) => {
    // Conditional formatting flag: firstRow in tblLook means "apply first-row styling"
    const isFirstRowStyled = rowIndex === 0 && !!look?.firstRow;
    const isLastRow = rowIndex === totalRows - 1 && !!look?.lastRow;

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
  tableLook?: TableLook,
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
  let attrs: TableRowAttrs = attrsWithoutStructuralChange;
  if (row.structuralChange?.type === "tableRowInsertion") {
    attrs = {
      ...attrsWithoutStructuralChange,
      trIns: {
        revisionId: row.structuralChange.info.id,
        author: row.structuralChange.info.author,
        date: row.structuralChange.info.date ?? null,
        ...(row.structuralChange.info.utcDate
          ? { utcDate: row.structuralChange.info.utcDate.value }
          : {}),
        ...(row.structuralChange.info.initials
          ? { initials: row.structuralChange.info.initials }
          : {}),
      },
    };
  } else if (row.structuralChange?.type === "tableRowDeletion") {
    attrs = {
      ...attrsWithoutStructuralChange,
      trDel: {
        revisionId: row.structuralChange.info.id,
        author: row.structuralChange.info.author,
        date: row.structuralChange.info.date ?? null,
        ...(row.structuralChange.info.utcDate
          ? { utcDate: row.structuralChange.info.utcDate.value }
          : {}),
        ...(row.structuralChange.info.initials
          ? { initials: row.structuralChange.info.initials }
          : {}),
      },
    };
  }

  const numCells = row.cells.length;
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
  if (effectiveCells.length === 0) {
    const fallback: TableCell = {
      type: "tableCell",
      content: [{ type: "paragraph", content: [] }],
    };
    if (totalCols > 1) {
      fallback.formatting = { gridSpan: totalCols };
    }
    effectiveCells = [fallback];
  }

  // Track column index for mapping to columnWidths (accounting for colspan)
  let colIndex = row.formatting?.gridBefore ?? 0;
  const cells: PMNode[] = [];

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
  if (cell.structuralChange?.type === "tableCellInsertion") {
    attrs.cellMarker = {
      kind: "ins",
      info: {
        revisionId: cell.structuralChange.info.id,
        author: cell.structuralChange.info.author,
        date: cell.structuralChange.info.date ?? null,
        ...(cell.structuralChange.info.utcDate
          ? { utcDate: cell.structuralChange.info.utcDate.value }
          : {}),
        ...(cell.structuralChange.info.initials
          ? { initials: cell.structuralChange.info.initials }
          : {}),
      },
    };
  } else if (cell.structuralChange?.type === "tableCellDeletion") {
    attrs.cellMarker = {
      kind: "del",
      info: {
        revisionId: cell.structuralChange.info.id,
        author: cell.structuralChange.info.author,
        date: cell.structuralChange.info.date ?? null,
        ...(cell.structuralChange.info.utcDate
          ? { utcDate: cell.structuralChange.info.utcDate.value }
          : {}),
        ...(cell.structuralChange.info.initials
          ? { initials: cell.structuralChange.info.initials }
          : {}),
      },
    };
  } else if (cell.structuralChange?.type === "tableCellMerge") {
    attrs.cellMarker = {
      kind: "merge",
      info: {
        revisionId: cell.structuralChange.info.id,
        author: cell.structuralChange.info.author,
        date: cell.structuralChange.info.date ?? null,
        ...(cell.structuralChange.info.utcDate
          ? { utcDate: cell.structuralChange.info.utcDate.value }
          : {}),
        ...(cell.structuralChange.info.initials
          ? { initials: cell.structuralChange.info.initials }
          : {}),
      },
      ...(cell.structuralChange.verticalMerge !== undefined
        ? { verticalMerge: cell.structuralChange.verticalMerge }
        : {}),
      ...(cell.structuralChange.verticalMergeOriginal !== undefined
        ? { verticalMergeOriginal: cell.structuralChange.verticalMergeOriginal }
        : {}),
    };
  }
  if (preserveVMergeRestart) {
    attrs._preserveVMergeRestart = true;
  }
  if (vMergeContinuationCells && vMergeContinuationCells.length > 0) {
    attrs._docxVMergeContinuationCells = vMergeContinuationCells;
  }

  // Convert cell content (paragraphs and nested tables)
  const contentNodes: PMNode[] = [];
  for (const content of cell.content) {
    if (content.type === "paragraph") {
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
    } else {
      // Nested tables - recursively convert
      contentNodes.push(convertTable(content, styleResolver, context));
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
  return convertTableCell({
    cell,
    styleResolver: null,
    context: {
      theme: null,
      nextTextBoxGroupId,
      nextHyperlinkInstanceIndex,
      pairedBookmarkIds: collectPairedBookmarkIds(cell.content),
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
  const runHasPageBreak = (run: Run): boolean =>
    run.content.some((content) => content.type === "break" && content.breakType === "page");
  if (field.type === "complexField" && field.fieldCode.some(runHasPageBreak)) {
    panic(
      "A complex-field instruction containing an explicit page break cannot be represented in the editor model",
    );
  }
  const hasPageBreakContent =
    field.type === "simpleField"
      ? field.content.some((content) =>
          content.type === "run"
            ? runHasPageBreak(content)
            : content.children.some((child) => child.type === "run" && runHasPageBreak(child)),
        )
      : field.fieldResult.some(runHasPageBreak);
  if (hasPageBreakContent) {
    assertPageBreakFieldResultIsRepresentable(field);
  }
  const hasStructuredSourceContent =
    hasPageBreakContent ||
    (field.type === "simpleField" && field.content.some((content) => content.type === "hyperlink"));
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
  const createStructuredField =
    hasConvertedPageBreakContent || (hasStructuredSourceContent && hasConvertedHyperlinkContent);
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
      fldLock: field.fldLock ?? false,
      dirty: field.dirty ?? false,
    },
    createStructuredField ? inlineNodes : undefined,
    marks,
  );
}

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
 */
function convertInlineSdt(
  sdt: InlineSdt,
  nextHyperlinkInstanceIndex: HyperlinkInstanceIndexAllocator,
  nextPageBreakRunOwnerId: PageBreakRunOwnerIdAllocator,
  getInheritedRunFormatting: RunFormattingResolver,
  styleResolver?: StyleEngine | null,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
): PMNode | null {
  const props = sdt.properties;
  const inlineNodes: PMNode[] = [];

  for (const content of sdt.content) {
    if (content.type === "run") {
      const runNodes = convertRun(
        content,
        getInheritedRunFormatting(content.formatting),
        nextPageBreakRunOwnerId,
        styleResolver,
        textBoxAnchors,
      );
      inlineNodes.push(...runNodes);
    } else if (content.type === "hyperlink") {
      const currentHyperlinkIndex = nextHyperlinkInstanceIndex();
      const linkNodes = convertHyperlink(content, {
        getInheritedRunFormatting,
        styleResolver,
        hyperlinkIndex: currentHyperlinkIndex,
        textBoxAnchors,
        nextPageBreakRunOwnerId,
      });
      inlineNodes.push(...linkNodes);
    } else if (content.type === "simpleField" || content.type === "complexField") {
      const fieldNode = convertField(content, {
        getInheritedRunFormatting,
        styleResolver,
        nextHyperlinkInstanceIndex,
        nextPageBreakRunOwnerId,
        textBoxAnchors,
      });
      if (fieldNode) {
        inlineNodes.push(fieldNode);
      }
    } else if (content.type === "inlineSdt") {
      const nestedSdt = convertInlineSdt(
        content,
        nextHyperlinkInstanceIndex,
        nextPageBreakRunOwnerId,
        getInheritedRunFormatting,
        styleResolver,
        textBoxAnchors,
      );
      if (nestedSdt) {
        inlineNodes.push(nestedSdt);
      }
    } else if (content.type === "insertion") {
      inlineNodes.push(
        ...convertTrackedChange(
          content,
          "insertion",
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          getInheritedRunFormatting,
          styleResolver,
          null,
          textBoxAnchors,
        ),
      );
    } else if (content.type === "deletion") {
      inlineNodes.push(
        ...convertTrackedChange(
          content,
          "deletion",
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          getInheritedRunFormatting,
          styleResolver,
          null,
          textBoxAnchors,
        ),
      );
    } else if (content.type === "moveTo") {
      inlineNodes.push(
        ...convertTrackedChange(
          content,
          "insertion",
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          getInheritedRunFormatting,
          styleResolver,
          "moveTo",
          textBoxAnchors,
        ),
      );
    } else if (content.type === "moveFrom") {
      inlineNodes.push(
        ...convertTrackedChange(
          content,
          "deletion",
          nextHyperlinkInstanceIndex,
          nextPageBreakRunOwnerId,
          getInheritedRunFormatting,
          styleResolver,
          "moveFrom",
          textBoxAnchors,
        ),
      );
    } else {
      // content.type === "mathEquation" — narrowed by exhaustion of the
      // InlineSdt['content'] union above.
      const mathNode = convertMathEquation(content);
      if (mathNode) {
        inlineNodes.push(mathNode);
      }
    }
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
  assertPageBreakSourceRunIsRepresentable(run);
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
    const contentNodes = convertRunContent(content, marks, mergedFormatting, textBoxAnchors);
    nodes.push(...contentNodes);
  }

  return nodes;
}

function assertPageBreakSourceRunIsRepresentable(run: Run): void {
  const hasPageBreak = run.content.some(
    (content) => content.type === "break" && content.breakType === "page",
  );
  if (!hasPageBreak) {
    return;
  }

  assertRunContentIsRepresentableBesidePageBreak(run, "A page-break-bearing run");
}

function assertPageBreakFieldResultIsRepresentable(field: SimpleField | ComplexField): void {
  if (field.type === "complexField") {
    for (const run of field.fieldResult) {
      assertRunContentIsRepresentableBesidePageBreak(
        run,
        "A field result with an explicit page break",
      );
    }
    return;
  }

  for (const content of field.content) {
    if (content.type === "run") {
      assertRunContentIsRepresentableBesidePageBreak(
        content,
        "A field result with an explicit page break",
      );
      continue;
    }
    for (const child of content.children) {
      if (child.type === "run") {
        assertRunContentIsRepresentableBesidePageBreak(
          child,
          "A field result with an explicit page break",
        );
      }
    }
  }
}

function assertRunContentIsRepresentableBesidePageBreak(run: Run, ownerDescription: string): void {
  for (const content of run.content) {
    switch (content.type) {
      case "break":
      case "drawing":
      case "endnoteRef":
      case "footnoteRef":
      case "renderedPageBreak":
      case "symbol":
      case "tab":
      case "text":
        continue;
      case "shape":
        if (content.shape.textBody) {
          panic(
            `${ownerDescription} containing a text-box shape cannot be represented in the editor model`,
          );
        }
        continue;
      case "fieldChar":
      case "instrText":
      case "noBreakHyphen":
      case "softHyphen":
        panic(
          `${ownerDescription} containing ${content.type} cannot be represented in the editor model`,
        );
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

type AuthoredRunFormattingCarrier = "preserve" | "reconstruct";

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
      marks: textFormattingToMarks(inherited.formatting, {
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
  const marks = textFormattingToMarks(mergedFormatting, {
    overrideFormatting,
    directFormatting: runFormatting,
    authoredCarrier,
  });

  if (styleId) {
    marks.push(schema.mark("characterStyle", { styleId }));
  }

  return { marks, mergedFormatting };
}

const addDirectFontProvenance = (
  marks: ReturnType<typeof schema.mark>[],
  directFormatting: TextFormatting | undefined,
): void => {
  const directFontProperties: ("fontFamily" | "fontSize" | "color")[] = [];
  if (directFormatting?.fontFamily !== undefined) {
    directFontProperties.push("fontFamily");
  }
  if (directFormatting?.fontSize !== undefined) {
    directFontProperties.push("fontSize");
  }
  if (directFormatting?.color !== undefined) {
    directFontProperties.push("color");
  }
  if (directFontProperties.length === 0) {
    return;
  }

  const index = marks.findIndex(({ type }) => type.name === "runFormattingOverride");
  const existing = index >= 0 ? marks.at(index) : undefined;
  const override = schema.mark("runFormattingOverride", {
    ...existing?.attrs,
    directFontProperties,
  });
  if (index >= 0) {
    marks[index] = override;
    return;
  }
  marks.push(override);
};

const COMPLEX_SCRIPT_MIRRORS = [
  { ordinary: "bold", complex: "boldCs" },
  { ordinary: "italic", complex: "italicCs" },
  { ordinary: "fontSize", complex: "fontSizeCs" },
] as const satisfies readonly {
  ordinary: keyof TextFormatting;
  complex: ComplexScriptRunPropertyKey;
}[];

const addComplexScriptAbsenceProvenance = (
  marks: ReturnType<typeof schema.mark>[],
  directFormatting: TextFormatting | undefined,
): void => {
  const absent = COMPLEX_SCRIPT_MIRRORS.filter(
    ({ ordinary, complex }) =>
      directFormatting?.[ordinary] !== undefined && directFormatting[complex] === undefined,
  ).map(({ complex }) => complex);
  if (absent.length === 0) {
    return;
  }

  const index = marks.findIndex(({ type }) => type.name === "runFormattingOverride");
  const existing = index >= 0 ? marks.at(index) : undefined;
  const existingAbsences = new Set(existing?.attrs["complexScriptPropertyAbsences"] ?? []);
  for (const property of absent) {
    existingAbsences.add(property);
  }
  const override = schema.mark("runFormattingOverride", {
    ...existing?.attrs,
    complexScriptPropertyAbsences: COMPLEX_SCRIPT_RUN_PROPERTY_KEYS.filter((property) =>
      existingAbsences.has(property),
    ),
  });
  if (index >= 0) {
    marks[index] = override;
    return;
  }
  marks.push(override);
};

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
 * Convert RunContent to ProseMirror nodes
 */
function convertRunContent(
  content: RunContent,
  marks: ReturnType<typeof schema.mark>[],
  formatting?: TextFormatting,
  textBoxAnchors?: ReadonlyMap<Shape, string>,
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
      return [withRunBoundaryMarks(convertShape(shp), marks)];
    }

    case "footnoteRef": {
      // Footnote reference - render as superscript number with footnoteRef mark
      const footnoteMark = schema.mark("footnoteRef", {
        id: content.id.toString(),
        noteType: "footnote",
        vertAlign:
          formatting?.vertAlign === "baseline" || formatting?.vertAlign === "superscript"
            ? formatting.vertAlign
            : null,
      });
      return [schema.text(content.id.toString(), [...marks, footnoteMark])];
    }

    case "endnoteRef": {
      // Endnote reference - render as superscript number with footnoteRef mark
      const endnoteMark = schema.mark("footnoteRef", {
        id: content.id.toString(),
        noteType: "endnote",
        vertAlign:
          formatting?.vertAlign === "baseline" || formatting?.vertAlign === "superscript"
            ? formatting.vertAlign
            : null,
      });
      return [schema.text(content.id.toString(), [...marks, endnoteMark])];
    }

    case "fieldChar":
    case "instrText":
      // Complex field structure markers — handled at the run/paragraph
      // level via `convertField`, not as standalone inline content.
      return [];

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
};

function convertImage({ image, rawXml, rawXmlMode }: ConvertImageOptions): PMNode {
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

  // Build transform string if needed (rotation, flip)
  let transform: string | undefined;
  if (image.transform) {
    const transforms: string[] = [];
    if (image.transform.rotation) {
      transforms.push(`rotate(${image.transform.rotation}deg)`);
    }
    if (image.transform.flipH) {
      transforms.push("scaleX(-1)");
    }
    if (image.transform.flipV) {
      transforms.push("scaleY(-1)");
    }
    if (transforms.length > 0) {
      transform = transforms.join(" ");
    }
  }

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
    // Convert EMU to pixels (1 EMU = 1/914400 inch, 1 inch = 96 px)
    borderWidth = Math.round((image.outline.width / 914_400) * 96 * 100) / 100;
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
    docPrName: image.docPrName,
    alt: image.alt,
    title: image.title,
    width: widthPx,
    height: heightPx,
    rId: image.rId,
    wrapType,
    displayMode,
    cssFloat,
    transform,
    // eigenpal #424 (opacity render pipeline). PR #513 added Image.opacity
    // on the model; thread it onto the PM node so the layout-bridge and
    // painter can honor it.
    opacity: image.opacity,
    distTop,
    distBottom,
    distLeft,
    distRight,
    // eigenpal #424: thread wp:srcRect crop fractions through PM attrs.
    cropTop: image.crop?.top,
    cropRight: image.crop?.right,
    cropBottom: image.crop?.bottom,
    cropLeft: image.crop?.left,
    position,
    layoutInCell: image.layoutInCell,
    borderWidth,
    borderColor,
    borderStyle,
    wrapText,
    hlinkHref: image.hlinkHref,
    hlinkRId: image.hlinkRId,
    _docxRawXml: rawXml,
    _docxRawXmlMode: rawXmlMode,
    _docxObjectPreview:
      rawXml !== undefined && /<(?:[A-Za-z_][\w.-]*:)?object(?:\s|>)/u.test(rawXml),
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
          },
          undefined,
          [linkMark],
        ),
      );
      continue;
    }
    if (child.type === "bookmarkEnd") {
      nodes.push(
        schema.node("bookmarkBoundary", { type: "end", id: child.id }, undefined, [linkMark]),
      );
      continue;
    }
    if (child.type === "run") {
      assertPageBreakSourceRunIsRepresentable(child);
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
        nodes.push(...convertRunContent(content, allMarks, mergedFormatting, textBoxAnchors));
      }
    }
  }

  return nodes;
}

/**
 * Convert TextFormatting to ProseMirror marks
 */
type TextFormattingToMarksOptions = {
  overrideFormatting: TextFormatting | undefined;
  /** Direct standard font properties whose authored provenance must survive. */
  directFormatting?: TextFormatting | undefined;
  /** Whether direct authorship is reconstructible from structural marks and the style context. */
  authoredCarrier?: AuthoredRunFormattingCarrier;
};

export function textFormattingToMarks(
  formatting: TextFormatting | undefined,
  options?: TextFormattingToMarksOptions,
): ReturnType<typeof schema.mark>[] {
  if (!formatting) {
    return [];
  }

  const marks: ReturnType<typeof schema.mark>[] = [];
  const overrideFormatting = options ? options.overrideFormatting : formatting;
  let overrideAttrs: ReturnType<typeof buildRunFormattingOverrideAttrs>;
  if (options?.authoredCarrier === "reconstruct") {
    overrideAttrs = buildRunFormattingOverrideAttrs(overrideFormatting, {
      type: "structural-only",
    });
  } else if (options) {
    overrideAttrs = buildRunFormattingOverrideAttrs(overrideFormatting, {
      type: "authored-baseline",
      formatting: options.directFormatting,
    });
  } else {
    overrideAttrs = buildRunFormattingOverrideAttrs(overrideFormatting);
  }

  if (overrideAttrs) {
    marks.push(schema.mark("runFormattingOverride", overrideAttrs));
  }

  // Bold
  if (formatting.bold) {
    marks.push(schema.mark("bold"));
  }

  // Italic
  if (formatting.italic) {
    marks.push(schema.mark("italic"));
  }

  // Underline
  if (formatting.underline && formatting.underline.style !== "none") {
    marks.push(
      schema.mark("underline", {
        style: formatting.underline.style,
        color: formatting.underline.color,
      }),
    );
  }

  // Strikethrough
  if (formatting.strike || formatting.doubleStrike) {
    marks.push(
      schema.mark("strike", {
        double: formatting.doubleStrike || false,
      }),
    );
  }

  // Text color
  if (formatting.color && !formatting.color.auto) {
    marks.push(
      schema.mark("textColor", {
        rgb: formatting.color.rgb,
        themeColor: formatting.color.themeColor,
        themeTint: formatting.color.themeTint,
        themeShade: formatting.color.themeShade,
      }),
    );
  }

  // Highlight
  if (formatting.highlight && formatting.highlight !== "none") {
    marks.push(
      schema.mark("highlight", {
        color: formatting.highlight,
      }),
    );
  }

  // Run shading (w:shd) used as a run background. Folio models highlight as a
  // strict OOXML named-palette union, so an arbitrary fill (e.g. a Word/Google
  // Docs run background) round-trips as a dedicated runShading mark instead of
  // silently disappearing at PM conversion. eigenpal #722 (#712).
  const runShadingMarkAttrs = shadingToRunShadingAttrs(formatting.shading);
  if (runShadingMarkAttrs) {
    marks.push(schema.mark("runShading", runShadingMarkAttrs));
  }

  // Font size
  if (formatting.fontSize) {
    marks.push(
      schema.mark("fontSize", {
        size: formatting.fontSize,
      }),
    );
  }

  // Font family
  if (formatting.fontFamily) {
    marks.push(
      schema.mark("fontFamily", {
        ascii: formatting.fontFamily.ascii,
        hAnsi: formatting.fontFamily.hAnsi,
        eastAsia: formatting.fontFamily.eastAsia,
        cs: formatting.fontFamily.cs,
        hint: formatting.fontFamily.hint,
        asciiTheme: formatting.fontFamily.asciiTheme,
        hAnsiTheme: formatting.fontFamily.hAnsiTheme,
        eastAsiaTheme: formatting.fontFamily.eastAsiaTheme,
        csTheme: formatting.fontFamily.csTheme,
      }),
    );
  }

  if (formatting.language) {
    marks.push(schema.mark("language", formatting.language));
  }

  // Superscript/Subscript
  if (formatting.vertAlign === "superscript") {
    marks.push(schema.mark("superscript"));
  } else if (formatting.vertAlign === "subscript") {
    marks.push(schema.mark("subscript"));
  }

  // All caps (w:caps)
  if (formatting.allCaps) {
    marks.push(schema.mark("allCaps"));
  }

  // Small caps (w:smallCaps)
  if (formatting.smallCaps) {
    marks.push(schema.mark("smallCaps"));
  }

  // Character spacing (spacing, position, scale, kerning)
  const spacing = typeof formatting.spacing === "number" ? formatting.spacing : null;
  const position = typeof formatting.position === "number" ? formatting.position : null;
  const scale = normalizeHorizontalScalePercent(formatting.scale) ?? null;
  const kerning = typeof formatting.kerning === "number" ? formatting.kerning : null;
  if (spacing !== null || position !== null || scale !== null || kerning !== null) {
    marks.push(
      schema.mark("characterSpacing", {
        spacing,
        position,
        scale,
        kerning,
      }),
    );
  }

  // Hidden text (w:vanish). eigenpal #424 (gap 9).
  if (formatting.hidden === true) {
    marks.push(schema.mark("hidden"));
  }

  // Emboss (w:emboss)
  if (formatting.emboss) {
    marks.push(schema.mark("emboss"));
  }

  // Imprint/Engrave (w:imprint)
  if (formatting.imprint) {
    marks.push(schema.mark("imprint"));
  }

  // Text shadow (w:shadow)
  if (formatting.shadow) {
    marks.push(schema.mark("textShadow"));
  }

  // Emphasis mark (w:em)
  if (formatting.emphasisMark && formatting.emphasisMark !== "none") {
    marks.push(schema.mark("emphasisMark", { type: formatting.emphasisMark }));
  }

  // Text outline (w:outline)
  if (formatting.outline) {
    marks.push(schema.mark("textOutline"));
  }

  // eigenpal #424 (gap 10) — per-run RTL direction (w:rtl)
  if (formatting.rtl) {
    marks.push(schema.mark("rtl"));
  }

  // eigenpal #424 (gap 11) — text effect animation (w:effect)
  if (formatting.effect && formatting.effect !== "none") {
    marks.push(schema.mark("textEffect", { effect: formatting.effect }));
  }

  addDirectFontProvenance(marks, options?.directFormatting);
  if (options?.authoredCarrier === "preserve") {
    addComplexScriptAbsenceProvenance(marks, options.directFormatting);
  }

  return marks;
}

// ============================================================================
// SHAPE CONVERSION
// ============================================================================

/**
 * Convert a Shape to a ProseMirror shape node (inline SVG)
 */
function convertShape(shape: Shape): PMNode {
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
  let outlineHeadEnd: NonNullable<Shape["outline"]>["headEnd"] | undefined;
  let outlineTailEnd: NonNullable<Shape["outline"]>["tailEnd"] | undefined;
  if (shape.outline) {
    if (shape.outline.width) {
      outlineWidth = Math.round((shape.outline.width / 914_400) * 96 * 100) / 100;
    }
    if (shape.outline.color) {
      outlineColorValue = shape.outline.color;
      outlineColor = resolveColorValueToHex(shape.outline.color);
    }
    outlineStyle = shape.outline.style || "solid";
    outlineCap = shape.outline.cap;
    outlineHeadEnd = shape.outline.headEnd;
    outlineTailEnd = shape.outline.tailEnd;
  } else {
    outlineWidth = 0;
  }

  let transform: string | undefined;
  if (shape.transform) {
    const transforms: string[] = [];
    if (shape.transform.rotation) {
      transforms.push(`rotate(${shape.transform.rotation}deg)`);
    }
    if (shape.transform.flipH) {
      transforms.push("scaleX(-1)");
    }
    if (shape.transform.flipV) {
      transforms.push("scaleY(-1)");
    }
    if (transforms.length > 0) {
      transform = transforms.join(" ");
    }
  }

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
    shapeType: shapeAttrs.shapeType ?? "rect",
    geometryAdjustments:
      shape.geometryAdjustments === undefined
        ? undefined
        : JSON.stringify(shape.geometryAdjustments),
    shapeId: shape.id,
    width: widthPx,
    height: heightPx,
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
    outlineHeadEnd,
    outlineTailEnd,
    transform,
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
    context.nextHyperlinkInstanceIndex,
    context.pairedBookmarkIds,
    undefined,
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
  for (const { textBox, anchorId, trackedChange, inlineSdts } of textBoxes) {
    nodes.push(
      convertTextBox(textBox, styleResolver, {
        placement:
          isEmptyAfterExtraction && !keepWrapperParagraph ? "standalone" : "inlineWithPrevious",
        groupId: textBoxGroupId,
        anchorId,
        context,
        trackedChange,
        inlineSdts,
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
  if (shape.position) {
    textBox.position = shape.position;
  }
  if (shape.wrap) {
    textBox.wrap = shape.wrap;
  }
  if (shape.fill) {
    textBox.fill = shape.fill;
  }
  if (shape.outline) {
    textBox.outline = shape.outline;
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
  },
): PMNode {
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
    outlineWidth = Math.round((textBox.outline.width / 914_400) * 96 * 100) / 100;
    if (textBox.outline.color?.rgb) {
      outlineColor = `#${textBox.outline.color.rgb}`;
    }
    outlineStyle = textBox.outline.style || "solid";
  }

  // Convert margins from EMU to pixels
  const marginTop = textBox.margins?.top !== undefined ? emuToPixels(textBox.margins.top) : 4;
  const marginBottom =
    textBox.margins?.bottom !== undefined ? emuToPixels(textBox.margins.bottom) : 4;
  const marginLeft = textBox.margins?.left !== undefined ? emuToPixels(textBox.margins.left) : 7;
  const marginRight = textBox.margins?.right !== undefined ? emuToPixels(textBox.margins.right) : 7;

  // Convert text box content to PM nodes
  const contentNodes: PMNode[] = [];
  for (const block of textBox.content) {
    if (block.type === "paragraph") {
      contentNodes.push(
        ...convertParagraphWithTextBoxes(block, styleResolver, {
          textBoxGroupId: options.context.nextTextBoxGroupId(),
          context: options.context,
        }),
      );
      continue;
    }
    contentNodes.push(convertTable(block, styleResolver, options.context));
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

  const distTop = textBox.wrap?.distT ? emuToPixels(textBox.wrap.distT) : undefined;
  const distBottom = textBox.wrap?.distB ? emuToPixels(textBox.wrap.distB) : undefined;
  const distLeft = textBox.wrap?.distL ? emuToPixels(textBox.wrap.distL) : undefined;
  const distRight = textBox.wrap?.distR ? emuToPixels(textBox.wrap.distR) : undefined;

  return schema.node(
    "textBox",
    {
      width: widthPx,
      height: heightPx,
      autoFit: textBox.autoFit,
      textWrap: textBox.textWrap,
      verticalAlign: textBox.verticalAlign,
      textBoxId: textBox.id,
      fillColor,
      outlineWidth,
      outlineColor,
      outlineStyle,
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
      _docxTrackedChange: options.trackedChange,
      _docxInlineSdts: options.inlineSdts.length > 0 ? options.inlineSdts : undefined,
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
  };

  const convertBlocks = (blocks: BlockContent[]): PMNode[] => {
    const out: PMNode[] = [];
    for (const block of blocks) {
      if (block.type === "paragraph") {
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
      } else if (block.type === "table") {
        out.push(convertTable(block, styleResolver, conversionContext));
      } else {
        out.push(convertBlockSdt(block, convertBlocks));
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
