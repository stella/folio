/**
 * Converts tables, rows, and cells to TableBlocks, and text-box nodes to
 * TextBoxBlocks. The two live together because they are mutually recursive:
 * table cells host text boxes and text boxes host tables.
 */

import type { Node as PMNode } from "prosemirror-model";
import type {
  FlowBlock,
  ParagraphBlock,
  TableBlock,
  TableRow,
  TableCell,
  CellBorders,
  TextBoxBlock,
  FloatingTablePosition,
} from "../../layout-engine/types";
import { setTextBoxGroupId } from "../../layout-engine/textBoxGroup";
import { DEFAULT_TEXTBOX_MARGINS, DEFAULT_TEXTBOX_WIDTH } from "../../layout-engine/types";
import { resolveParagraphMarkFormatting } from "./paragraphMarkFormatting";
import {
  expectParagraphAttrs,
  expectTableAttrs,
  expectTableCellAttrs,
  expectTableRowAttrs,
  expectTextBoxAttrs,
} from "../../prosemirror/attrs";
import { textBoxHostParagraph } from "../../prosemirror/textBoxHostParagraph";
import type { TableAttrs } from "../../prosemirror/schema/nodes";
import type { Theme } from "../../types/document";
import { normalizeShapeTextAnchor } from "../../types/documentEnumValues";
import { groupParagraphFrames } from "./paragraphFrames";
import { twipsToPixels, nextBlockId } from "./flowConversionShared";
import type { FlowConversionOptions } from "./flowConversionShared";
import type { PageBreakRunProjection } from "./paragraphRuns";
import { convertBorderSpecToLayout } from "./flowBorders";
import type { ConvertibleBorder } from "./flowBorders";
import { convertParagraphAttrs } from "./paragraphAttrs";
import { convertParagraph } from "./paragraphConversion";
import { countPageBreakRuns, hasSingleLeadingProjectedPageBreak } from "./pageBreakSplitting";
import { inlineEffectExtentPx, resolveLinearGradientFill } from "./textBoxFill";

const TEXT_BOX_ANCHOR_BLOCK_ID = Symbol.for("stll.textBoxAnchorBlockId");
const DEFAULT_TABLE_CELL_MARGIN_TWIPS = {
  top: 0,
  right: 108,
  bottom: 0,
  left: 108,
} as const;
type TablePaddingSide = keyof typeof DEFAULT_TABLE_CELL_MARGIN_TWIPS;

/**
 * Extract cell borders from ProseMirror attributes.
 * Borders are full BorderSpec objects with style/size/color.
 */
function extractCellBorders(
  borders: Record<string, ConvertibleBorder> | null | undefined,
  theme?: Theme | null,
): CellBorders | undefined {
  if (!borders) {
    return undefined;
  }

  const result: CellBorders = {};
  const sides = ["top", "bottom", "left", "right"] as const;

  for (const side of sides) {
    const border = borders[side];
    const converted = border ? convertBorderSpecToLayout(border, theme) : undefined;
    if (!converted) {
      result[side] = { width: 0, style: "none" };
      continue;
    }

    result[side] = border?.size === 0 ? { ...converted, width: 0 } : converted;
  }

  const diagonalSides = ["topLeftToBottomRight", "topRightToBottomLeft"] as const;
  for (const side of diagonalSides) {
    const border = borders[side];
    const converted = border ? convertBorderSpecToLayout(border, theme) : undefined;
    if (converted) {
      result[side] = converted;
    }
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

/**
 * Convert a table cell node.
 */
type ConvertedTableCell = {
  cell: TableCell;
  breakBefore?: "page";
};

function convertTableCell(
  node: PMNode,
  startPos: number,
  options: FlowConversionOptions,
  tableCellMargins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  },
): ConvertedTableCell {
  const blocks: FlowBlock[] = [];
  let offset = startPos + 1; // +1 for opening tag
  const authoredPageBreakPosition = options.firstPageBreakRunPosition(node);
  // The break belongs to the cell's opening paragraph; what follows it in the
  // cell rides along, because the row moves as a unit.
  // An interior break, or a second one, needs table-fragment ownership, which
  // cell-local flow cannot model. The cell is then laid out whole: the run is
  // still projected and still saved, and refusing here would stop a document
  // Word opens from being laid out or exported at all.
  const leadingParagraph =
    node.firstChild?.type.name === "paragraph" &&
    countPageBreakRuns(node) === 1 &&
    options.firstPageBreakRunPosition(node.firstChild) === authoredPageBreakPosition
      ? node.firstChild
      : undefined;
  const pageBreaks: PageBreakRunProjection[] = [];

  const convertCellChild = (child: PMNode, childStart: number): void => {
    if (child.type.name === "paragraph") {
      const block = convertParagraph(
        child,
        childStart,
        options,
        child === leadingParagraph ? pageBreaks : undefined,
      );
      blocks.push(block);
    } else if (child.type.name === "table") {
      blocks.push(convertTable(child, childStart, options));
    } else if (child.type.name === "textBox") {
      blocks.push(convertTextBoxNode(child, childStart, options));
    } else if (child.type.name === "blockSdt") {
      // A `w:sdt` among a cell's block content wraps blocks of that cell:
      // what its `w:sdtContent` holds lays out as the cell's own blocks.
      let sdtChildStart = childStart + 1;
      // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
      child.forEach((sdtChild) => {
        convertCellChild(sdtChild, sdtChildStart);
        sdtChildStart += sdtChild.nodeSize;
      });
    }
  };
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    convertCellChild(child, offset);
    offset += child.nodeSize;
  });

  // A table cell whose final block is a nested table needs a trailing paragraph
  // as its editable cell-end marker. Keep that marker as a zero-height anchor.
  // Empty paragraphs after ordinary prose are authored content and retain their
  // normal line height.
  const trailingBlock = blocks.at(-1);
  const precedingBlock = blocks.at(-2);
  if (
    precedingBlock?.kind === "table" &&
    trailingBlock?.kind === "paragraph" &&
    trailingBlock.runs.every((run) => run.kind === "text" && run.text.length === 0)
  ) {
    trailingBlock.attrs = { ...trailingBlock.attrs, suppressEmptyParagraphHeight: true };
  }

  const attrs = expectTableCellAttrs(node);
  if (
    attrs.hideMark &&
    trailingBlock?.kind === "paragraph" &&
    trailingBlock.runs.every((run) => run.kind === "text" && run.text.length === 0)
  ) {
    trailingBlock.attrs = { ...trailingBlock.attrs, suppressEmptyParagraphHeight: true };
  }

  // Convert cell margins (twips) to pixel padding
  // OOXML TableNormal defaults: top=0, bottom=0, left=108 twips (~7px), right=108 twips (~7px)
  const margins = attrs.margins;
  const resolvePaddingSide = (
    side: TablePaddingSide,
    cellTwips: number | undefined,
    tableTwips: number | undefined,
  ): number => {
    if (cellTwips !== undefined) {
      return twipsToPixels(cellTwips);
    }
    if (tableTwips !== undefined) {
      return twipsToPixels(tableTwips);
    }
    return twipsToPixels(DEFAULT_TABLE_CELL_MARGIN_TWIPS[side]);
  };
  const padding = {
    top: resolvePaddingSide("top", margins?.top, tableCellMargins?.top),
    right: resolvePaddingSide("right", margins?.right, tableCellMargins?.right),
    bottom: resolvePaddingSide("bottom", margins?.bottom, tableCellMargins?.bottom),
    left: resolvePaddingSide("left", margins?.left, tableCellMargins?.left),
  };

  const cell: TableCell = {
    id: nextBlockId(),
    blocks: groupParagraphFrames(blocks, nextBlockId),
    colSpan: attrs.colspan,
    rowSpan: attrs.rowspan,
    padding,
  };
  if (attrs.width) {
    cell.width = twipsToPixels(attrs.width);
  }
  if (attrs.verticalAlign) {
    cell.verticalAlign = attrs.verticalAlign;
  }
  if (attrs.textDirection) {
    cell.textDirection = attrs.textDirection;
  }
  if (attrs.backgroundColor) {
    cell.background = `#${attrs.backgroundColor}`;
  }
  const cellBorders = extractCellBorders(attrs.borders, options.theme);
  if (cellBorders) {
    cell.borders = cellBorders;
  }
  if (attrs.noWrap) {
    cell.noWrap = true;
  }
  // A break the projection suppressed, a deleted one above all, is no boundary.
  if (authoredPageBreakPosition === undefined || pageBreaks.length === 0) {
    return { cell };
  }
  const paragraph = blocks.at(0);
  if (
    paragraph?.kind !== "paragraph" ||
    !hasSingleLeadingProjectedPageBreak(paragraph.runs, pageBreaks)
  ) {
    // An interior break needs table-fragment ownership, which cell-local flow
    // cannot model. The row stays whole and the run still saves; refusing here
    // would stop a document Word opens from being laid out at all.
    return { cell };
  }
  return { cell, breakBefore: "page" };
}

/**
 * Resolve an `ST_JcTable` placement to the flow engine's logical alignment.
 *
 * `start` and `end` name an edge of the writing direction, not a side of the
 * page, the same distinction `ST_Jc` draws for a paragraph. The flow engine
 * applies `w:bidiVisual` when it resolves that logical alignment to a physical
 * side, so the bridge must not mirror it first.
 */
const resolveTablePlacement = (
  placement: NonNullable<TableAttrs["justification"]>,
): NonNullable<TableBlock["justification"]> => {
  switch (placement) {
    case "start":
      return "left";
    case "end":
      return "right";
    case "left":
    case "center":
    case "right":
      return placement;
    default:
      return placement satisfies never;
  }
};

/**
 * Convert a table row node.
 */
type TableRowConversionContext = {
  cellMargins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
};

function convertTableRow(
  node: PMNode,
  startPos: number,
  options: FlowConversionOptions,
  { cellMargins }: TableRowConversionContext,
): TableRow {
  const cells: TableCell[] = [];
  let offset = startPos + 1; // +1 for opening tag
  let breakBefore: TableRow["breakBefore"];

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    if (child.type.name === "tableCell" || child.type.name === "tableHeader") {
      const attrs = expectTableCellAttrs(child);
      if (!attrs._omittedGridSlot) {
        const converted = convertTableCell(child, offset, options, cellMargins);
        if (converted.breakBefore !== undefined) {
          breakBefore = converted.breakBefore;
        }
        cells.push(converted.cell);
      }
    }
    offset += child.nodeSize;
  });

  const attrs = expectTableRowAttrs(node);
  const row: TableRow = {
    id: nextBlockId(),
    cells,
  };
  if (attrs._originalFormatting?.gridBefore) {
    row.gridBefore = attrs._originalFormatting.gridBefore;
  }
  if (attrs._originalFormatting?.gridAfter) {
    row.gridAfter = attrs._originalFormatting.gridAfter;
  }
  if (attrs.height) {
    row.height = twipsToPixels(attrs.height);
  }
  if (attrs.heightRule) {
    row.heightRule = attrs.heightRule;
  }
  if (attrs.isHeader) {
    row.isHeader = attrs.isHeader;
  }
  if (attrs._originalFormatting?.cantSplit) {
    row.cantSplit = true;
  }
  if (breakBefore !== undefined) {
    row.breakBefore = breakBefore;
  }
  if (attrs.hidden) {
    row.hidden = attrs.hidden;
  }
  const effectiveJustification =
    attrs._originalFormatting?.justification ?? attrs._resolvedJustification;
  if (effectiveJustification) {
    row.justification = resolveTablePlacement(effectiveJustification);
  }
  return row;
}

/**
 * Convert a table node to a TableBlock.
 */
export function convertTable(
  node: PMNode,
  startPos: number,
  options: FlowConversionOptions,
): TableBlock {
  const rows: TableRow[] = [];
  let offset = startPos + 1; // +1 for opening tag
  const attrs = expectTableAttrs(node);
  const rightToLeft = (attrs._resolvedBidi ?? attrs._originalFormatting?.bidi) === true;
  const rowContext: TableRowConversionContext = {
    ...(attrs.cellMargins === undefined ? {} : { cellMargins: attrs.cellMargins }),
  };

  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child) => {
    if (child.type.name === "tableRow") {
      rows.push(convertTableRow(child, offset, options, rowContext));
    }
    offset += child.nodeSize;
  });

  // Extract columnWidths from node attributes and convert from twips to pixels
  const columnWidthsTwips = attrs.columnWidths;
  let columnWidths = columnWidthsTwips?.map(twipsToPixels);

  const width = attrs.width;
  const widthType = attrs.widthType;

  // Fallback: compute column widths from first row cell widths if table attr is missing
  if (!columnWidths && rows.length > 0) {
    // SAFETY: rows.length > 0 verified by condition above
    const firstRow = rows[0]!;
    const cellWidths = firstRow.cells.map((cell) => cell.width);
    // Only use if all cells have widths defined
    if (cellWidths.every((w) => w !== undefined && w > 0)) {
      columnWidths = cellWidths as number[];
    }
  }

  // Keep authored justification separate in ProseMirror so style-derived
  // placement never becomes direct formatting on save.
  const authoredJustification = attrs.justification ?? attrs._resolvedJustification;
  const justification =
    authoredJustification === undefined ? undefined : resolveTablePlacement(authoredJustification);

  // Extract table indent + RTL column order from _originalFormatting
  // (w:tblInd, w:bidiVisual). bidiVisual is import-only — folio has no UI to
  // toggle it — so reading the preserved formatting is sufficient
  // (eigenpal/docx-editor#940).
  const originalFormatting = attrs._originalFormatting;
  const resolvedIndent = attrs._resolvedIndent;
  const effectiveIndent = resolvedIndent ?? originalFormatting?.indent;
  const indentPx =
    effectiveIndent?.value !== undefined && effectiveIndent?.type === "dxa"
      ? twipsToPixels(effectiveIndent.value)
      : undefined;
  // An indent measurement folio cannot apply must not half-apply: pairing the
  // text-edge compensation with a dropped `w:tblInd` would shift the table by
  // the leading cell margin alone, which no indent asked for.
  const dropsAuthoredIndent = effectiveIndent?.value !== undefined && indentPx === undefined;

  const floating = attrs.floating as
    | {
        horzAnchor?: "margin" | "page" | "text";
        vertAnchor?: "margin" | "page" | "text";
        tblpX?: number;
        tblpXSpec?: "left" | "center" | "right" | "inside" | "outside";
        tblpY?: number;
        tblpYSpec?: "top" | "center" | "bottom" | "inside" | "outside" | "inline";
        topFromText?: number;
        bottomFromText?: number;
        leftFromText?: number;
        rightFromText?: number;
      }
    | undefined;

  let floatingPx: FloatingTablePosition | undefined;
  if (floating) {
    const fp: FloatingTablePosition = {};
    if (floating.horzAnchor) {
      fp.horzAnchor = floating.horzAnchor;
    }
    if (floating.vertAnchor) {
      fp.vertAnchor = floating.vertAnchor;
    }
    if (floating.tblpX !== undefined) {
      fp.tblpX = twipsToPixels(floating.tblpX);
    }
    if (floating.tblpXSpec) {
      fp.tblpXSpec = floating.tblpXSpec;
    }
    if (floating.tblpY !== undefined) {
      fp.tblpY = twipsToPixels(floating.tblpY);
    }
    if (floating.tblpYSpec) {
      fp.tblpYSpec = floating.tblpYSpec;
    }
    if (floating.topFromText !== undefined) {
      fp.topFromText = twipsToPixels(floating.topFromText);
    }
    if (floating.bottomFromText !== undefined) {
      fp.bottomFromText = twipsToPixels(floating.bottomFromText);
    }
    if (floating.leftFromText !== undefined) {
      fp.leftFromText = twipsToPixels(floating.leftFromText);
    }
    if (floating.rightFromText !== undefined) {
      fp.rightFromText = twipsToPixels(floating.rightFromText);
    }
    floatingPx = fp;
  }

  const tableBlock: TableBlock = {
    kind: "table",
    id: nextBlockId(),
    rows,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
  if (columnWidths) {
    tableBlock.columnWidths = columnWidths;
  }
  if (width !== undefined) {
    tableBlock.width = width;
  }
  if (widthType !== undefined) {
    tableBlock.widthType = widthType;
  }
  if (originalFormatting?.layout !== undefined) {
    tableBlock.layout = originalFormatting.layout;
  }
  if (justification) {
    tableBlock.justification = justification;
  }
  if (indentPx !== undefined) {
    tableBlock.indent = indentPx;
  }
  // `w:tblCellSpacing` is only defined as an absolute width; `pct` and `auto`
  // have no length to resolve against.
  const cellSpacing = originalFormatting?.cellSpacing;
  if (cellSpacing?.type === "dxa" && cellSpacing.value > 0) {
    tableBlock.cellSpacing = twipsToPixels(cellSpacing.value);
  }
  if (options.tableIndentCompatibility && !dropsAuthoredIndent) {
    tableBlock.indentCompatibility = options.tableIndentCompatibility;
  }
  if (floatingPx) {
    tableBlock.floating = floatingPx;
  }
  if (rightToLeft) {
    tableBlock.bidi = true;
  }
  return tableBlock;
}

/** Convert a textBox PM node to a TextBoxBlock. */
export function convertTextBoxNode(
  node: PMNode,
  startPos: number,
  opts: FlowConversionOptions,
): TextBoxBlock {
  // A break inside a text box paginates nothing: the box is placed as a unit.
  // The run is still projected and still saved; refusing here would stop a
  // document Word opens from being laid out or exported at all.
  const attrs = expectTextBoxAttrs(node);
  const contentBlocks: (ParagraphBlock | TableBlock)[] = [];

  // Convert child blocks inside the text box
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  node.forEach((child, offset) => {
    const childPos = startPos + 1 + offset;
    if (child.type.name === "paragraph") {
      contentBlocks.push(convertParagraph(child, childPos, opts));
      return;
    }
    if (child.type.name === "table") {
      contentBlocks.push(convertTable(child, childPos, opts));
    }
  });

  const textBox: TextBoxBlock = {
    kind: "textBox",
    id: nextBlockId(),
    width: attrs.width ?? DEFAULT_TEXTBOX_WIDTH,
    margins: {
      top: attrs.marginTop ?? DEFAULT_TEXTBOX_MARGINS.top,
      bottom: attrs.marginBottom ?? DEFAULT_TEXTBOX_MARGINS.bottom,
      left: attrs.marginLeft ?? DEFAULT_TEXTBOX_MARGINS.left,
      right: attrs.marginRight ?? DEFAULT_TEXTBOX_MARGINS.right,
    },
    content: contentBlocks,
    pmStart: startPos,
    pmEnd: startPos + node.nodeSize,
  };
  const verticalAlign = normalizeShapeTextAnchor(attrs.verticalAlign);
  if (verticalAlign !== undefined) {
    textBox.verticalAlign = verticalAlign;
  }
  if (attrs.height !== undefined) {
    textBox.height = attrs.height;
  }
  if (attrs.autoFit !== undefined) {
    textBox.autoFit = attrs.autoFit;
  }
  if (attrs.textWrap !== undefined) {
    textBox.textWrap = attrs.textWrap;
  }
  if (attrs.fillColor !== undefined) {
    textBox.fillColor = attrs.fillColor;
  } else {
    const fillGradient = resolveLinearGradientFill(attrs.gradientFill, opts.theme);
    if (fillGradient !== undefined) {
      textBox.fillGradient = fillGradient;
    }
  }
  if (attrs.outlineWidth !== undefined) {
    textBox.outlineWidth = attrs.outlineWidth;
  }
  if (attrs.outlineColor !== undefined) {
    textBox.outlineColor = attrs.outlineColor;
  }
  if (attrs.outlineStyle !== undefined) {
    textBox.outlineStyle = attrs.outlineStyle;
  }
  if (attrs.transform !== undefined) {
    textBox.transform = attrs.transform;
  }
  // Carry anchored-textbox wrap attributes through so the page renderer can
  // build exclusion rects (eigenpal #474).
  if (attrs.displayMode !== undefined) {
    textBox.displayMode = attrs.displayMode;
  }
  if (attrs.cssFloat !== undefined) {
    textBox.cssFloat = attrs.cssFloat;
  }
  if (attrs.wrapType !== undefined) {
    textBox.wrapType = attrs.wrapType;
  }
  if (attrs.wrapText !== undefined) {
    textBox.wrapText = attrs.wrapText;
  }
  if (attrs.distTop !== undefined) {
    textBox.distTop = attrs.distTop;
  }
  if (attrs.distBottom !== undefined) {
    textBox.distBottom = attrs.distBottom;
  }
  if (attrs.distLeft !== undefined) {
    textBox.distLeft = attrs.distLeft;
  }
  if (attrs.distRight !== undefined) {
    textBox.distRight = attrs.distRight;
  }
  if (attrs.position !== undefined) {
    textBox.position = attrs.position;
  }
  const effectExtent = inlineEffectExtentPx(attrs.wrapType, attrs.wrapEffectExtentSlots?.drawing);
  if (effectExtent !== undefined) {
    textBox.effectExtent = effectExtent;
  }
  const host = textBoxHostParagraph(node);
  if (host && attrs.position === undefined && (attrs.wrapType ?? "inline") === "inline") {
    const hostAttrs = expectParagraphAttrs(host);
    textBox.hostParagraph = convertParagraphAttrs(hostAttrs, {
      theme: opts.theme,
      fontAlternates: opts.fontAlternates,
      listCounterStreams: opts.listCounterStreams,
      defaultTabStopTwips: opts.defaultTabStopTwips,
      paragraphMarkFormatting: () => resolveParagraphMarkFormatting(hostAttrs, opts.styleResolver),
    });
  }
  if (attrs._docxGroupId !== undefined) {
    setTextBoxGroupId(textBox, attrs._docxGroupId);
  }
  const anchorBlockId = attrs._docxAnchorId
    ? opts.textBoxAnchorBlockIds.get(attrs._docxAnchorId)
    : undefined;
  if (anchorBlockId !== undefined) {
    Reflect.set(textBox, TEXT_BOX_ANCHOR_BLOCK_ID, anchorBlockId);
  }
  return textBox;
}
