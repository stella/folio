/**
 * Footnote Layout Utilities
 *
 * Handles scanning for footnote references, mapping them to pages,
 * converting footnote content to measurable FlowBlocks, and computing
 * per-page footnote area heights for layout space reservation.
 */

import type {
  FlowBlock,
  Measure,
  Page,
  ParagraphBlock,
  Run,
  TableBlock,
  TableCellMeasure,
  TableMeasure,
  TableRowMeasure,
  TextRun,
  FootnoteContent,
} from "../../layout-engine/types";
import {
  DEFAULT_TEXTBOX_MARGINS as TEXTBOX_MARGINS,
  FOOTNOTE_ENTRY_MARGIN_BOTTOM,
  FOOTNOTE_FALLBACK_LINE_HEIGHT,
  FOOTNOTE_SEPARATOR_HEIGHT,
} from "../../layout-engine/types";
import { footnoteToProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Footnote, StyleDefinitions, Theme } from "../../types/document";
import { measureParagraph } from "../engine/measuring";
import { layoutTextBoxContent } from "../../layout-engine/measure/textBoxParagraphLayout";
import { toFlowBlocks } from "./toFlowBlocks";
import type { ToFlowBlocksOptions } from "./toFlowBlocks";

// Re-exported for back-compat with existing callers that imported the
// constants from this module before they moved to `layout-engine/types`.
export { FOOTNOTE_SEPARATOR_HEIGHT, FOOTNOTE_ENTRY_MARGIN_BOTTOM, FOOTNOTE_FALLBACK_LINE_HEIGHT };

/** Default footnote font size in points */
const FOOTNOTE_FONT_SIZE = 8;

export type MeasureBlocksFn = (blocks: FlowBlock[], contentWidth: number) => Measure[];

export type ConvertFootnoteOptions = {
  styles?: StyleDefinitions | null;
  theme?: Theme | null;
  measureBlocks?: MeasureBlocksFn;
  /** Document-wide `w:defaultTabStop` in twips — forwarded to toFlowBlocks. */
  defaultTabStopTwips?: number;
  lineBreakRules?: ToFlowBlocksOptions["lineBreakRules"];
  justificationCompatibility?: ToFlowBlocksOptions["justificationCompatibility"];
  automaticHyphenation?: ToFlowBlocksOptions["automaticHyphenation"];
};

// ============================================================================
// 1. Scan FlowBlocks for footnote references
// ============================================================================

/**
 * Scan FlowBlocks for runs with footnoteRefId set.
 * Returns a list of { footnoteId, pmPos } in document order.
 *
 * Recurses into table cells and text-box content so footnote references
 * nested in tables (incl. tables-within-cells) and text boxes still reach
 * the page-assignment step. Without this walk, inline footnote markers
 * render inside the body but never get an entry in the per-page footnote
 * area.
 */
export function collectFootnoteRefs(blocks: FlowBlock[]): { footnoteId: number; pmPos: number }[] {
  return collectNoteRefs(blocks, "footnoteRefId").map(({ noteId, pmPos }) => ({
    footnoteId: noteId,
    pmPos,
  }));
}

/**
 * Scan FlowBlocks for runs with endnoteRefId set, in document order.
 * Same recursive walk as `collectFootnoteRefs`.
 */
export function collectEndnoteRefs(blocks: FlowBlock[]): { endnoteId: number; pmPos: number }[] {
  return collectNoteRefs(blocks, "endnoteRefId").map(({ noteId, pmPos }) => ({
    endnoteId: noteId,
    pmPos,
  }));
}

function collectNoteRefs(
  blocks: FlowBlock[],
  idKey: "footnoteRefId" | "endnoteRefId",
): { noteId: number; pmPos: number }[] {
  const refs: { noteId: number; pmPos: number }[] = [];

  const walk = (containerBlocks: FlowBlock[]): void => {
    for (const block of containerBlocks) {
      if (block.kind === "paragraph") {
        for (const run of block.runs) {
          if (run.kind !== "text") {
            continue;
          }
          const noteId = run[idKey];
          if (noteId !== undefined) {
            refs.push({ noteId, pmPos: run.pmStart ?? 0 });
          }
        }
      } else if (block.kind === "table") {
        for (const row of block.rows) {
          for (const cell of row.cells) {
            walk(cell.blocks);
          }
        }
      } else if (block.kind === "textBox") {
        walk(block.content);
      }
    }
  };

  walk(blocks);
  return refs;
}

// ============================================================================
// 1.5 Sequential display numbers + body marker remap
// ============================================================================

type NumberableNote = {
  id: number;
  noteType?: Footnote["noteType"];
};

/**
 * Assign sequential display numbers (1, 2, 3, …) to notes in order of first
 * reference. Word numbers footnotes/endnotes by reference order, not by their
 * `w:id` values, which may be non-contiguous or out of document order. Only
 * `normal` notes are numbered; separator/continuation notes and refs to
 * missing notes consume no number.
 */
export function computeNoteDisplayNumbers(
  notes: readonly NumberableNote[],
  refNoteIds: readonly number[],
): Map<number, number> {
  const numberable = new Set<number>();
  for (const note of notes) {
    if (note.noteType === "normal") {
      numberable.add(note.id);
    }
  }

  const displayNumbers = new Map<number, number>();
  let nextNumber = 1;
  for (const noteId of refNoteIds) {
    if (displayNumbers.has(noteId) || !numberable.has(noteId)) {
      continue;
    }
    displayNumbers.set(noteId, nextNumber);
    nextNumber++;
  }
  return displayNumbers;
}

export type NoteDisplayNumberMaps = {
  /** footnote `w:id` → sequential display number */
  footnoteNumbers?: ReadonlyMap<number, number>;
  /** endnote `w:id` → sequential display number */
  endnoteNumbers?: ReadonlyMap<number, number>;
};

/**
 * Rewrite body reference-marker run text from the raw `w:id` (which the PM
 * doc stores as the marker text; the save path serializes from the mark
 * attrs, never from this text) to the sequential display number, so the
 * inline marker matches the number rendered in the note area. Must run
 * before measurement so marker widths match the painted digits. Untouched
 * blocks/runs are returned by reference so the painter's fingerprinting and
 * the incremental measure path see them unchanged. Runs keep their PM range,
 * so click-to-position still resolves into the marker (same contract as the
 * template preview substitution).
 */
export function remapNoteMarkerText(blocks: FlowBlock[], maps: NoteDisplayNumberMaps): FlowBlock[] {
  if ((maps.footnoteNumbers?.size ?? 0) === 0 && (maps.endnoteNumbers?.size ?? 0) === 0) {
    return blocks;
  }

  let changed = false;
  const next = blocks.map((block) => {
    const remapped = remapNoteMarkerBlock(block, maps);
    changed ||= remapped !== block;
    return remapped;
  });
  return changed ? next : blocks;
}

function remapNoteMarkerBlock(block: FlowBlock, maps: NoteDisplayNumberMaps): FlowBlock {
  if (block.kind === "paragraph") {
    return remapNoteMarkerParagraph(block, maps);
  }
  if (block.kind === "table") {
    return remapNoteMarkerTable(block, maps);
  }
  if (block.kind === "textBox") {
    let boxChanged = false;
    const content = block.content.map((contentBlock) => {
      const remapped =
        contentBlock.kind === "table"
          ? remapNoteMarkerTable(contentBlock, maps)
          : remapNoteMarkerParagraph(contentBlock, maps);
      boxChanged ||= remapped !== contentBlock;
      return remapped;
    });
    return boxChanged ? { ...block, content } : block;
  }
  return block;
}

function remapNoteMarkerTable(block: TableBlock, maps: NoteDisplayNumberMaps): TableBlock {
  let tableChanged = false;
  const rows = block.rows.map((row) => {
    const cells = row.cells.map((cell) => {
      let cellChanged = false;
      const cellBlocks = cell.blocks.map((cellBlock) => {
        const remapped = remapNoteMarkerBlock(cellBlock, maps);
        cellChanged ||= remapped !== cellBlock;
        return remapped;
      });
      return cellChanged ? { ...cell, blocks: cellBlocks } : cell;
    });
    const rowChanged = cells.some((cell, index) => cell !== row.cells[index]);
    tableChanged ||= rowChanged;
    return rowChanged ? { ...row, cells } : row;
  });
  return tableChanged ? { ...block, rows } : block;
}

function remapNoteMarkerParagraph(
  block: ParagraphBlock,
  maps: NoteDisplayNumberMaps,
): ParagraphBlock {
  let changed = false;
  const runs = block.runs.map((run) => {
    const remapped = remapNoteMarkerRun(run, maps);
    changed ||= remapped !== run;
    return remapped;
  });
  return changed ? { ...block, runs } : block;
}

function remapNoteMarkerRun(run: Run, maps: NoteDisplayNumberMaps): Run {
  if (run.kind !== "text") {
    return run;
  }
  const displayNumber = getRunDisplayNumber(run, maps);
  if (displayNumber === undefined) {
    return run;
  }
  const text = String(displayNumber);
  return run.text === text ? run : { ...run, text };
}

function getRunDisplayNumber(run: TextRun, maps: NoteDisplayNumberMaps): number | undefined {
  if (run.footnoteRefId !== undefined) {
    return maps.footnoteNumbers?.get(run.footnoteRefId);
  }
  if (run.endnoteRefId !== undefined) {
    return maps.endnoteNumbers?.get(run.endnoteRefId);
  }
  return undefined;
}

// ============================================================================
// 2. Map footnote references to pages
// ============================================================================

/**
 * After layout, determine which footnotes appear on which pages.
 * Checks each page's fragments to see if any footnoteRef PM positions fall within.
 *
 * Returns Map<pageNumber, footnoteId[]> in document order.
 */
export function mapFootnotesToPages(
  pages: Page[],
  footnoteRefs: { footnoteId: number; pmPos: number }[],
): Map<number, number[]> {
  const pageFootnotes = new Map<number, number[]>();

  if (footnoteRefs.length === 0) {
    return pageFootnotes;
  }

  // For each footnote ref, find which page it lands on
  for (const ref of footnoteRefs) {
    for (const page of pages) {
      let found = false;
      for (const fragment of page.fragments) {
        const pmStart = fragment.pmStart ?? -1;
        const pmEnd = fragment.pmEnd ?? -1;
        if (pmStart >= 0 && pmEnd >= 0 && ref.pmPos >= pmStart && ref.pmPos < pmEnd) {
          const existing = pageFootnotes.get(page.number) ?? [];
          // Avoid duplicates (same footnote shouldn't appear twice on same page)
          if (!existing.includes(ref.footnoteId)) {
            existing.push(ref.footnoteId);
          }
          pageFootnotes.set(page.number, existing);
          found = true;
          break;
        }
      }
      if (found) {
        break;
      }
    }
  }

  return pageFootnotes;
}

// ============================================================================
// 3. Convert footnote content to FlowBlocks + Measures
// ============================================================================

/**
 * Convert a Footnote's content paragraphs to FlowBlocks suitable for rendering.
 * Prepends the display number to the first run of the first paragraph.
 */
export function convertFootnoteToContent(
  footnote: Footnote,
  displayNumber: number,
  contentWidth: number,
  options: ConvertFootnoteOptions = {},
): FootnoteContent {
  const proseOptions: Parameters<typeof footnoteToProseDoc>[1] = {};
  if (options.styles) {
    proseOptions.styles = options.styles;
  }
  if (options.theme !== undefined) {
    proseOptions.theme = options.theme;
  }
  const pmDoc = footnoteToProseDoc(footnote.content, proseOptions);
  const flowOptions: Parameters<typeof toFlowBlocks>[1] = {};
  if (options.theme !== undefined) {
    flowOptions.theme = options.theme;
  }
  if (options.defaultTabStopTwips !== undefined) {
    flowOptions.defaultTabStopTwips = options.defaultTabStopTwips;
  }
  if (options.lineBreakRules) {
    flowOptions.lineBreakRules = options.lineBreakRules;
  }
  if (options.justificationCompatibility) {
    flowOptions.justificationCompatibility = options.justificationCompatibility;
  }
  if (options.automaticHyphenation) {
    flowOptions.automaticHyphenation = options.automaticHyphenation;
  }
  const blocks = applyFootnotePresentation(toFlowBlocks(pmDoc, flowOptions), displayNumber);

  const measures = options.measureBlocks
    ? options.measureBlocks(blocks, contentWidth)
    : measureFootnoteBlocks(blocks, contentWidth);

  let totalHeight = 0;
  for (const measure of measures) {
    if (measure.kind === "paragraph") {
      totalHeight += measure.totalHeight;
    } else if (measure.kind === "table") {
      totalHeight += measure.totalHeight;
    } else if (measure.kind === "image" || measure.kind === "textBox") {
      totalHeight += measure.height;
    }
  }

  return {
    id: footnote.id,
    displayNumber,
    blocks,
    measures,
    height: totalHeight,
  };
}

function measureFootnoteBlocks(blocks: FlowBlock[], contentWidth: number): Measure[] {
  return blocks.map((block) => measureFootnoteBlock(block, contentWidth));
}

function measureFootnoteBlock(block: FlowBlock, contentWidth: number): Measure {
  switch (block.kind) {
    case "paragraph":
      return measureParagraph(block, contentWidth);

    case "table":
      return measureFootnoteTable(block, contentWidth);

    case "image":
      return {
        kind: "image",
        width: block.width,
        height: block.height,
      };

    case "textBox": {
      const margins = block.margins ?? TEXTBOX_MARGINS;
      const width = block.width;
      const innerWidth = Math.max(1, width - margins.left - margins.right);
      const innerMeasures = block.content.map((contentBlock) => {
        if (contentBlock.kind === "table") {
          return measureFootnoteTable(contentBlock, innerWidth);
        }
        return measureParagraph(contentBlock, innerWidth);
      });
      const contentHeight = layoutTextBoxContent(block.content, innerMeasures).totalHeight;

      return {
        kind: "textBox",
        width,
        height: block.height ?? contentHeight + margins.top + margins.bottom,
        innerMeasures,
      };
    }

    case "pageBreak":
      return { kind: "pageBreak" };

    case "columnBreak":
      return { kind: "columnBreak" };

    case "sectionBreak":
      return { kind: "sectionBreak" };

    default: {
      const exhaustive: never = block;
      return exhaustive;
    }
  }
}

function resolveFootnoteTableWidth(
  width: number | undefined,
  widthType: string | undefined,
  contentWidth: number,
): number | undefined {
  if (!width) {
    return undefined;
  }
  if (widthType === "pct") {
    return (contentWidth * width) / 5000;
  }
  if (widthType === "dxa" || !widthType || widthType === "auto") {
    return (width / 1440) * 96;
  }
  return undefined;
}

function measureFootnoteTable(tableBlock: TableBlock, contentWidth: number): TableMeasure {
  let columnWidths = tableBlock.columnWidths ?? [];
  const explicitWidth = resolveFootnoteTableWidth(
    tableBlock.width,
    tableBlock.widthType,
    contentWidth,
  );

  if (columnWidths.length === 0) {
    const firstRow = tableBlock.rows.at(0);
    let colCount = 0;
    for (const cell of firstRow?.cells ?? []) {
      colCount += cell.colSpan ?? 1;
    }
    const totalWidth = explicitWidth ?? contentWidth;
    const equalWidth = totalWidth / Math.max(1, colCount);
    columnWidths = Array.from({ length: Math.max(1, colCount) }, () => equalWidth);
  } else if (explicitWidth) {
    const totalWidth = sumColumnWidths(columnWidths);
    if (totalWidth > 0 && Math.abs(totalWidth - explicitWidth) > 1) {
      const scale = explicitWidth / totalWidth;
      columnWidths = columnWidths.map((width) => width * scale);
    }
  }

  const rowSpanEnds: number[] = [];
  const rows: TableRowMeasure[] = tableBlock.rows.map((row, rowIndex) => {
    let columnIndex = 0;
    const cells: TableCellMeasure[] = row.cells.map((cell) => {
      while ((rowSpanEnds[columnIndex] ?? 0) > rowIndex) {
        columnIndex++;
      }

      const colSpan = cell.colSpan ?? 1;
      let cellWidth = 0;
      for (
        let offset = 0;
        offset < colSpan && columnIndex + offset < columnWidths.length;
        offset++
      ) {
        cellWidth += columnWidths[columnIndex + offset] ?? 0;
      }
      if (cellWidth === 0) {
        cellWidth = cell.width ?? 100;
      }
      columnIndex += colSpan;

      const padding = cell.padding ?? { top: 0, right: 7, bottom: 0, left: 7 };
      const innerWidth = Math.max(1, cellWidth - padding.left - padding.right);
      const blocks = measureFootnoteBlocks(cell.blocks, innerWidth);
      let height = padding.top + padding.bottom;
      for (const measure of blocks) {
        height += getMeasureHeight(measure);
      }

      const cellMeasure: TableCellMeasure = {
        blocks,
        width: cellWidth,
        height,
      };
      if (cell.colSpan !== undefined) {
        cellMeasure.colSpan = cell.colSpan;
      }
      if (cell.rowSpan !== undefined) {
        cellMeasure.rowSpan = cell.rowSpan;
      }
      if ((cell.rowSpan ?? 1) > 1) {
        const rowSpanEnd = rowIndex + (cell.rowSpan ?? 1);
        for (let offset = 0; offset < colSpan; offset++) {
          rowSpanEnds[columnIndex - colSpan + offset] = rowSpanEnd;
        }
      }
      return cellMeasure;
    });

    let contentHeight = 0;
    for (const cell of cells) {
      contentHeight = Math.max(contentHeight, cell.height);
    }
    const explicitHeight = row.height;
    let height = contentHeight;
    if (explicitHeight !== undefined && row.heightRule === "exact") {
      height = explicitHeight;
    } else if (explicitHeight !== undefined) {
      height = Math.max(contentHeight, explicitHeight);
    }

    return { cells, height };
  });

  let totalHeight = 0;
  for (const row of rows) {
    totalHeight += row.height;
  }

  return {
    kind: "table",
    rows,
    columnWidths,
    totalWidth: sumColumnWidths(columnWidths) || explicitWidth || contentWidth,
    totalHeight,
  };
}

function sumColumnWidths(columnWidths: number[]): number {
  let total = 0;
  for (const width of columnWidths) {
    total += width;
  }
  return total;
}

function getMeasureHeight(measure: Measure): number {
  if (measure.kind === "paragraph" || measure.kind === "table") {
    return measure.totalHeight;
  }
  if (measure.kind === "image" || measure.kind === "textBox") {
    return measure.height;
  }
  return 0;
}

export function applyFootnotePresentation(blocks: FlowBlock[], displayNumber: number): FlowBlock[] {
  if (blocks.length === 0) {
    return [
      {
        kind: "paragraph",
        id: `fn-empty-${displayNumber}`,
        runs: [
          {
            kind: "text",
            text: `${displayNumber}  `,
            fontSize: FOOTNOTE_FONT_SIZE,
            superscript: true,
          },
        ],
      } satisfies ParagraphBlock,
    ];
  }

  const output = blocks.map(applyFootnoteBlockPresentation);

  const first = output[0];
  if (first?.kind === "paragraph") {
    output[0] = {
      ...first,
      runs: [createFootnoteNumberRun(displayNumber, first), ...first.runs],
    };
  } else {
    output.unshift({
      kind: "paragraph",
      id: `fn-number-${displayNumber}`,
      runs: [
        {
          kind: "text",
          text: `${displayNumber}  `,
          fontSize: FOOTNOTE_FONT_SIZE,
          superscript: true,
        },
      ],
    });
  }

  return output;
}

function createFootnoteNumberRun(displayNumber: number, paragraph: ParagraphBlock): TextRun {
  const firstTextRun = paragraph.runs.find((run) => run.kind === "text");
  const firstFormattedRun = paragraph.runs.find(
    (run) => run.kind === "text" || run.kind === "tab" || run.kind === "field",
  );
  const text = firstTextRun?.text.match(/^\s/u) ? `${displayNumber}` : `${displayNumber} `;
  const numberRun: TextRun = {
    kind: "text",
    text,
    fontSize: firstFormattedRun?.fontSize ?? paragraph.attrs?.defaultFontSize ?? FOOTNOTE_FONT_SIZE,
    superscript: true,
  };
  const fontFamily = firstFormattedRun?.fontFamily ?? paragraph.attrs?.defaultFontFamily;
  if (fontFamily) {
    numberRun.fontFamily = fontFamily;
  }
  return numberRun;
}

function applyFootnoteBlockPresentation(block: FlowBlock): FlowBlock {
  if (block.kind === "paragraph") {
    return applyFootnoteParagraphPresentation(block);
  }
  if (block.kind === "table") {
    return applyFootnoteTablePresentation(block);
  }
  if (block.kind === "textBox") {
    return {
      ...block,
      content: block.content.map((contentBlock) =>
        contentBlock.kind === "table"
          ? applyFootnoteTablePresentation(contentBlock)
          : applyFootnoteParagraphPresentation(contentBlock),
      ),
    };
  }
  return block;
}

function applyFootnoteTablePresentation(block: TableBlock): TableBlock {
  return {
    ...block,
    rows: block.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        blocks: cell.blocks.map(applyFootnoteBlockPresentation),
      })),
    })),
  };
}

function applyFootnoteParagraphPresentation(block: ParagraphBlock): ParagraphBlock {
  return {
    ...block,
    runs: block.runs.map(applyFootnoteRunPresentation),
  };
}

function applyFootnoteRunPresentation(run: Run): Run {
  if (
    (run.kind === "text" || run.kind === "tab" || run.kind === "field") &&
    run.fontSize === undefined
  ) {
    return { ...run, fontSize: FOOTNOTE_FONT_SIZE };
  }
  return run;
}

// ============================================================================
// 4. Build per-page footnote content and reserved heights
// ============================================================================

/**
 * Build footnote content for all footnotes referenced in the document.
 * Returns a Map<footnoteId, FootnoteContent>.
 */
export function buildFootnoteContentMap(
  footnotes: Footnote[],
  footnoteRefs: { footnoteId: number }[],
  contentWidth: number,
  options: ConvertFootnoteOptions = {},
): Map<number, FootnoteContent> {
  const contentMap = new Map<number, FootnoteContent>();
  const footnoteById = new Map<number, Footnote>();

  for (const fn of footnotes) {
    if (fn.noteType === "normal") {
      footnoteById.set(fn.id, fn);
    }
  }

  // Display numbers follow first-appearance order — the same map that
  // drives the body marker remap, so area and marker can never disagree.
  const displayNumbers = computeNoteDisplayNumbers(
    footnotes,
    footnoteRefs.map((ref) => ref.footnoteId),
  );

  for (const [footnoteId, displayNumber] of displayNumbers) {
    const footnote = footnoteById.get(footnoteId);
    if (!footnote) {
      continue;
    }
    contentMap.set(
      footnoteId,
      convertFootnoteToContent(footnote, displayNumber, contentWidth, options),
    );
  }

  return contentMap;
}

/**
 * Calculate per-page footnote reserved heights.
 * Returns Map<pageNumber, reservedHeight>.
 */
export function calculateFootnoteReservedHeights(
  pageFootnoteMap: Map<number, number[]>,
  footnoteContentMap: Map<number, { height: number }>,
): Map<number, number> {
  const reserved = new Map<number, number>();

  for (const [pageNumber, footnoteIds] of pageFootnoteMap) {
    let totalHeight = 0;

    for (const fnId of footnoteIds) {
      const content = footnoteContentMap.get(fnId);
      if (content) {
        totalHeight += content.height;
      }
    }

    if (totalHeight > 0) {
      // Add separator + any wrapper margin so the static reservation matches
      // what `renderFootnoteArea` actually paints. In Word-like rendering the
      // wrapper margin is zero; paragraph spacing inside each footnote content
      // carries the source DOCX spacing.
      totalHeight += FOOTNOTE_SEPARATOR_HEIGHT;
      totalHeight += footnoteIds.length * FOOTNOTE_ENTRY_MARGIN_BOTTOM;
      reserved.set(pageNumber, totalHeight);
    }
  }

  return reserved;
}
