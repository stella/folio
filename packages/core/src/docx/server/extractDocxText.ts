import type { DocxArchive } from "./boundedArchive";
import { loadDocxArchive } from "./boundedArchive";
import { escapeTableCell } from "../../markdown/escape";
import { parseRelationships, RELATIONSHIP_TYPES } from "../relsParser";
import {
  findAllDeep,
  findDeep,
  getAttribute,
  getAttributeByNamespaceUri,
  getLocalName,
  getNamespaceUri,
  getTextContent,
  parseXml,
  type XmlElement,
} from "../xmlParser";

const DOCUMENT_RELS_PATH = "word/_rels/document.xml.rels";
const WORDPROCESSINGML_NAMESPACES: ReadonlySet<string> = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);

/** Document part containing an extracted paragraph. */
export type DocxParagraphSource = "header" | "body" | "footer";

/**
 * Role of an emitted markdown table row.
 *
 * - `cells` — a `w:tr` rendered as a pipe row, including the first row when the
 *   table declares it as its header.
 * - `syntheticHeader` — the empty header row emitted for a table that declares
 *   no header row; GFM has no headerless table.
 * - `delimiter` — the `| --- |` line GFM requires under the header.
 */
export type DocxTableRowKind = "cells" | "syntheticHeader" | "delimiter";

/** Source paragraph inside a structured table cell. */
export type ExtractedDocxTableCellParagraph = {
  text: string;
  style?: string;
  bold?: boolean;
  fontSize?: number;
  alignment?: "left" | "center" | "right" | "both";
};

/** Source paragraphs contained by one physical table cell. */
export type ExtractedDocxTableCell = {
  /** Source `w:p` records in document order; empty padding cells have no entries. */
  paragraphs: readonly ExtractedDocxTableCellParagraph[];
};

/** Table membership of a paragraph whose `text` is a markdown table row. */
export type DocxTableRowPosition =
  | {
      /** 0-based index of the source `w:tbl`, in extraction order across all parts. */
      table: number;
      kind: "cells";
      /** Source cells aligned to the rendered GFM columns. */
      cells: readonly ExtractedDocxTableCell[];
    }
  | {
      /** 0-based index of the source `w:tbl`, in extraction order across all parts. */
      table: number;
      kind: "syntheticHeader" | "delimiter";
    };

/** Paragraph text and lightweight formatting metadata from a DOCX archive. */
export type ExtractedDocxParagraph = {
  index: number;
  text: string;
  source: DocxParagraphSource;
  style?: string;
  bold?: boolean;
  fontSize?: number;
  alignment?: "left" | "center" | "right" | "both";
  /**
   * Present only when `text` is a markdown table row rendered from a `w:tbl`,
   * absent for ordinary prose paragraphs. Consumers that join `text` across
   * paragraphs need no change; consumers that want to regroup a table's rows,
   * or drop the rows GFM forced into existence, can key off this.
   */
  tableRow?: DocxTableRowPosition;
};

/** Accepted-revision paragraph text extracted in deterministic part order. */
export type ExtractedDocxText = {
  paragraphs: ExtractedDocxParagraph[];
  charCount: number;
  view: "accepted";
};

type ParagraphProperties = Pick<ExtractedDocxParagraph, "style" | "alignment">;

type RunMetrics = {
  bold: boolean;
  fontSize?: number;
  chars: number;
};

const childElements = (element: XmlElement): XmlElement[] =>
  element.elements?.filter((child) => child.type === "element") ?? [];

const wordElementName = (element: XmlElement): string | null =>
  WORDPROCESSINGML_NAMESPACES.has(getNamespaceUri(element) ?? "")
    ? getLocalName(element.name)
    : null;

const findWordChild = (
  parent: XmlElement | null | undefined,
  localName: string,
): XmlElement | null => {
  if (!parent) {
    return null;
  }
  return childElements(parent).find((child) => wordElementName(child) === localName) ?? null;
};

const getWordAttribute = (
  element: XmlElement | null | undefined,
  localName: string,
): string | null => getAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACES, localName);

const collectText = (element: XmlElement): string => {
  let text = "";

  const walk = (node: XmlElement) => {
    const localName = wordElementName(node);
    if (localName === "t") {
      text += getTextContent(node);
      return;
    }
    if (localName === "br") {
      text += "\n";
      return;
    }
    if (localName === "tab") {
      text += "\t";
      return;
    }
    if (localName === "del" || localName === "delText" || localName === "moveFrom") {
      return;
    }
    for (const child of childElements(node)) {
      walk(child);
    }
  };

  walk(element);
  return text;
};

const countAcceptedTextChars = (element: XmlElement): number => {
  let chars = 0;

  const walk = (node: XmlElement) => {
    const localName = wordElementName(node);
    if (localName === "t") {
      chars += getTextContent(node).length;
      return;
    }
    if (localName === "del" || localName === "delText" || localName === "moveFrom") {
      return;
    }
    for (const child of childElements(node)) {
      walk(child);
    }
  };

  walk(element);
  return chars;
};

const readParagraphProperties = (paragraph: XmlElement): ParagraphProperties => {
  const properties = findWordChild(paragraph, "pPr");
  if (!properties) {
    return {};
  }

  const result: ParagraphProperties = {};
  const style = findWordChild(properties, "pStyle");
  const styleValue = getWordAttribute(style, "val");
  if (styleValue !== null) {
    result.style = styleValue;
  }

  const justification = findWordChild(properties, "jc");
  const alignment = getWordAttribute(justification, "val");
  if (
    alignment === "left" ||
    alignment === "center" ||
    alignment === "right" ||
    alignment === "both"
  ) {
    result.alignment = alignment;
  }
  return result;
};

const readRunMetrics = (paragraph: XmlElement): RunMetrics[] => {
  const metrics: RunMetrics[] = [];

  for (const run of childElements(paragraph)) {
    if (wordElementName(run) !== "r") {
      continue;
    }

    const properties = findWordChild(run, "rPr");
    const boldProperty = findWordChild(properties, "b");
    const boldValue = getWordAttribute(boldProperty, "val");
    const bold = boldProperty !== null && boldValue !== "0" && boldValue !== "false";

    const sizeProperty = findWordChild(properties, "sz");
    const sizeValue = getWordAttribute(sizeProperty, "val");
    const parsedSize = sizeValue === null ? Number.NaN : Number.parseInt(sizeValue, 10);
    const fontSize = Number.isFinite(parsedSize) && parsedSize > 0 ? parsedSize : undefined;

    const chars = countAcceptedTextChars(run);
    if (chars === 0) {
      continue;
    }

    const entry: RunMetrics = { bold, chars };
    if (fontSize !== undefined) {
      entry.fontSize = fontSize;
    }
    metrics.push(entry);
  }

  return metrics;
};

const readParagraph = (paragraph: XmlElement): ExtractedDocxTableCellParagraph => {
  const entry: ExtractedDocxTableCellParagraph = { text: collectText(paragraph) };
  const { style, alignment } = readParagraphProperties(paragraph);
  if (style !== undefined) {
    entry.style = style;
  }
  if (alignment !== undefined) {
    entry.alignment = alignment;
  }

  const runs = readRunMetrics(paragraph);
  if (runs.length === 0) {
    return entry;
  }
  const totalChars = runs.reduce((sum, run) => sum + run.chars, 0);
  const boldChars = runs.reduce((sum, run) => sum + (run.bold ? run.chars : 0), 0);
  if (boldChars > totalChars / 2) {
    entry.bold = true;
  }
  const firstFontSize = runs.find((run) => run.fontSize !== undefined)?.fontSize;
  if (firstFontSize !== undefined) {
    entry.fontSize = firstFontSize;
  }
  return entry;
};

// ---------------------------------------------------------------------------
// Tables
//
// A `w:tbl` renders as GFM pipe rows, one extracted paragraph per row. Emitting
// the cells as a flat row-major paragraph list — every value, then every label —
// destroys the column-to-label association that downstream review and indexing
// depend on.
//
// This is deliberately GFM-only, unlike `markdown/renderTable.ts`, which falls
// back to an inline HTML `<table>` for merged or nested cells. That renderer
// targets a markdown document, where a lossless `colspan`/`rowspan` grid is
// worth the HTML; this one targets a flat paragraph list consumed as text, where
// one line per row is the property that matters. Merges therefore flatten into
// grid columns (see `readTableCell`) rather than switching output modes.
// ---------------------------------------------------------------------------

const TABLE_DELIMITER_CELL = "---";

/** GFM cannot nest tables; an inner table joins its cells inside the outer cell. */
const NESTED_TABLE_CELL_SEPARATOR = " / ";

/** Word supports 63 table columns; cap well above that so a hostile `w:gridSpan` cannot balloon a row. */
const MAX_TABLE_COLUMNS = 256;

/** Bound the mutual recursion between a cell and the tables nested inside it. */
const MAX_NESTED_TABLE_DEPTH = 8;

/** Rows collected from one `w:tbl`. The column cap alone leaves row count unbounded. */
const MAX_TABLE_ROWS = 8192;

/**
 * Characters one extraction emits, shared by the body and every header/footer
 * part. Element count is bounded at unzip, but a bounded element count still
 * renders an unbounded number of table rows once `w:gridSpan` padding and the
 * GFM scaffolding are counted, so the emitted side carries its own ceiling.
 */
const MAX_EXTRACTED_CHARS = 8_000_000;

/**
 * Collect a table's `w:tr`, or a row's `w:tc`, seeing through the wrappers Word
 * puts around them (`w:sdt` / `w:sdtContent` content controls, `w:customXml`).
 * The walk stops at `w:tbl` and `w:p` so a nested table's rows and cells never
 * leak into the grid of the table that contains them.
 */
const collectTableParts = (
  parent: XmlElement,
  localName: "tr" | "tc",
  limit: number,
): XmlElement[] => {
  const parts: XmlElement[] = [];

  const walk = (node: XmlElement) => {
    for (const child of childElements(node)) {
      // Bound the collection itself: a caller that drops the tail afterwards
      // has already paid for the whole array, and every table walk shares this.
      if (parts.length >= limit) {
        return;
      }
      const childName = wordElementName(child);
      if (childName === localName) {
        parts.push(child);
        continue;
      }
      if (childName === "tbl" || childName === "p") {
        continue;
      }
      walk(child);
    }
  };

  walk(parent);
  return parts;
};

/**
 * Raw (unescaped) text of one cell: its paragraphs in order, one per line.
 * Blank paragraphs are dropped so a cell padded with empty paragraphs does not
 * render as a run of `<br>`. A nested table contributes one line per inner row.
 */
const readCellSourceParagraphs = (
  cell: XmlElement,
  depth: number,
): ExtractedDocxTableCellParagraph[] => {
  const paragraphs: ExtractedDocxTableCellParagraph[] = [];

  const walk = (node: XmlElement) => {
    for (const child of childElements(node)) {
      const childName = wordElementName(child);
      if (childName === "p") {
        const paragraph = readParagraph(child);
        if (paragraph.text.length > 0) {
          paragraphs.push(paragraph);
        }
        // `collectText` already descended for text, including any textbox
        // inside the paragraph; descending again would duplicate the cell text.
        continue;
      }
      if (childName === "tbl") {
        if (depth >= MAX_NESTED_TABLE_DEPTH) {
          continue;
        }
        for (const row of collectTableParts(child, "tr", MAX_TABLE_ROWS)) {
          for (const nestedCell of collectTableParts(row, "tc", MAX_TABLE_COLUMNS)) {
            for (const paragraph of readCellSourceParagraphs(nestedCell, depth + 1)) {
              paragraphs.push(paragraph);
            }
          }
        }
        continue;
      }
      walk(child);
    }
  };

  walk(cell);
  return paragraphs;
};

const readCellRenderedLines = (cell: XmlElement, depth: number): string[] => {
  const lines: string[] = [];

  const walk = (node: XmlElement) => {
    for (const child of childElements(node)) {
      const childName = wordElementName(child);
      if (childName === "p") {
        const text = collectText(child);
        if (text.length > 0) {
          lines.push(text);
        }
        continue;
      }
      if (childName === "tbl") {
        if (depth < MAX_NESTED_TABLE_DEPTH) {
          for (const line of flattenNestedTable(child, depth + 1)) {
            lines.push(line);
          }
        }
        continue;
      }
      walk(child);
    }
  };

  walk(cell);
  return lines;
};

const flattenNestedTable = (table: XmlElement, depth: number): string[] => {
  const lines: string[] = [];

  for (const row of collectTableParts(table, "tr", MAX_TABLE_ROWS)) {
    const cells = collectTableParts(row, "tc", MAX_TABLE_COLUMNS).map((cell) =>
      readCellRenderedLines(cell, depth).join("\n"),
    );
    if (cells.some((text) => text.length > 0)) {
      lines.push(cells.join(NESTED_TABLE_CELL_SEPARATOR));
    }
  }

  return lines;
};

type ExtractedTableCell = {
  /** Raw cell text; escaped only when it reaches a row line. */
  text: string;
  /** Source paragraphs before GFM joins them with `<br>`. */
  paragraphs: ExtractedDocxTableCellParagraph[];
  /** Grid columns the cell occupies (`w:gridSpan`), at least 1. */
  gridSpan: number;
};

const emptyTableCell = (): ExtractedTableCell => ({
  text: "",
  paragraphs: [],
  gridSpan: 1,
});

const readTableCell = (cell: XmlElement, depth: number): ExtractedTableCell => {
  const properties = findWordChild(cell, "tcPr");

  const gridSpanValue = getWordAttribute(findWordChild(properties, "gridSpan"), "val");
  const parsedGridSpan = gridSpanValue === null ? 1 : Number.parseInt(gridSpanValue, 10);
  const gridSpan = Number.isFinite(parsedGridSpan) && parsedGridSpan > 1 ? parsedGridSpan : 1;

  // A `w:vMerge` without `w:val="restart"` continues the cell above it. Word
  // renders no content of its own there, so emit an empty grid column: it holds
  // the row's column alignment without repeating the anchor cell's value or
  // surfacing content Word itself never shows.
  const vMerge = findWordChild(properties, "vMerge");
  if (vMerge !== null && getWordAttribute(vMerge, "val") !== "restart") {
    return { text: "", paragraphs: [], gridSpan };
  }

  const paragraphs = readCellSourceParagraphs(cell, depth);
  return {
    text: readCellRenderedLines(cell, depth).join("\n"),
    paragraphs,
    gridSpan,
  };
};

/**
 * Does the row declare itself a header? `w:tblHeader` is the only OOXML signal
 * that says so: it marks the row Word repeats at the top of each page. The
 * neighbouring `w:tblLook/@w:firstRow` is conditional *formatting* that Word
 * writes on essentially every table (its default `w:val="04A0"`), so keying off
 * it would promote the first data row of almost every document.
 *
 * Without the flag the table is headerless and GFM gets a synthetic empty header
 * row: a table's first row is data until the document says otherwise, and column
 * names invented here would be read back as facts about the document.
 */
const declaresHeaderRow = (row: XmlElement): boolean => {
  const header = findWordChild(findWordChild(row, "trPr"), "tblHeader");
  if (header === null) {
    return false;
  }
  const value = getWordAttribute(header, "val");
  return value !== "0" && value !== "false";
};

const readRowGridOffset = (row: XmlElement, localName: "gridBefore" | "gridAfter"): number => {
  const properties = findWordChild(row, "trPr");
  const value = getWordAttribute(findWordChild(properties, localName), "val");
  const parsed = value === null ? 0 : Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? Math.min(parsed, MAX_TABLE_COLUMNS) : 0;
};

type TableGrid = {
  /** Raw cell text per row, already expanded across the grid columns each cell spans. */
  rows: ExtractedTableCell[][];
  columnCount: number;
  firstRowIsHeader: boolean;
};

const readTableGrid = (table: XmlElement): TableGrid => {
  const rows: ExtractedTableCell[][] = [];
  let columnCount = 0;
  let firstRowIsHeader = false;

  for (const [rowIndex, row] of collectTableParts(table, "tr", MAX_TABLE_ROWS).entries()) {
    if (rowIndex === 0) {
      firstRowIsHeader = declaresHeaderRow(row);
    }
    const columns: ExtractedTableCell[] = [];
    const gridBefore = readRowGridOffset(row, "gridBefore");
    for (let index = 0; index < gridBefore; index += 1) {
      columns.push(emptyTableCell());
    }
    for (const cell of collectTableParts(row, "tc", MAX_TABLE_COLUMNS)) {
      if (columns.length >= MAX_TABLE_COLUMNS) {
        break;
      }
      const extractedCell = readTableCell(cell, 0);
      columns.push(extractedCell);
      // A horizontally merged cell holds its text in the first column it spans;
      // the rest stay empty so every row keeps the same column boundaries.
      const padding = Math.min(extractedCell.gridSpan - 1, MAX_TABLE_COLUMNS - columns.length);
      for (let index = 0; index < padding; index++) {
        columns.push(emptyTableCell());
      }
    }
    const gridAfter = Math.min(
      readRowGridOffset(row, "gridAfter"),
      MAX_TABLE_COLUMNS - columns.length,
    );
    for (let index = 0; index < gridAfter; index += 1) {
      columns.push(emptyTableCell());
    }
    if (columns.length > columnCount) {
      columnCount = columns.length;
    }
    rows.push(columns);
  }

  return { rows, columnCount, firstRowIsHeader };
};

/** Pad a row to the table's column count and escape each cell into a pipe row. */
const toRowLine = (columns: readonly ExtractedTableCell[], columnCount: number): string => {
  const cells: string[] = [];
  for (let column = 0; column < columnCount; column++) {
    cells.push(escapeTableCell(columns[column]?.text ?? ""));
  }
  return `| ${cells.join(" | ")} |`;
};

type RenderedTableRow = {
  text: string;
  position: DocxTableRowPosition;
};

/** Render a `w:tbl` as GFM rows. A table with no cell at all renders nothing. */
const renderTableRows = (table: XmlElement, tableIndex: number): RenderedTableRow[] => {
  const { rows, columnCount, firstRowIsHeader } = readTableGrid(table);
  const [firstRow, ...remainingRows] = rows;
  if (columnCount === 0 || firstRow === undefined) {
    return [];
  }

  const rendered: RenderedTableRow[] = [];
  const pushScaffolding = (text: string, kind: "syntheticHeader" | "delimiter") => {
    rendered.push({ text, position: { table: tableIndex, kind } });
  };
  const pushCells = (cells: readonly ExtractedTableCell[]) => {
    rendered.push({
      text: toRowLine(cells, columnCount),
      position: {
        table: tableIndex,
        kind: "cells",
        cells: Array.from({ length: columnCount }, (_, column) => ({
          paragraphs: cells.at(column)?.paragraphs ?? [],
        })),
      },
    });
  };

  if (firstRowIsHeader) {
    pushCells(firstRow);
  } else {
    pushScaffolding(toRowLine([], columnCount), "syntheticHeader");
  }
  pushScaffolding(
    toRowLine(
      Array.from({ length: columnCount }, () => ({
        text: TABLE_DELIMITER_CELL,
        paragraphs: [],
        gridSpan: 1,
      })),
      columnCount,
    ),
    "delimiter",
  );
  for (const row of firstRowIsHeader ? remainingRows : rows) {
    pushCells(row);
  }

  return rendered;
};

// ---------------------------------------------------------------------------
// Container walk
// ---------------------------------------------------------------------------

/** Mutable remainder of {@link MAX_EXTRACTED_CHARS}, shared across parts. */
type CharBudget = { remaining: number };

const createCharBudget = (): CharBudget => ({ remaining: MAX_EXTRACTED_CHARS });

type ExtractContainerOptions = {
  container: XmlElement;
  source: DocxParagraphSource;
  startIndex: number;
  startTableIndex: number;
  budget: CharBudget;
};

type ExtractContainerResult = {
  paragraphs: ExtractedDocxParagraph[];
  charCount: number;
  /** `w:tbl` elements rendered, so the next part continues the table numbering. */
  tableCount: number;
};

const extractContainer = ({
  container,
  source,
  startIndex,
  startTableIndex,
  budget,
}: ExtractContainerOptions): ExtractContainerResult => {
  const paragraphs: ExtractedDocxParagraph[] = [];
  let charCount = 0;
  let tableCount = 0;

  const pushProse = (paragraph: XmlElement) => {
    const extracted = readParagraph(paragraph);
    const { text } = extracted;
    const entry: ExtractedDocxParagraph = {
      index: startIndex + paragraphs.length,
      source,
      ...extracted,
    };

    paragraphs.push(entry);
    charCount += text.length;
    budget.remaining -= text.length;
  };

  const pushTableRow = ({ text, position }: RenderedTableRow) => {
    paragraphs.push({
      index: startIndex + paragraphs.length,
      text,
      source,
      tableRow: position,
    });
    charCount += text.length;
    budget.remaining -= text.length;
  };

  /**
   * Walk block content in document order. Descent mirrors the previous
   * `findAllDeep(container, "w", "p")` — every wrapper (`w:sdt`, textboxes) is
   * still entered — except that a `w:tbl` is consumed as a table instead of
   * having its cell paragraphs emitted individually.
   */
  const walkBlocks = (node: XmlElement) => {
    for (const child of childElements(node)) {
      if (budget.remaining <= 0) {
        return;
      }
      const childName = wordElementName(child);
      if (childName === "tbl") {
        for (const row of renderTableRows(child, startTableIndex + tableCount)) {
          pushTableRow(row);
        }
        tableCount += 1;
        continue;
      }
      if (childName === "p") {
        pushProse(child);
      }
      walkBlocks(child);
    }
  };

  walkBlocks(container);
  return { paragraphs, charCount, tableCount };
};

type ExtractPartsOptions = {
  archive: DocxArchive;
  source: "header" | "footer";
  rootName: "hdr" | "ftr";
  startIndex: number;
  startTableIndex: number;
  /** Part paths to read, in extraction order — see {@link resolveReferencedHeaderFooterParts}. */
  paths: readonly string[];
  budget: CharBudget;
};

const extractParts = async ({
  archive,
  source,
  rootName,
  startIndex,
  startTableIndex,
  paths,
  budget,
}: ExtractPartsOptions): Promise<ExtractContainerResult> => {
  const paragraphs: ExtractedDocxParagraph[] = [];
  let charCount = 0;
  let tableCount = 0;
  let nextIndex = startIndex;

  for (const path of paths) {
    // oxlint-disable-next-line no-await-in-loop -- part order defines stable paragraph indices
    const xml = await archive.readEntryString(path);
    if (xml === null) {
      continue;
    }
    const root = parseXml(xml);
    const container = findDeep(root, "w", rootName);
    if (!container) {
      continue;
    }
    const result = extractContainer({
      container,
      source,
      startIndex: nextIndex,
      startTableIndex: startTableIndex + tableCount,
      budget,
    });
    for (const paragraph of result.paragraphs) {
      paragraphs.push(paragraph);
    }
    charCount += result.charCount;
    tableCount += result.tableCount;
    nextIndex += result.paragraphs.length;
  }

  return { paragraphs, charCount, tableCount };
};

/** A `word/_rels/document.xml.rels` `Target` is relative to `word/`; resolve it to a full archive-entry path. */
const resolveWordPartPath = (target: string): string =>
  target.startsWith("/") ? target.slice(1) : `word/${target}`;

type ReferencedHeaderFooterParts = {
  headers: string[];
  footers: string[];
};

/**
 * Resolve the header/footer parts actually wired into the document via
 * `word/_rels/document.xml.rels` + each section's `w:headerReference` /
 * `w:footerReference`, instead of extracting every `word/header*.xml` /
 * `word/footer*.xml` entry by filename. A DOCX can carry an orphaned
 * header/footer part (stale, or planted by an attacker) that no section
 * references — reading it unconditionally would surface prompt-injection or
 * stale content that Word itself never renders.
 */
const resolveReferencedHeaderFooterParts = async (
  archive: DocxArchive,
  documentRoot: XmlElement,
): Promise<ReferencedHeaderFooterParts> => {
  const relsXml = await archive.readEntryString(DOCUMENT_RELS_PATH);
  if (relsXml === null) {
    return { headers: [], footers: [] };
  }
  const relationships = parseRelationships(relsXml);

  const headerRIds = new Set<string>();
  const footerRIds = new Set<string>();
  for (const sectPr of findAllDeep(documentRoot, "w", "sectPr")) {
    for (const ref of findAllDeep(sectPr, "w", "headerReference")) {
      const rId = getAttribute(ref, "r", "id");
      if (rId !== null) {
        headerRIds.add(rId);
      }
    }
    for (const ref of findAllDeep(sectPr, "w", "footerReference")) {
      const rId = getAttribute(ref, "r", "id");
      if (rId !== null) {
        footerRIds.add(rId);
      }
    }
  }

  const resolvePaths = (rIds: Set<string>, relationshipType: string): string[] => {
    const paths = new Set<string>();
    for (const rId of rIds) {
      const relationship = relationships.get(rId);
      if (
        !relationship ||
        relationship.type !== relationshipType ||
        relationship.targetMode === "External"
      ) {
        continue;
      }
      paths.add(resolveWordPartPath(relationship.target));
    }
    return [...paths].toSorted();
  };

  return {
    headers: resolvePaths(headerRIds, RELATIONSHIP_TYPES.header),
    footers: resolvePaths(footerRIds, RELATIONSHIP_TYPES.footer),
  };
};

const createEmptyResult = (): ExtractedDocxText => ({
  paragraphs: [],
  charCount: 0,
  view: "accepted",
});

/** Extract paragraph text and formatting metadata from a DOCX archive. */
export const extractDocxText = async (
  bytes: ArrayBuffer | Uint8Array,
): Promise<ExtractedDocxText> => {
  const archive = await loadDocxArchive(bytes);
  const documentXml = await archive.readEntryString("word/document.xml");
  if (documentXml === null) {
    return createEmptyResult();
  }

  const root = parseXml(documentXml);
  const body = findDeep(root, "w", "body");
  if (!body) {
    return createEmptyResult();
  }

  const referencedParts = await resolveReferencedHeaderFooterParts(archive, root);
  const budget = createCharBudget();

  const headers = await extractParts({
    archive,
    source: "header",
    rootName: "hdr",
    startIndex: 0,
    startTableIndex: 0,
    paths: referencedParts.headers,
    budget,
  });
  const bodyResult = extractContainer({
    container: body,
    source: "body",
    startIndex: headers.paragraphs.length,
    startTableIndex: headers.tableCount,
    budget,
  });
  const footers = await extractParts({
    archive,
    source: "footer",
    rootName: "ftr",
    startIndex: headers.paragraphs.length + bodyResult.paragraphs.length,
    startTableIndex: headers.tableCount + bodyResult.tableCount,
    paths: referencedParts.footers,
    budget,
  });

  return {
    paragraphs: [...headers.paragraphs, ...bodyResult.paragraphs, ...footers.paragraphs],
    charCount: headers.charCount + bodyResult.charCount + footers.charCount,
    view: "accepted",
  };
};
