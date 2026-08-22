/**
 * Headless document builders.
 *
 * Thin constructors for the `@stll/docx-core` model so a server can assemble
 * a report in code (`createEmptyDocument` → push blocks → `createDocx`)
 * without hand-writing model literals. They build model values only; the
 * serializer owns the OOXML.
 */

import { TaggedError } from "better-result";

import type { ShadingProperties } from "../../types/colors";
import type {
  BookmarkEnd,
  BookmarkStart,
  ComplexField,
  Endnote,
  Hyperlink,
  Paragraph,
  ParagraphContent,
  Run,
  Table,
  TableCell,
  TableRow,
} from "../../types/content";
import type { Document } from "../../types/document";
import type { ParagraphFormatting, TextFormatting } from "../../types/formatting";

/** Character style applied to hyperlink text; present in the bundled style sets. */
const HYPERLINK_STYLE_ID = "Hyperlink";
/** Character style applied to the endnote reference mark. */
const ENDNOTE_REFERENCE_STYLE_ID = "EndnoteReference";
/** Paragraph style applied to endnote body paragraphs. */
const ENDNOTE_TEXT_STYLE_ID = "EndnoteText";
/** Table style used by `table()`; present in the bundled style sets. */
const TABLE_STYLE_ID = "TableGrid";
/** `w:tblW w:type="pct"` is in fiftieths of a percent: 5000 = 100%. */
const FULL_WIDTH_PCT = 5000;

export const HEADING_LEVELS = [1, 2, 3, 4, 5, 6] as const;
export type HeadingLevel = (typeof HEADING_LEVELS)[number];

/** Outline levels a `TOC \o` switch may name (ECMA-376 `w:outlineLvl` 0-8). */
const TOC_LEVEL_MIN = 1;
const TOC_LEVEL_MAX = 9;

/**
 * A builder received a value outside the model's domain (a zero `gridSpan`,
 * a negative column width, a reversed TOC range). Thrown at the boundary so
 * the invalid value never reaches the serializer.
 */
export class InvalidFolioReportBuilderOptionsError extends TaggedError(
  "InvalidFolioReportBuilderOptionsError",
)<{
  message: string;
  path: string;
}> {}

const assertPositiveInteger = (value: number, path: string): void => {
  if (!Number.isInteger(value) || value < 1) {
    throw new InvalidFolioReportBuilderOptionsError({
      message: `${path} must be a positive integer, got ${String(value)}`,
      path,
    });
  }
};

const assertHeadingLevel = (level: number): void => {
  if (!HEADING_LEVELS.some((known) => known === level)) {
    throw new InvalidFolioReportBuilderOptionsError({
      message: `level must be one of ${HEADING_LEVELS.join(", ")}, got ${String(level)}`,
      path: "level",
    });
  }
};

const assertTocLevels = ({ from, to }: { from: number; to: number }): void => {
  const inRange = (value: number) =>
    Number.isInteger(value) && value >= TOC_LEVEL_MIN && value <= TOC_LEVEL_MAX;
  if (!inRange(from) || !inRange(to) || from > to) {
    throw new InvalidFolioReportBuilderOptionsError({
      message: `levels must satisfy ${TOC_LEVEL_MIN} <= from <= to <= ${TOC_LEVEL_MAX}, got ${String(from)}-${String(to)}`,
      path: "levels",
    });
  }
};

export const run = (text: string, formatting?: TextFormatting): Run => ({
  type: "run",
  ...(formatting ? { formatting } : {}),
  content: [{ type: "text", text }],
});

export const paragraph = (
  content: string | ParagraphContent[],
  formatting?: ParagraphFormatting,
): Paragraph => ({
  type: "paragraph",
  ...(formatting ? { formatting } : {}),
  content: typeof content === "string" ? [run(content)] : content,
});

type HeadingOptions = {
  text: string;
  level: HeadingLevel;
};

/** A paragraph in the `Heading<level>` style. */
export const heading = ({ text, level }: HeadingOptions): Paragraph => {
  assertHeadingLevel(level);
  return paragraph(text, { styleId: `Heading${level}` });
};

/** An empty paragraph carrying a hard page break. */
export const pageBreak = (): Paragraph => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "break", breakType: "page" }] }],
});

type HyperlinkTarget = { href: string; anchor?: never } | { anchor: string; href?: never };

type HyperlinkOptions = HyperlinkTarget & {
  text: string;
  formatting?: TextFormatting;
  tooltip?: string;
};

/**
 * An external (`href`) or in-document (`anchor`, a bookmark name) link. The
 * relationship for an external link is minted when the document is written.
 */
export const hyperlink = ({
  text,
  formatting,
  tooltip,
  href,
  anchor,
}: HyperlinkOptions): Hyperlink => ({
  type: "hyperlink",
  ...(href !== undefined ? { href } : { anchor }),
  ...(tooltip !== undefined ? { tooltip } : {}),
  children: [run(text, { styleId: HYPERLINK_STYLE_ID, ...formatting })],
});

type BookmarkOptions = {
  name: string;
  content: ParagraphContent[];
  /**
   * Bookmark id, unique per document. Defaults to a process-wide counter,
   * which is unique for documents built from scratch; pass an explicit id
   * when adding bookmarks to a parsed document.
   */
  id?: number;
};

let nextBookmarkId = 0;

/** `content` wrapped in a named bookmark, the target of `hyperlink({ anchor })`. */
export const bookmark = ({ name, content, id }: BookmarkOptions): ParagraphContent[] => {
  const bookmarkId = id ?? nextBookmarkId++;
  const start: BookmarkStart = { type: "bookmarkStart", id: bookmarkId, name };
  const end: BookmarkEnd = { type: "bookmarkEnd", id: bookmarkId };
  return [start, ...content, end];
};

export type TableCellSpec =
  | string
  | {
      content: Paragraph[];
      shading?: ShadingProperties;
      gridSpan?: number;
      vMerge?: "restart" | "continue";
    };

type TableOptions = {
  /** Header row labels; rendered bold, optionally shaded and repeated per page. */
  header?: string[];
  rows: TableCellSpec[][];
  /** Grid column widths in twips. Omitted: the consumer autofits. */
  columnWidths?: number[];
  headerShading?: ShadingProperties;
  /** Repeat the header row at the top of every page (default true). */
  repeatHeader?: boolean;
};

type BuildCellOptions = {
  spec: TableCellSpec;
  columnWidths: number[] | undefined;
  /** Grid column this cell starts at, for width lookup. */
  gridIndex: number;
  shading: ShadingProperties | undefined;
  textFormatting: TextFormatting | undefined;
};

const cellWidth = (
  columnWidths: number[] | undefined,
  gridIndex: number,
  gridSpan: number,
): number | undefined => {
  if (!columnWidths) {
    return undefined;
  }
  let width = 0;
  for (let column = gridIndex; column < gridIndex + gridSpan; column++) {
    const columnWidth = columnWidths.at(column);
    if (columnWidth === undefined) {
      return undefined;
    }
    width += columnWidth;
  }
  return width;
};

const buildCell = ({
  spec,
  columnWidths,
  gridIndex,
  shading,
  textFormatting,
}: BuildCellOptions): TableCell => {
  const resolved =
    typeof spec === "string"
      ? { content: [paragraph([run(spec, textFormatting)])], shading }
      : { ...spec, shading: spec.shading ?? shading };
  const gridSpan = resolved.gridSpan ?? 1;
  assertPositiveInteger(gridSpan, "gridSpan");
  const width = cellWidth(columnWidths, gridIndex, gridSpan);
  const formatting = {
    ...(width !== undefined ? { width: { type: "dxa" as const, value: width } } : {}),
    ...(resolved.shading ? { shading: resolved.shading } : {}),
    ...(resolved.gridSpan !== undefined ? { gridSpan: resolved.gridSpan } : {}),
    ...(resolved.vMerge !== undefined ? { vMerge: resolved.vMerge } : {}),
  };
  return {
    type: "tableCell",
    ...(Object.keys(formatting).length > 0 ? { formatting } : {}),
    // A cell must end with a paragraph; a merged-away cell is typically empty.
    content: resolved.content.length > 0 ? resolved.content : [paragraph([])],
  };
};

type BuildRowOptions = {
  cells: TableCellSpec[];
  columnWidths: number[] | undefined;
  shading: ShadingProperties | undefined;
  textFormatting: TextFormatting | undefined;
  header: boolean;
};

const buildRow = ({
  cells,
  columnWidths,
  shading,
  textFormatting,
  header,
}: BuildRowOptions): TableRow => {
  const built: TableCell[] = [];
  let gridIndex = 0;
  for (const spec of cells) {
    const cell = buildCell({ spec, columnWidths, gridIndex, shading, textFormatting });
    built.push(cell);
    gridIndex += cell.formatting?.gridSpan ?? 1;
  }
  return {
    type: "tableRow",
    ...(header ? { formatting: { header: true, cantSplit: true } } : {}),
    cells: built,
  };
};

/**
 * A full-width grid table in the `TableGrid` style. A string cell becomes one
 * plain paragraph; an object cell supplies its own paragraphs plus optional
 * shading and horizontal (`gridSpan`) or vertical (`vMerge`) merge.
 */
export const table = ({
  header,
  rows,
  columnWidths,
  headerShading,
  repeatHeader = true,
}: TableOptions): Table => {
  columnWidths?.forEach((width, index) => assertPositiveInteger(width, `columnWidths[${index}]`));
  const builtRows: TableRow[] = [];
  if (header) {
    builtRows.push(
      buildRow({
        cells: header,
        columnWidths,
        shading: headerShading,
        textFormatting: { bold: true },
        header: repeatHeader,
      }),
    );
  }
  for (const cells of rows) {
    builtRows.push(
      buildRow({
        cells,
        columnWidths,
        shading: undefined,
        textFormatting: undefined,
        header: false,
      }),
    );
  }
  return {
    type: "table",
    formatting: {
      styleId: TABLE_STYLE_ID,
      width: { type: "pct", value: FULL_WIDTH_PCT },
      layout: columnWidths ? "fixed" : "autofit",
    },
    ...(columnWidths ? { columnWidths } : {}),
    rows: builtRows,
  };
};

/**
 * Register an endnote on `doc` and return the reference run to place in body
 * text. Allocates the next free endnote id (Word reserves 0 and -1 for the
 * separator notes) and pushes the note into `doc.package.endnotes`.
 */
export const endnote = (doc: Document, content: string | Paragraph[]): Run => {
  const endnotes = doc.package.endnotes ?? [];
  doc.package.endnotes = endnotes;
  const id = Math.max(0, ...endnotes.map((note) => note.id)) + 1;
  const body =
    typeof content === "string"
      ? [paragraph(content, { styleId: ENDNOTE_TEXT_STYLE_ID })]
      : content;
  const note: Endnote = { type: "endnote", id, content: body };
  endnotes.push(note);
  return {
    type: "run",
    formatting: { styleId: ENDNOTE_REFERENCE_STYLE_ID },
    content: [{ type: "endnoteRef", id }],
  };
};

type TableOfContentsOptions = {
  /** Heading levels to include, `1 <= from <= to <= 9` (default 1-3). */
  levels?: { from: number; to: number };
  /** Make entries hyperlinks (`\h`, default true). */
  hyperlinks?: boolean;
  /** Result text shown until the consumer recomputes the field. */
  placeholderText?: string;
};

const DEFAULT_TOC_LEVELS = { from: 1, to: 3 };
const DEFAULT_TOC_PLACEHOLDER = "Update the field to build the table of contents.";

/**
 * A paragraph holding a dirty `TOC` field, so the consumer recomputes the
 * table on open. Set `package.settings.updateFields` as well to have Word
 * recompute without prompting for each field.
 */
export const createTableOfContentsField = ({
  levels = DEFAULT_TOC_LEVELS,
  hyperlinks = true,
  placeholderText = DEFAULT_TOC_PLACEHOLDER,
}: TableOfContentsOptions = {}): Paragraph => {
  assertTocLevels(levels);
  const switches = [`\\o "${levels.from}-${levels.to}"`];
  if (hyperlinks) {
    switches.push("\\h");
  }
  switches.push("\\z", "\\u");
  const field: ComplexField = {
    type: "complexField",
    instruction: `TOC ${switches.join(" ")}`,
    fieldType: "TOC",
    fieldCode: [],
    fieldResult: [run(placeholderText)],
    dirty: true,
  };
  return { type: "paragraph", content: [field] };
};
