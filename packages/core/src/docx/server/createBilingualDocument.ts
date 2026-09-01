/**
 * Bilingual document transform: body -> two-column table, one row per block.
 *
 * The left column keeps every source block untouched; the right column holds a
 * copy of the same block whose numbering and numbered paragraph styles are
 * cloned per language, so both columns count independently (1. / 1. instead of
 * 1. / 2.) and stay live in Word. Right-column paragraphs receive fresh
 * `paraId`s so callers can address each row later (for example to replace the
 * placeholder copy with a translation by block id). Horizontal paragraph
 * geometry is projected into the half-width cells: full-page indents and tab
 * stops otherwise place signature fields outside their column and let prose
 * overlap the translation.
 *
 * Section breaks cannot live inside a table cell, so the body is split at
 * paragraphs carrying `sectionProperties`: each section becomes its own table
 * and the break paragraph stays between the tables. Source tables (parties,
 * signature blocks) can either span both columns for inline translation or be
 * followed by a full-width target copy, depending on `tableLayout`.
 * Structural-only paragraphs are also kept once between tables; they have no
 * independently editable text for a translation row to address.
 */

import { TaggedError } from "better-result";

import { getParagraphText } from "../paragraphParser";
import type {
  AbstractNumbering,
  BlockContent,
  Document,
  ListRendering,
  NumberingDefinitions,
  NumberingInstance,
  Paragraph,
  ParagraphFormatting,
  SectionProperties,
  Style,
  StyleDefinitions,
  Table,
  TableBorders,
  TableCell,
} from "../../types/document";
import { deterministicHexId } from "../../utils/hexId";

export type BilingualRowKind = "paragraph" | "heading" | "listItem";

/** One translatable unit: a source paragraph and its right-column copy. */
export type BilingualParagraphRef = {
  /** `paraId` of the untouched source paragraph (left column). */
  sourceParaId: string | undefined;
  /** `paraId` minted for the right-column copy; stable across re-runs. */
  targetParaId: string;
  /** Plain text of the source paragraph. */
  sourceText: string;
};

export type BilingualRow =
  | ({
      kind: BilingualRowKind;
      /** Equals `targetParaId`; the handle callers use to address the row. */
      rowId: string;
    } & BilingualParagraphRef)
  | {
      kind: "table";
      layout: "inline";
      rowId: string;
      /**
       * Every paragraph inside the table, in document order. The table is not
       * copied, so these are the paragraphs to translate in place.
       */
      paragraphs: BilingualTableParagraphRef[];
    }
  | {
      kind: "table";
      layout: "stacked";
      /** Equals the first target paragraph id. */
      rowId: string;
      /** Source/target paragraph pairs in table document order. */
      paragraphs: BilingualParagraphRef[];
    };

export type BilingualTableParagraphRef = {
  paraId: string | undefined;
  sourceText: string;
};

export type BilingualBorders = "none" | "grid";

export const BILINGUAL_TABLE_LAYOUTS = ["inline", "stacked"] as const;
export type BilingualTableLayout = (typeof BILINGUAL_TABLE_LAYOUTS)[number];

export type CreateBilingualDocumentOptions = {
  /**
   * Suffix for cloned style ids and names (for example `"en"` turns
   * `Heading1` into `Heading1-en`). Must be a non-empty token of letters,
   * digits, or `-`.
   */
  targetStyleSuffix: string;
  /** Table borders; legal practice is usually `"none"`. Default `"none"`. */
  borders?: BilingualBorders;
  /**
   * Layout for source tables. `"inline"` keeps one full-width table whose
   * paragraphs are translated in place. `"stacked"` keeps the source table
   * and adds an independently addressable target copy below it. Default
   * `"inline"`.
   */
  tableLayout?: BilingualTableLayout;
  /**
   * Paragraph handles exposed by Folio's canonical AI-edit snapshot for the
   * source DOCX. Only these paragraphs may become translation rows.
   */
  editableParagraphIds: ReadonlySet<string>;
};

export type CreateBilingualDocumentResult = {
  document: Document;
  rows: BilingualRow[];
  /** Non-fatal fidelity notes (for example an unresolvable numbering style link). */
  warnings: string[];
};

const STYLE_SUFFIX_PATTERN = /^[A-Za-z0-9-]+$/u;
const DEFAULT_TABLE_LAYOUT = BILINGUAL_TABLE_LAYOUTS[0];

export class InvalidBilingualDocumentOptionsError extends TaggedError(
  "InvalidBilingualDocumentOptionsError",
)<{
  message: string;
  option: "tableLayout" | "targetStyleSuffix";
}> {}
class UneditableBilingualManifestError extends TaggedError("UneditableBilingualManifestError")<{
  message: string;
  missingHandleCount: number;
}> {}
const FULL_WIDTH_PCT = 5000;
const HALF_WIDTH_PCT = 2500;
const A4_TEXT_WIDTH_TWIPS = 9072;
const ROW_ID_NAMESPACE = "folio-bilingual";
const BILINGUAL_TABLE_STYLE_ID = "FolioBilingualTranslation";
const MIN_COLUMN_TEXT_WIDTH_TWIPS = 720;
const MIN_TAB_TRAILING_WIDTH_TWIPS = 360;

const GRID_BORDER = { style: "single", size: 4, space: 0 } as const;

const TABLE_BORDERS: Record<BilingualBorders, TableBorders> = {
  none: {
    top: { style: "nil" },
    bottom: { style: "nil" },
    left: { style: "nil" },
    right: { style: "nil" },
    insideH: { style: "nil" },
    insideV: { style: "nil" },
  },
  grid: {
    top: GRID_BORDER,
    bottom: GRID_BORDER,
    left: GRID_BORDER,
    right: GRID_BORDER,
    insideH: GRID_BORDER,
    insideV: GRID_BORDER,
  },
};

const isBilingualTableLayout = (value: unknown): value is BilingualTableLayout =>
  typeof value === "string" && BILINGUAL_TABLE_LAYOUTS.some((layout) => layout === value);

export function createBilingualDocument(
  source: Document,
  options: CreateBilingualDocumentOptions,
): CreateBilingualDocumentResult {
  if (!STYLE_SUFFIX_PATTERN.test(options.targetStyleSuffix)) {
    throw new InvalidBilingualDocumentOptionsError({
      message: `targetStyleSuffix must match ${STYLE_SUFFIX_PATTERN}; received ${JSON.stringify(options.targetStyleSuffix)}`,
      option: "targetStyleSuffix",
    });
  }
  if (options.tableLayout !== undefined && !isBilingualTableLayout(options.tableLayout)) {
    throw new InvalidBilingualDocumentOptionsError({
      message: `tableLayout must be one of ${BILINGUAL_TABLE_LAYOUTS.join(", ")}; received ${JSON.stringify(options.tableLayout)}`,
      option: "tableLayout",
    });
  }
  const borders = options.borders ?? "none";
  const tableLayout = options.tableLayout ?? DEFAULT_TABLE_LAYOUT;
  const warnings: string[] = [];

  const styles = source.package.styles;
  const numbering = source.package.numbering;
  const styleById = new Map((styles?.styles ?? []).map((style) => [style.styleId, style]));

  const blocks = flattenBlocks(source.package.document.content);
  const cloner = createNumberingCloner({ numbering, styleById, warnings });
  const styleCloner = createStyleCloner({
    styleById,
    suffix: options.targetStyleSuffix,
    cloner,
  });
  const paraIds = createParaIdMinter(collectPackageParaIds(source.package));
  const bookmarkIds = createBookmarkIdMinter(source.package);

  const rows: BilingualRow[] = [];
  const content: BlockContent[] = [];
  let sectionRows: Table["rows"] = [];
  const textWidth = resolveTextWidthTwips(source);

  const flushSection = (): void => {
    if (sectionRows.length > 0) {
      content.push(buildTable(sectionRows, borders, textWidth));
    }
    sectionRows = [];
  };

  const copyParagraph = (paragraph: Paragraph): { copy: Paragraph; ref: BilingualParagraphRef } => {
    const targetParaId = paraIds.mint(paragraph.paraId);
    const copy = cloneParagraphForTarget(paragraph, targetParaId, styleCloner, cloner, bookmarkIds);
    return {
      copy,
      ref: {
        sourceParaId: paragraph.paraId,
        targetParaId,
        sourceText: getParagraphText(paragraph),
      },
    };
  };

  for (const block of blocks) {
    if (block.type === "paragraph" && block.sectionProperties) {
      flushSection();
      content.push(block);
      continue;
    }
    if (block.type === "paragraph") {
      if (isEmptyParagraph(block)) {
        continue;
      }
      if (block.paraId === undefined || !options.editableParagraphIds.has(block.paraId)) {
        flushSection();
        content.push(block);
        continue;
      }
      const { copy, ref } = copyParagraph(block);
      rows.push({ kind: classifyParagraph(block, styleById), rowId: ref.targetParaId, ...ref });
      sectionRows.push(buildRow(block, copy, styleById, textWidth));
      continue;
    }
    const paragraphs = collectTableParagraphs(block)
      .filter(
        (paragraph): paragraph is Paragraph & { paraId: string } =>
          paragraph.paraId !== undefined && options.editableParagraphIds.has(paragraph.paraId),
      )
      .map((paragraph) => ({
        paraId: paragraph.paraId,
        sourceText: getParagraphText(paragraph),
      }));
    if (paragraphs.length === 0) {
      flushSection();
      content.push(block);
      continue;
    }
    if (tableLayout === "inline") {
      rows.push({
        kind: "table",
        layout: "inline",
        rowId: paragraphs.at(0)?.paraId ?? tableRowHandle(rows.length),
        paragraphs,
      });
      sectionRows.push(buildInlineTableRow(block));
      continue;
    }

    if (tableLayout === "stacked") {
      const target = cloneTableForTarget({
        table: block,
        editableParagraphIds: options.editableParagraphIds,
        paraIds,
        styleCloner,
        cloner,
        bookmarkIds,
      });
      rows.push({
        kind: "table",
        layout: "stacked",
        rowId: target.paragraphs.at(0)?.targetParaId ?? tableRowHandle(rows.length),
        paragraphs: target.paragraphs,
      });
      sectionRows.push(buildStackedTableRow({ source: block, target: target.table }));
      continue;
    }
    tableLayout satisfies never;
  }
  flushSection();

  const document: Document = {
    ...source,
    package: {
      ...source.package,
      document: { ...source.package.document, content },
      ...(cloner.hasClones() && { numbering: cloner.toDefinitions() }),
      ...(styleCloner.hasClones() && styles && { styles: styleCloner.toDefinitions(styles) }),
    },
  };

  return { document, rows, warnings };
}

// ----------------------------------------------------------------------------
// Blocks
// ----------------------------------------------------------------------------

type BodyBlock = Paragraph | Table;

/** Handle for a table row whose paragraphs carry no `paraId`: its position in
 *  the manifest, which creation and reading derive identically. */
const tableRowHandle = (index: number): string => `table-${index}`;

/** Top-level body blocks with content controls flattened to their children. */
const flattenBlocks = (content: BlockContent[]): BodyBlock[] => {
  const out: BodyBlock[] = [];
  const visit = (block: BlockContent): void => {
    if (block.type === "paragraph" || block.type === "table") {
      out.push(block);
      return;
    }
    for (const child of block.content) {
      visit(child);
    }
  };
  for (const block of content) {
    visit(block);
  }
  return out;
};

const isEmptyParagraph = (paragraph: Paragraph): boolean => {
  if (getParagraphText(paragraph).trim().length > 0) {
    return false;
  }
  // Anything that is not a visible plain-text run (drawings, fields, breaks,
  // hidden text, content controls) must be preserved rather than discarded.
  return paragraph.content.every(
    (item) =>
      item.type === "run" &&
      item.formatting?.hidden !== true &&
      item.content.every((part) => part.type === "text"),
  );
};

/** Heading style families across Word UI languages (en, cs/sk, de, fr, pl). */
const HEADING_STYLE_PATTERN = /heading|nadpis|berschrift|titre|nag[łl]/iu;

const classifyParagraph = (
  paragraph: Paragraph,
  styleById: Map<string, Style>,
): BilingualRowKind => {
  const formatting = paragraph.formatting;
  const style = formatting?.styleId ? styleById.get(formatting.styleId) : undefined;
  const outlineLevel = formatting?.outlineLevel ?? resolveInheritedOutlineLevel(style, styleById);
  if (outlineLevel !== undefined && outlineLevel < 9) {
    return "heading";
  }
  if (
    style &&
    (HEADING_STYLE_PATTERN.test(style.styleId) || HEADING_STYLE_PATTERN.test(style.name ?? ""))
  ) {
    return "heading";
  }
  if (effectiveNumPr(paragraph, styleById) !== undefined) {
    return "listItem";
  }
  return "paragraph";
};

const resolveInheritedOutlineLevel = (
  style: Style | undefined,
  styleById: Map<string, Style>,
): number | undefined => {
  const seen = new Set<string>();
  let current = style;
  while (current && !seen.has(current.styleId)) {
    seen.add(current.styleId);
    if (current.pPr?.outlineLevel !== undefined) {
      return current.pPr.outlineLevel;
    }
    current = current.basedOn ? styleById.get(current.basedOn) : undefined;
  }
  return undefined;
};

type NumPr = NonNullable<ParagraphFormatting["numPr"]>;

/** The numbering a paragraph renders with: direct `numPr`, else the style chain's. */
const effectiveNumPr = (paragraph: Paragraph, styleById: Map<string, Style>): NumPr | undefined => {
  const direct = paragraph.formatting?.numPr;
  if (direct?.numId !== undefined) {
    return direct.numId === 0 ? undefined : direct;
  }
  const styleId = paragraph.formatting?.styleId;
  return styleId ? styleNumPr(styleById.get(styleId), styleById) : undefined;
};

const styleNumPr = (style: Style | undefined, styleById: Map<string, Style>): NumPr | undefined => {
  const seen = new Set<string>();
  let current = style;
  while (current && !seen.has(current.styleId)) {
    seen.add(current.styleId);
    const numPr = current.pPr?.numPr;
    if (numPr?.numId !== undefined) {
      return numPr.numId === 0 ? undefined : numPr;
    }
    current = current.basedOn ? styleById.get(current.basedOn) : undefined;
  }
  return undefined;
};

// ----------------------------------------------------------------------------
// Numbering clones
// ----------------------------------------------------------------------------

type NumberingCloner = {
  /** Cloned `numId` for a source `numId`; minted on first use. */
  cloneNumId: (numId: number) => number;
  /** Cloned `abstractNumId` for a source one, when it was cloned. */
  clonedAbstractNumId: (abstractNumId: number) => number | undefined;
  hasClones: () => boolean;
  toDefinitions: () => NumberingDefinitions;
};

type CreateNumberingClonerOptions = {
  numbering: NumberingDefinitions | undefined;
  styleById: Map<string, Style>;
  warnings: string[];
};

const createNumberingCloner = ({
  numbering,
  styleById,
  warnings,
}: CreateNumberingClonerOptions): NumberingCloner => {
  const abstractNums = numbering?.abstractNums ?? [];
  const nums = numbering?.nums ?? [];
  const abstractById = new Map(abstractNums.map((item) => [item.abstractNumId, item]));
  const numById = new Map(nums.map((item) => [item.numId, item]));
  let nextAbstractNumId = Math.max(0, ...abstractNums.map((item) => item.abstractNumId)) + 1;
  let nextNumId = Math.max(0, ...nums.map((item) => item.numId)) + 1;

  const clonedAbstract = new Map<number, AbstractNumbering>();
  const clonedNum = new Map<number, NumberingInstance>();

  /**
   * Word keys list counters by the abstract definition a `w:num` points at;
   * two instances sharing one abstract continue the same sequence. A clone
   * therefore needs its own abstract, and an abstract that only links to a
   * numbering style must be materialized from that style's levels, otherwise
   * both clones resolve to the same linked definition and share counters.
   */
  const cloneAbstract = (sourceId: number): AbstractNumbering | undefined => {
    const existing = clonedAbstract.get(sourceId);
    if (existing) {
      return existing;
    }
    const source = abstractById.get(sourceId);
    if (!source) {
      return undefined;
    }
    const resolved = resolveLinkedAbstract(source);
    const { numStyleLink: _numStyleLink, styleLink: _styleLink, ...rest } = resolved;
    const clone: AbstractNumbering = {
      ...rest,
      abstractNumId: nextAbstractNumId,
      levels: structuredClone(resolved.levels),
    };
    nextAbstractNumId += 1;
    clonedAbstract.set(sourceId, clone);
    return clone;
  };

  const resolveLinkedAbstract = (abstract: AbstractNumbering): AbstractNumbering => {
    const seen = new Set<number>();
    let current = abstract;
    while (current.numStyleLink && !seen.has(current.abstractNumId)) {
      seen.add(current.abstractNumId);
      const linkedStyle = styleById.get(current.numStyleLink);
      const linkedNumId = linkedStyle?.pPr?.numPr?.numId;
      const linkedNum = linkedNumId === undefined ? undefined : numById.get(linkedNumId);
      const linkedAbstract = linkedNum ? abstractById.get(linkedNum.abstractNumId) : undefined;
      if (!linkedAbstract) {
        warnings.push(
          `Numbering style link "${current.numStyleLink}" on abstractNum ${current.abstractNumId} could not be resolved; the clone keeps the link.`,
        );
        return current;
      }
      current = linkedAbstract;
    }
    return current;
  };

  const cloneNumId = (numId: number): number => {
    const existing = clonedNum.get(numId);
    if (existing) {
      return existing.numId;
    }
    const source = numById.get(numId);
    if (!source) {
      warnings.push(
        `Numbering instance ${numId} is not defined; paragraphs using it keep the source instance.`,
      );
      return numId;
    }
    const abstract = cloneAbstract(source.abstractNumId);
    if (!abstract) {
      warnings.push(
        `Numbering instance ${numId} references abstractNum ${source.abstractNumId}, which is not defined; its copy shares the source counters.`,
      );
    }
    const clone: NumberingInstance = {
      ...source,
      numId: nextNumId,
      abstractNumId: abstract ? abstract.abstractNumId : source.abstractNumId,
    };
    nextNumId += 1;
    clonedNum.set(numId, clone);
    return clone.numId;
  };

  return {
    cloneNumId,
    clonedAbstractNumId: (abstractNumId) => clonedAbstract.get(abstractNumId)?.abstractNumId,
    hasClones: () => clonedNum.size > 0,
    toDefinitions: () => ({
      abstractNums: [...abstractNums, ...clonedAbstract.values()],
      nums: [...nums, ...clonedNum.values()],
    }),
  };
};

// ----------------------------------------------------------------------------
// Style clones
// ----------------------------------------------------------------------------

type StyleCloner = {
  /** Cloned style id for a source style id, or the source id when no clone is needed. */
  styleIdFor: (styleId: string) => string;
  hasClones: () => boolean;
  toDefinitions: (styles: StyleDefinitions) => StyleDefinitions;
};

type CreateStyleClonerOptions = {
  styleById: Map<string, Style>;
  suffix: string;
  cloner: NumberingCloner;
};

/**
 * A paragraph style is cloned only when its chain carries numbering. The clone
 * keeps `basedOn` and every other property; only `pPr.numPr` is rewritten to
 * the cloned instance, so indent precedence stays "style-sourced" exactly as
 * in the source (see `ParagraphFormatting.numPrFromStyle`).
 */
const createStyleCloner = ({
  styleById,
  suffix,
  cloner,
}: CreateStyleClonerOptions): StyleCloner => {
  const clones = new Map<string, Style>();

  const styleIdFor = (styleId: string): string => {
    const existing = clones.get(styleId);
    if (existing) {
      return existing.styleId;
    }
    const style = styleById.get(styleId);
    if (!style || style.type !== "paragraph") {
      return styleId;
    }
    const numPr = styleNumPr(style, styleById);
    if (numPr?.numId === undefined) {
      return styleId;
    }
    const cloneId = `${styleId}-${suffix}`;
    if (styleById.has(cloneId)) {
      return cloneId;
    }
    const clone: Style = {
      ...style,
      styleId: cloneId,
      name: `${style.name ?? style.styleId} (${suffix})`,
      ...(style.next !== undefined && { next: style.next === styleId ? cloneId : style.next }),
      default: false,
      pPr: { ...style.pPr, numPr: { ...numPr, numId: cloner.cloneNumId(numPr.numId) } },
    };
    clones.set(styleId, clone);
    return cloneId;
  };

  return {
    styleIdFor,
    hasClones: () => clones.size > 0,
    toDefinitions: (styles) => ({ ...styles, styles: [...styles.styles, ...clones.values()] }),
  };
};

// ----------------------------------------------------------------------------
// Paragraph / table copies
// ----------------------------------------------------------------------------

const cloneParagraphForTarget = (
  paragraph: Paragraph,
  targetParaId: string,
  styleCloner: StyleCloner,
  cloner: NumberingCloner,
  bookmarkIds: BookmarkIdMinter,
): Paragraph => {
  const { textId: _textId, sectionProperties: _sectionProperties, ...rest } = paragraph;
  const formatting = paragraph.formatting;
  const nextFormatting: ParagraphFormatting | undefined = formatting && {
    ...formatting,
    ...(formatting.styleId !== undefined && {
      styleId: styleCloner.styleIdFor(formatting.styleId),
    }),
    ...(formatting.numPr?.numId !== undefined &&
      formatting.numPr.numId !== 0 && {
        numPr: { ...formatting.numPr, numId: cloner.cloneNumId(formatting.numPr.numId) },
      }),
    ...(formatting.numPrFromStyle?.numId !== undefined &&
      formatting.numPrFromStyle.numId !== 0 && {
        numPrFromStyle: {
          ...formatting.numPrFromStyle,
          numId: cloner.cloneNumId(formatting.numPrFromStyle.numId),
        },
      }),
  };
  return {
    ...rest,
    // Own content nodes: repacking assigns rIds to images and hyperlinks in
    // place, which must not hit one shared node graph twice.
    content: remapClonedBookmarkIds(structuredClone(paragraph.content), bookmarkIds),
    paraId: targetParaId,
    ...(nextFormatting && { formatting: nextFormatting }),
    ...(paragraph.listRendering && {
      listRendering: remapListRendering(paragraph.listRendering, cloner),
    }),
  };
};

type BookmarkIdMinter = {
  mint: (sourceId: number) => number;
};

const createBookmarkIdMinter = (source: unknown): BookmarkIdMinter => {
  let nextId = 0;
  const remapped = new Map<number, number>();
  const visit = (value: unknown, seen: Set<object>): void => {
    if (typeof value !== "object" || value === null || seen.has(value)) return;
    seen.add(value);
    if (Array.isArray(value)) {
      value.forEach((item) => visit(item, seen));
      return;
    }
    const record = value as Record<string, unknown>;
    if (record["type"] === "bookmarkStart" || record["type"] === "bookmarkEnd") {
      const id = record["id"];
      if (typeof id === "number") nextId = Math.max(nextId, id + 1);
    }
    Object.values(record).forEach((item) => visit(item, seen));
  };
  visit(source, new Set());
  return {
    mint: (sourceId) => {
      const existing = remapped.get(sourceId);
      if (existing !== undefined) return existing;
      const id = nextId++;
      remapped.set(sourceId, id);
      return id;
    },
  };
};

const remapClonedBookmarkIds = <T>(value: T, bookmarkIds: BookmarkIdMinter): T => {
  const visit = (item: unknown, seen: Set<object>): void => {
    if (typeof item !== "object" || item === null || seen.has(item)) return;
    seen.add(item);
    if (Array.isArray(item)) {
      item.forEach((child) => visit(child, seen));
      return;
    }
    const record = item as Record<string, unknown>;
    if (record["type"] === "bookmarkStart" || record["type"] === "bookmarkEnd") {
      const id = record["id"];
      if (typeof id === "number") record["id"] = bookmarkIds.mint(id);
    }
    Object.values(record).forEach((child) => visit(child, seen));
  };
  visit(value, new Set());
  return value;
};

const remapListRendering = (rendering: ListRendering, cloner: NumberingCloner): ListRendering => {
  const clonedAbstract =
    rendering.abstractNumId === undefined
      ? undefined
      : cloner.clonedAbstractNumId(rendering.abstractNumId);
  return {
    ...rendering,
    numId: cloner.cloneNumId(rendering.numId),
    ...(clonedAbstract !== undefined && { abstractNumId: clonedAbstract }),
  };
};

const collectTableParagraphs = (table: Table): Paragraph[] => {
  const out: Paragraph[] = [];
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const item of cell.content) {
        if (item.type === "paragraph") {
          out.push(item);
        } else {
          out.push(...collectTableParagraphs(item));
        }
      }
    }
  }
  return out;
};

type CloneTableForTargetOptions = {
  table: Table;
  editableParagraphIds: ReadonlySet<string>;
  paraIds: ParaIdMinter;
  styleCloner: StyleCloner;
  cloner: NumberingCloner;
  bookmarkIds: BookmarkIdMinter;
};

type CloneTableForTargetResult = {
  table: Table;
  paragraphs: BilingualParagraphRef[];
};

const cloneTableForTarget = ({
  table,
  editableParagraphIds,
  paraIds,
  styleCloner,
  cloner,
  bookmarkIds,
}: CloneTableForTargetOptions): CloneTableForTargetResult => {
  const paragraphs: BilingualParagraphRef[] = [];
  const cloneTable = (source: Table): Table => ({
    ...structuredClone(source),
    rows: source.rows.map((row) => ({
      ...structuredClone(row),
      cells: row.cells.map((cell) => ({
        ...structuredClone(cell),
        content: cell.content.map((item) => {
          if (item.type === "table") {
            return cloneTable(item);
          }
          const targetParaId = paraIds.mint(item.paraId);
          const copy = cloneParagraphForTarget(
            item,
            targetParaId,
            styleCloner,
            cloner,
            bookmarkIds,
          );
          if (item.paraId !== undefined && editableParagraphIds.has(item.paraId)) {
            paragraphs.push({
              sourceParaId: item.paraId,
              targetParaId,
              sourceText: getParagraphText(item),
            });
          }
          return copy;
        }),
      })),
    })),
  });

  return { table: cloneTable(table), paragraphs };
};

// ----------------------------------------------------------------------------
// Output table
// ----------------------------------------------------------------------------

const buildRow = (
  left: Paragraph,
  right: Paragraph,
  styleById: Map<string, Style>,
  textWidth: number,
): Table["rows"][number] => {
  const columnWidth = Math.floor(textWidth / 2);
  const geometry = resolveHorizontalParagraphGeometry(left, styleById);
  return {
    type: "tableRow",
    formatting: { cantSplit: true },
    cells: [
      buildCell(projectParagraphIntoColumn(left, geometry, textWidth, columnWidth)),
      buildCell(projectParagraphIntoColumn(right, geometry, textWidth, columnWidth)),
    ],
  };
};

type HorizontalParagraphGeometry = Pick<
  ParagraphFormatting,
  "indentLeft" | "indentRight" | "indentFirstLine" | "hangingIndent" | "tabs"
>;

const HORIZONTAL_PARAGRAPH_KEYS = [
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
  "tabs",
] as const satisfies readonly (keyof HorizontalParagraphGeometry)[];

/** Resolve only the paragraph properties whose coordinates change when a
 * full-width paragraph is placed in a half-width cell. Direct pPr wins over
 * the basedOn style chain, matching Word's paragraph-style cascade. */
const resolveHorizontalParagraphGeometry = (
  paragraph: Paragraph,
  styleById: Map<string, Style>,
): HorizontalParagraphGeometry => {
  const chain: Style[] = [];
  const seen = new Set<string>();
  let style = paragraph.formatting?.styleId
    ? styleById.get(paragraph.formatting.styleId)
    : undefined;
  while (style && !seen.has(style.styleId)) {
    seen.add(style.styleId);
    chain.push(style);
    style = style.basedOn ? styleById.get(style.basedOn) : undefined;
  }

  const geometry: HorizontalParagraphGeometry = {};
  for (const current of chain.toReversed()) {
    assignHorizontalParagraphGeometry(geometry, current.pPr);
  }
  assignHorizontalParagraphGeometry(geometry, paragraph.formatting);
  return geometry;
};

const assignHorizontalParagraphGeometry = (
  target: HorizontalParagraphGeometry,
  source: ParagraphFormatting | undefined,
): void => {
  for (const key of HORIZONTAL_PARAGRAPH_KEYS) {
    const value = source?.[key];
    if (value !== undefined) {
      Object.assign(target, { [key]: value });
    }
  }
};

const projectParagraphIntoColumn = (
  paragraph: Paragraph,
  geometry: HorizontalParagraphGeometry,
  sourceWidth: number,
  columnWidth: number,
): Paragraph => {
  const scale = columnWidth / sourceWidth;
  const maxSideIndent = Math.max(0, columnWidth - MIN_COLUMN_TEXT_WIDTH_TWIPS);
  let indentLeft = projectSideIndent(geometry.indentLeft, scale, maxSideIndent);
  let indentRight = projectSideIndent(geometry.indentRight, scale, maxSideIndent);
  const indentFirstLine =
    geometry.indentFirstLine === undefined
      ? undefined
      : Math.max(-maxSideIndent, Math.round(geometry.indentFirstLine * scale));
  const minimumLeftIndent = Math.min(maxSideIndent, Math.max(0, -(indentFirstLine ?? 0)));
  indentLeft = Math.max(indentLeft, minimumLeftIndent);
  const flexibleLeft = indentLeft - minimumLeftIndent;
  const flexibleTotal = flexibleLeft + indentRight;
  const flexibleMaximum = maxSideIndent - minimumLeftIndent;
  if (flexibleTotal > flexibleMaximum) {
    indentLeft = minimumLeftIndent + Math.round((flexibleLeft / flexibleTotal) * flexibleMaximum);
    indentRight = maxSideIndent - indentLeft;
  }

  const formatting: ParagraphFormatting = {
    ...paragraph.formatting,
    ...((geometry.indentLeft !== undefined || minimumLeftIndent > 0) && { indentLeft }),
    ...(geometry.indentRight !== undefined && { indentRight }),
    ...(indentFirstLine !== undefined && { indentFirstLine }),
    ...(geometry.hangingIndent !== undefined && { hangingIndent: geometry.hangingIndent }),
    ...(geometry.tabs !== undefined && {
      tabs: geometry.tabs.map((tab) => ({
        ...tab,
        position: Math.min(
          Math.max(0, Math.round(tab.position * scale)),
          Math.max(0, columnWidth - MIN_TAB_TRAILING_WIDTH_TWIPS),
        ),
      })),
    }),
  };

  return { ...paragraph, formatting };
};

const projectSideIndent = (value: number | undefined, scale: number, maximum: number): number =>
  Math.min(Math.max(0, Math.round((value ?? 0) * scale)), maximum);

const buildCell = (paragraph: Paragraph): TableCell => ({
  type: "tableCell",
  formatting: { width: { value: HALF_WIDTH_PCT, type: "pct" }, verticalAlign: "top" },
  content: [paragraph],
});

/** A source table kept once, across both columns. */
const buildInlineTableRow = (table: Table): Table["rows"][number] => ({
  type: "tableRow",
  formatting: { cantSplit: true },
  cells: [
    {
      type: "tableCell",
      formatting: {
        width: { value: FULL_WIDTH_PCT, type: "pct" },
        gridSpan: 2,
        verticalAlign: "top",
      },
      // A cell must end with a paragraph; a bare nested table is invalid OOXML.
      content: [table, { type: "paragraph", content: [] }],
    },
  ],
});

/** Full-width source and target tables, stacked without narrowing either one. */
type BuildStackedTableRowOptions = { source: Table; target: Table };

const buildStackedTableRow = ({
  source,
  target,
}: BuildStackedTableRowOptions): Table["rows"][number] => ({
  type: "tableRow",
  formatting: { cantSplit: false },
  cells: [
    {
      type: "tableCell",
      formatting: {
        width: { value: FULL_WIDTH_PCT, type: "pct" },
        gridSpan: 2,
        verticalAlign: "top",
      },
      content: [
        source,
        { type: "paragraph", content: [] },
        target,
        { type: "paragraph", content: [] },
      ],
    },
  ],
});

const buildTable = (rows: Table["rows"], borders: BilingualBorders, textWidth: number): Table => ({
  type: "table",
  formatting: {
    width: { value: FULL_WIDTH_PCT, type: "pct" },
    styleId: BILINGUAL_TABLE_STYLE_ID,
    layout: "fixed",
    borders: TABLE_BORDERS[borders],
    look: { firstRow: false, firstColumn: false, noHBand: true, noVBand: true },
  },
  columnWidths: [Math.floor(textWidth / 2), Math.ceil(textWidth / 2)],
  rows,
});

const resolveTextWidthTwips = (doc: Document): number => {
  const section: SectionProperties | undefined =
    doc.package.document.finalSectionProperties ?? doc.package.document.sections?.at(0)?.properties;
  if (!section?.pageWidth) {
    return A4_TEXT_WIDTH_TWIPS;
  }
  const width = section.pageWidth - (section.marginLeft ?? 0) - (section.marginRight ?? 0);
  return width > 0 ? width : A4_TEXT_WIDTH_TWIPS;
};

// ----------------------------------------------------------------------------
// paraIds
// ----------------------------------------------------------------------------

/**
 * Every `paraId` anywhere in the package (body, headers, footers, notes,
 * comments), so a minted id cannot collide with a part the body never sees.
 * Walks the model generically: any object with `type: "paragraph"` and a
 * string `paraId` counts.
 */
const collectPackageParaIds = (pkg: Document["package"]): Set<string> => {
  const ids = new Set<string>();
  const seen = new Set<object>();
  const visit = (value: unknown): void => {
    if (typeof value !== "object" || value === null || seen.has(value)) {
      return;
    }
    seen.add(value);
    if (ArrayBuffer.isView(value) || value instanceof ArrayBuffer) {
      return;
    }
    if (Array.isArray(value)) {
      for (const item of value) {
        visit(item);
      }
      return;
    }
    if (value instanceof Map) {
      for (const item of value.values()) {
        visit(item);
      }
      return;
    }
    if (isParagraphWithId(value)) {
      ids.add(value.paraId);
    }
    for (const child of Object.values(value)) {
      visit(child);
    }
  };
  visit(pkg);
  return ids;
};

type ParaIdMinter = { mint: (sourceParaId: string | undefined) => string };

/**
 * Fresh ids derived from the source id (so re-running on the same source
 * yields the same handles), salted past any id already in the document.
 */
const createParaIdMinter = (taken: Set<string>): ParaIdMinter => {
  let ordinal = 0;
  return {
    mint: (sourceParaId) => {
      ordinal += 1;
      const seed = `${ROW_ID_NAMESPACE}:${sourceParaId ?? `ordinal-${ordinal}`}`;
      let id = deterministicHexId(seed);
      for (let salt = 1; taken.has(id); salt += 1) {
        id = deterministicHexId(`${seed}:${salt}`);
      }
      taken.add(id);
      return id;
    },
  };
};

// ----------------------------------------------------------------------------
// Reading a bilingual document back
// ----------------------------------------------------------------------------

/**
 * Re-derive the row manifest from a document produced by
 * {@link createBilingualDocument}. A dedicated table-style discriminator keeps
 * ordinary two-column source tables from being mistaken for bilingual output;
 * the expected row structure is validated as a second check. Rows are returned
 * in document order.
 */
export function readBilingualDocument(
  document: Document,
  editableParagraphIds: ReadonlySet<string>,
): BilingualRow[] {
  const styleById = new Map(
    (document.package.styles?.styles ?? []).map((style) => [style.styleId, style]),
  );
  const rows: BilingualRow[] = [];
  let missingHandleCount = 0;
  for (const block of flattenBlocks(document.package.document.content)) {
    if (block.type !== "table" || !isBilingualTable(block)) {
      continue;
    }
    for (const row of block.rows) {
      const [left, right] = row.cells;
      if (!left) {
        continue;
      }
      if (!right) {
        const nestedTables = left.content.filter((item): item is Table => item.type === "table");
        const sourceTable = nestedTables.at(0);
        if (!sourceTable) {
          missingHandleCount += 1;
          continue;
        }
        if (nestedTables.length === 2) {
          const targetTable = nestedTables.at(1);
          if (!targetTable) {
            continue;
          }
          const sourceParagraphs = collectTableParagraphs(sourceTable);
          const targetParagraphs = collectTableParagraphs(targetTable);
          if (sourceParagraphs.length !== targetParagraphs.length) {
            missingHandleCount += 1;
            continue;
          }
          const paragraphs: BilingualParagraphRef[] = [];
          for (const [index, source] of sourceParagraphs.entries()) {
            if (source.paraId === undefined || !editableParagraphIds.has(source.paraId)) {
              continue;
            }
            const target = targetParagraphs.at(index);
            if (
              target?.paraId === undefined ||
              target.paraId === source.paraId ||
              !editableParagraphIds.has(target.paraId)
            ) {
              missingHandleCount += 1;
              continue;
            }
            paragraphs.push({
              sourceParaId: source.paraId,
              targetParaId: target.paraId,
              sourceText: getParagraphText(source),
            });
          }
          rows.push({
            kind: "table",
            layout: "stacked",
            rowId: paragraphs.at(0)?.targetParaId ?? tableRowHandle(rows.length),
            paragraphs,
          });
          continue;
        }
        if (nestedTables.length !== 1) {
          missingHandleCount += 1;
          continue;
        }
        const paragraphs = collectTableParagraphs(sourceTable)
          .filter(
            (paragraph): paragraph is Paragraph & { paraId: string } =>
              paragraph.paraId !== undefined && editableParagraphIds.has(paragraph.paraId),
          )
          .map((paragraph) => ({
            paraId: paragraph.paraId,
            sourceText: getParagraphText(paragraph),
          }));
        rows.push({
          kind: "table",
          layout: "inline",
          rowId: paragraphs.at(0)?.paraId ?? tableRowHandle(rows.length),
          paragraphs,
        });
        continue;
      }
      const source = left.content.at(0);
      const target = right.content.at(0);
      if (source?.type !== "paragraph" || target?.type !== "paragraph") {
        continue;
      }
      if (target.paraId === undefined) {
        missingHandleCount += 1;
        continue;
      }
      const sourceMissing = source.paraId === undefined || !editableParagraphIds.has(source.paraId);
      const targetMissing = !editableParagraphIds.has(target.paraId);
      if (sourceMissing) {
        missingHandleCount += 1;
      }
      if (targetMissing) {
        missingHandleCount += 1;
      }
      if (sourceMissing || targetMissing) {
        continue;
      }
      rows.push({
        kind: classifyParagraph(source, styleById),
        rowId: target.paraId,
        sourceParaId: source.paraId,
        targetParaId: target.paraId,
        sourceText: getParagraphText(source),
      });
    }
  }
  if (missingHandleCount > 0) {
    throw new UneditableBilingualManifestError({
      message: "Bilingual manifest structure or handles do not match the Folio AI-edit snapshot.",
      missingHandleCount,
    });
  }
  return rows;
}

const isBilingualTable = (table: Table): boolean => {
  if (table.formatting?.styleId !== BILINGUAL_TABLE_STYLE_ID) {
    return false;
  }
  for (const row of table.rows) {
    const cells = row.cells;
    if (cells.length === 2) {
      const singleParagraphCells = cells.every(
        (cell) => cell.content.length === 1 && cell.content[0]?.type === "paragraph",
      );
      if (!singleParagraphCells) {
        return false;
      }
      continue;
    }
    if (cells.length === 1 && cells[0]?.formatting?.gridSpan === 2) {
      continue;
    }
    return false;
  }
  return table.rows.length > 0;
};

const isParagraphWithId = (value: object): value is { type: "paragraph"; paraId: string } =>
  "type" in value &&
  value.type === "paragraph" &&
  "paraId" in value &&
  typeof value.paraId === "string";
