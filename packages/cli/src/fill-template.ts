/**
 * `folio fill`: author a document from a template and structured content.
 * The template keeps its sections, headers, footers, styles, numbering, and
 * the body blocks the caller keeps (the front matter, and optionally a
 * trailing block such as a signature); everything between is replaced, in
 * one direct write, by the content, given as Markdown or as a block list and
 * styled through a map from content constructs to the template's paragraph
 * styles and numbering. Fields then fill `{{name}}` placeholders, and exact
 * text replacements apply, in every story.
 */

import { panic, Result } from "better-result";

import {
  type BlockContent,
  compileMarkdownToContent,
  type ParagraphContent,
} from "@stll/docx-core";
import { SEQUENCE_CHILDREN } from "@stll/docx-core/schema";
import type { FolioDocxReviewer } from "@stll/folio-core/server";

import { cliError, FOLIO_CLI_ERROR_CODES, type FolioCliError } from "./errors";
import {
  attributeNamespace,
  childElements,
  createWElement,
  insertChildInOrder,
  isW,
  localNameOf,
  mainDocumentPart,
  type OoxmlPackage,
  RELATIONSHIP_TYPES,
  storyParts,
  walkElements,
  wAttribute,
  type XmlElement,
} from "./ooxml-package";
import { paragraphText, replaceParagraphText } from "./paragraph-text";

/** The content constructs a style map can style. */
export const CONTENT_CONSTRUCTS: readonly [
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "paragraph",
  "numberedParagraph",
  "list",
  "quote",
  "table",
] = [
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "paragraph",
  "numberedParagraph",
  "list",
  "quote",
  "table",
] as const;

type Construct = (typeof CONTENT_CONSTRUCTS)[number];

const HEADING_CONSTRUCTS = ["h1", "h2", "h3", "h4", "h5", "h6"] as const;

/** How one construct is styled: a paragraph (or table) style, and numbering. */
type ConstructStyle = { style?: string; numId?: number; level?: number };

type InlineRun = { text: string; bold?: boolean; italic?: boolean };

/** One block of content, whatever form it arrived in. */
type ContentBlock =
  | { type: "heading"; level: number; runs: InlineRun[] }
  | {
      type: "paragraph" | "numberedParagraph" | "quote";
      runs: InlineRun[];
      style?: string;
    }
  | { type: "list"; ordered: boolean; level: number; runs: InlineRun[] }
  | { type: "table"; rows: InlineRun[][][]; header: boolean };

export type FillReport = {
  /** Body blocks kept from the template: before the content, and after it. */
  kept: { front: number; tail: number };
  /** Template body blocks the content replaced. */
  removed: number;
  /** Body blocks the content added. */
  inserted: number;
  /** How many placeholders or texts each field and replacement filled. */
  fields: Record<string, number>;
  warnings: string[];
};

const PARAGRAPH_PROPERTIES = SEQUENCE_CHILDREN["paragraph-properties"];
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
/** Text width when the final section states no page size: A4 with 1 in margins. */
const DEFAULT_TEXT_WIDTH = 9026;
const MAX_PARAGRAPH_ID = 0x7f_ff_ff_ff;
const FIRST_PARAGRAPH_ID = 0x10_00_00_00;

const invalidInput = (message: string, hint?: string): FolioCliError =>
  cliError({ code: FOLIO_CLI_ERROR_CODES.invalidInput, message, hint });

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isConstruct = (value: string): value is Construct =>
  CONTENT_CONSTRUCTS.some((construct) => construct === value);

// --- arguments -----------------------------------------------------------------

const stringMap = (
  value: unknown,
  name: string,
): Result<ReadonlyMap<string, string>, FolioCliError> => {
  if (value === undefined) return Result.ok(new Map());
  if (!isRecord(value))
    return Result.err(invalidInput(`${name} must be an object of strings.`));
  const map = new Map<string, string>();
  for (const [key, entry] of Object.entries(value)) {
    if (typeof entry !== "string" || key === "")
      return Result.err(invalidInput(`${name}.${key} must be a string value.`));
    map.set(key, entry);
  }
  return Result.ok(map);
};

const styleMap = (
  value: unknown,
): Result<ReadonlyMap<Construct, ConstructStyle>, FolioCliError> => {
  if (value === undefined) return Result.ok(new Map());
  if (!isRecord(value))
    return Result.err(invalidInput("styles must be an object."));
  const map = new Map<Construct, ConstructStyle>();
  for (const [key, entry] of Object.entries(value)) {
    if (!isConstruct(key)) {
      return Result.err(
        invalidInput(
          `styles names ${JSON.stringify(key)}; expected ${CONTENT_CONSTRUCTS.join(", ")}.`,
        ),
      );
    }
    if (typeof entry === "string") {
      map.set(key, { style: entry });
      continue;
    }
    if (!isRecord(entry))
      return Result.err(
        invalidInput(`styles.${key} must be a style id or an object.`),
      );
    const { style, numId, level } = entry;
    if (
      (style !== undefined && typeof style !== "string") ||
      (numId !== undefined && !Number.isInteger(numId)) ||
      (level !== undefined && !Number.isInteger(level))
    ) {
      return Result.err(
        invalidInput(
          `styles.${key} takes { style?: string, numId?: integer, level?: integer }.`,
        ),
      );
    }
    map.set(key, {
      ...(typeof style === "string" && { style }),
      ...(typeof numId === "number" && { numId }),
      ...(typeof level === "number" && { level }),
    });
  }
  return Result.ok(map);
};

const runsFrom = (
  value: unknown,
  where: string,
): Result<InlineRun[], FolioCliError> => {
  if (typeof value === "string") return Result.ok([{ text: value }]);
  if (!Array.isArray(value))
    return Result.err(invalidInput(`${where} needs text or runs.`));
  const runs: InlineRun[] = [];
  for (const run of value) {
    if (typeof run === "string") {
      runs.push({ text: run });
      continue;
    }
    if (!isRecord(run) || typeof run["text"] !== "string")
      return Result.err(
        invalidInput(
          `${where}: each run is a string or { text, bold?, italic? }.`,
        ),
      );
    runs.push({
      text: run["text"],
      ...(run["bold"] === true && { bold: true }),
      ...(run["italic"] === true && { italic: true }),
    });
  }
  return Result.ok(runs);
};

const BLOCK_TYPES = [
  "heading",
  "paragraph",
  "numberedParagraph",
  "quote",
  "bullet",
  "numbered",
  "table",
] as const;

const isBlockType = (value: unknown): value is (typeof BLOCK_TYPES)[number] =>
  BLOCK_TYPES.some((type) => type === value);

/** The JSON block list form of the content. */
const blocksFromList = (
  value: unknown,
): Result<ContentBlock[], FolioCliError> => {
  if (!Array.isArray(value))
    return Result.err(invalidInput("blocks must be a list."));
  const blocks: ContentBlock[] = [];
  for (const [index, block] of value.entries()) {
    const where = `blocks[${String(index)}]`;
    const type = isRecord(block) ? block["type"] : undefined;
    if (!isRecord(block) || !isBlockType(type)) {
      return Result.err(
        invalidInput(`${where} needs a type: ${BLOCK_TYPES.join(", ")}.`),
      );
    }
    if (type === "table") {
      const rows = block["rows"];
      if (!Array.isArray(rows) || rows.some((row) => !Array.isArray(row)))
        return Result.err(invalidInput(`${where}.rows must be rows of cells.`));
      const parsedRows: InlineRun[][][] = [];
      for (const row of rows) {
        const cells: InlineRun[][] = [];
        for (const cell of Array.isArray(row) ? row : []) {
          const runs = runsFrom(cell, `${where} cell`);
          if (runs.isErr()) return Result.err(runs.error);
          cells.push(runs.value);
        }
        parsedRows.push(cells);
      }
      blocks.push({
        type: "table",
        rows: parsedRows,
        header: block["header"] === true,
      });
      continue;
    }
    const runs = runsFrom(block["runs"] ?? block["text"], where);
    if (runs.isErr()) return Result.err(runs.error);
    const level = block["level"] ?? (type === "heading" ? 1 : 0);
    if (
      !Number.isInteger(level) ||
      typeof level !== "number" ||
      level < 0 ||
      level > 8
    )
      return Result.err(invalidInput(`${where}.level must be an integer.`));
    switch (type) {
      case "heading":
        if (level < 1 || level > 6)
          return Result.err(invalidInput(`${where}.level must be 1 to 6.`));
        blocks.push({ type, level, runs: runs.value });
        break;
      case "bullet":
      case "numbered":
        blocks.push({
          type: "list",
          ordered: type === "numbered",
          level,
          runs: runs.value,
        });
        break;
      default: {
        const style = block["style"];
        if (style !== undefined && typeof style !== "string")
          return Result.err(invalidInput(`${where}.style must be a style id.`));
        blocks.push({
          type,
          runs: runs.value,
          ...(typeof style === "string" && { style }),
        });
      }
    }
  }
  return Result.ok(blocks);
};

/** Inline runs of compiled Markdown, flattening links to their text. */
const runsOfContent = (content: readonly ParagraphContent[]): InlineRun[] => {
  const runs: InlineRun[] = [];
  for (const item of content) {
    if (item.type === "hyperlink") {
      runs.push(...runsOfContent(item.children));
      continue;
    }
    if (item.type !== "run") continue;
    const text = item.content
      .map((piece) => {
        switch (piece.type) {
          case "text":
            return piece.text;
          case "tab":
            return "\t";
          case "break":
            return "\n";
          default:
            return "";
        }
      })
      .join("");
    if (text === "") continue;
    runs.push({
      text,
      ...(item.formatting?.bold === true && { bold: true }),
      ...(item.formatting?.italic === true && { italic: true }),
    });
  }
  return runs;
};

const HEADING_STYLE = /^Heading([1-6])$/u;

/** The Markdown form of the content, through folio's own Markdown compiler. */
const blocksFromMarkdown = (markdown: string): ContentBlock[] => {
  const toBlock = (block: BlockContent): ContentBlock | null => {
    if (block.type === "table") {
      return {
        type: "table",
        header: true,
        rows: block.rows.map((row) =>
          row.cells.map((cell) =>
            cell.content.flatMap((inner) =>
              inner.type === "paragraph" ? runsOfContent(inner.content) : [],
            ),
          ),
        ),
      };
    }
    if (block.type !== "paragraph") return null;
    const runs = runsOfContent(block.content);
    const heading = HEADING_STYLE.exec(block.formatting?.styleId ?? "");
    if (heading?.[1] !== undefined)
      return { type: "heading", level: Number(heading[1]), runs };
    if (block.listRendering !== undefined) {
      return {
        type: "list",
        ordered: !block.listRendering.isBullet,
        level: block.listRendering.level,
        runs,
      };
    }
    if (block.formatting?.styleId === "Quote") return { type: "quote", runs };
    return { type: "paragraph", runs };
  };
  return compileMarkdownToContent(markdown).content.flatMap((block) => {
    const converted = toBlock(block);
    return converted === null ? [] : [converted];
  });
};

// --- the template ------------------------------------------------------------------

type TemplateFacts = {
  paragraphStyles: ReadonlySet<string>;
  tableStyles: ReadonlySet<string>;
  numIds: ReadonlySet<number>;
};

const templateFacts = (
  styles: XmlElement | null,
  numbering: XmlElement | null,
): TemplateFacts => {
  const paragraphStyles = new Set<string>();
  const tableStyles = new Set<string>();
  for (const style of childElements(styles ?? {})) {
    if (!isW(style, "style")) continue;
    const id = wAttribute(style, "styleId");
    if (id === undefined) continue;
    const type = wAttribute(style, "type");
    if (type === "paragraph") paragraphStyles.add(id);
    if (type === "table") tableStyles.add(id);
  }
  const numIds = new Set<number>();
  for (const instance of childElements(numbering ?? {})) {
    if (isW(instance, "num")) numIds.add(Number(wAttribute(instance, "numId")));
  }
  return { paragraphStyles, tableStyles, numIds };
};

/** Where a top-level body block sits relative to the boundaries. */
const matchesBoundary = (element: XmlElement, boundary: string): boolean => {
  const upper = boundary.toUpperCase();
  let matched = false;
  walkElements(element, (inner) => {
    if (matched || !isW(inner, "p")) return;
    const id = Object.entries(inner.attributes ?? {}).find(
      ([attribute]) =>
        localNameOf(attribute) === "paraId" &&
        attributeNamespace(inner, attribute) === W14,
    )?.[1];
    if (String(id ?? "").toUpperCase() === upper) matched = true;
    else if (paragraphText(inner).includes(boundary)) matched = true;
  });
  return matched;
};

const paraIdOf = (element: XmlElement): string | undefined => {
  const id = Object.entries(element.attributes ?? {}).find(
    ([attribute]) =>
      localNameOf(attribute) === "paraId" &&
      attributeNamespace(element, attribute) === W14,
  )?.[1];
  return id === undefined ? undefined : String(id).toUpperCase();
};

// --- writing ---------------------------------------------------------------------

type BlockBuilder = {
  body: XmlElement;
  w14Prefix: string | undefined;
  nextParaId: () => string;
  facts: TemplateFacts;
  styles: ReadonlyMap<Construct, ConstructStyle>;
  textWidth: number;
};

const runElement = (builder: BlockBuilder, run: InlineRun): XmlElement => {
  const element = createWElement(builder.body, "r");
  const children: XmlElement[] = [];
  if (run.bold === true || run.italic === true) {
    const properties = createWElement(element, "rPr");
    properties.elements = [
      ...(run.bold === true ? [createWElement(properties, "b")] : []),
      ...(run.italic === true ? [createWElement(properties, "i")] : []),
    ];
    children.push(properties);
  }
  const pieces = run.text.split(/(\t|\n)/u);
  for (const piece of pieces) {
    if (piece === "") continue;
    if (piece === "\t") {
      children.push(createWElement(element, "tab"));
      continue;
    }
    if (piece === "\n") {
      children.push(createWElement(element, "br"));
      continue;
    }
    const text = createWElement(element, "t");
    text.attributes = { "xml:space": "preserve" };
    text.elements = [{ type: "text", text: piece }];
    children.push(text);
  }
  element.elements = children;
  return element;
};

type ParagraphOptions = {
  style?: string | undefined;
  numbering?: { numId: number; level: number } | undefined;
  runs: readonly InlineRun[];
};

const paragraphElement = (
  builder: BlockBuilder,
  { style, numbering, runs }: ParagraphOptions,
): XmlElement => {
  const paragraph = createWElement(builder.body, "p");
  if (builder.w14Prefix !== undefined) {
    paragraph.attributes = {
      [`${builder.w14Prefix}:paraId`]: builder.nextParaId(),
    };
  }
  const children: XmlElement[] = [];
  if (style !== undefined || numbering !== undefined) {
    const properties = createWElement(paragraph, "pPr");
    properties.elements = [];
    if (style !== undefined) {
      insertChildInOrder(
        properties,
        createWElement(properties, "pStyle", { val: style }),
        PARAGRAPH_PROPERTIES,
      );
    }
    if (numbering !== undefined) {
      const numPr = createWElement(properties, "numPr");
      numPr.elements = [
        createWElement(numPr, "ilvl", { val: String(numbering.level) }),
        createWElement(numPr, "numId", { val: String(numbering.numId) }),
      ];
      insertChildInOrder(properties, numPr, PARAGRAPH_PROPERTIES);
    }
    children.push(properties);
  }
  for (const run of runs) children.push(runElement(builder, run));
  paragraph.elements = children;
  return paragraph;
};

const tableElement = (
  builder: BlockBuilder,
  block: Extract<ContentBlock, { type: "table" }>,
): XmlElement => {
  const table = createWElement(builder.body, "tbl");
  const columns = Math.max(1, ...block.rows.map((row) => row.length));
  const width = Math.floor(builder.textWidth / columns);
  const style = builder.styles.get("table")?.style;
  const properties = createWElement(table, "tblPr");
  properties.elements = [
    ...(style === undefined
      ? []
      : [createWElement(properties, "tblStyle", { val: style })]),
    createWElement(properties, "tblW", { w: "0", type: "auto" }),
  ];
  if (style === undefined) {
    const borders = createWElement(properties, "tblBorders");
    borders.elements = [
      "top",
      "left",
      "bottom",
      "right",
      "insideH",
      "insideV",
    ].map((edge) =>
      createWElement(borders, edge, {
        val: "single",
        sz: "4",
        space: "0",
        color: "auto",
      }),
    );
    properties.elements.push(borders);
  }
  const grid = createWElement(table, "tblGrid");
  grid.elements = Array.from({ length: columns }, () =>
    createWElement(grid, "gridCol", { w: String(width) }),
  );
  const rows = block.rows.map((cells, index) => {
    const row = createWElement(table, "tr");
    row.elements = [];
    if (index === 0 && block.header) {
      const rowProperties = createWElement(row, "trPr");
      rowProperties.elements = [createWElement(rowProperties, "tblHeader")];
      row.elements.push(rowProperties);
    }
    for (let column = 0; column < columns; column += 1) {
      const cell = createWElement(row, "tc");
      const cellProperties = createWElement(cell, "tcPr");
      cellProperties.elements = [
        createWElement(cellProperties, "tcW", {
          w: String(width),
          type: "dxa",
        }),
      ];
      cell.elements = [
        cellProperties,
        paragraphElement(builder, { runs: cells[column] ?? [] }),
      ];
      row.elements.push(cell);
    }
    return row;
  });
  table.elements = [properties, grid, ...rows];
  return table;
};

const resolvedStyle = (
  builder: BlockBuilder,
  construct: Construct,
  fallback: string | undefined,
): ConstructStyle => {
  const mapped = builder.styles.get(construct);
  if (mapped !== undefined) return mapped;
  return fallback !== undefined && builder.facts.paragraphStyles.has(fallback)
    ? { style: fallback }
    : {};
};

const blockElements = (
  builder: BlockBuilder,
  block: ContentBlock,
  warnings: Set<string>,
): XmlElement => {
  switch (block.type) {
    case "table":
      return tableElement(builder, block);
    case "heading": {
      const styled = resolvedStyle(
        builder,
        HEADING_CONSTRUCTS[block.level - 1] ?? "h1",
        `Heading${String(block.level)}`,
      );
      return paragraphElement(builder, {
        style: styled.style,
        numbering:
          styled.numId === undefined
            ? undefined
            : { numId: styled.numId, level: styled.level ?? 0 },
        runs: block.runs,
      });
    }
    case "list": {
      const construct = block.ordered ? "numberedParagraph" : "list";
      const styled =
        builder.styles.get(construct) ?? builder.styles.get("list") ?? {};
      if (styled.numId === undefined && styled.style === undefined) {
        warnings.add(
          `${block.ordered ? "Numbered" : "Bullet"} items have no style or numbering; map ${construct} in styles.`,
        );
      }
      return paragraphElement(builder, {
        style: styled.style,
        numbering:
          styled.numId === undefined
            ? undefined
            : { numId: styled.numId, level: (styled.level ?? 0) + block.level },
        runs: block.runs,
      });
    }
    case "paragraph":
    case "numberedParagraph":
    case "quote": {
      const styled = resolvedStyle(
        builder,
        block.type,
        block.type === "quote" ? "Quote" : undefined,
      );
      return paragraphElement(builder, {
        style: block.style ?? styled.style,
        numbering:
          styled.numId === undefined
            ? undefined
            : { numId: styled.numId, level: styled.level ?? 0 },
        runs: block.runs,
      });
    }
    default: {
      const unreachable: never = block;
      return panic("Unhandled content block", { unreachable });
    }
  }
};

/** Every style id and numbering instance the content uses, checked against the template. */
const checkReferences = (
  blocks: readonly ContentBlock[],
  styles: ReadonlyMap<Construct, ConstructStyle>,
  facts: TemplateFacts,
): Result<void, FolioCliError> => {
  for (const [construct, { style, numId }] of styles) {
    const known =
      construct === "table" ? facts.tableStyles : facts.paragraphStyles;
    if (style !== undefined && !known.has(style)) {
      return Result.err(
        invalidInput(
          `styles.${construct} names ${JSON.stringify(style)}, which the template defines no ${construct === "table" ? "table" : "paragraph"} style for.`,
        ),
      );
    }
    if (numId !== undefined && !facts.numIds.has(numId)) {
      return Result.err(
        invalidInput(
          `styles.${construct} names numbering instance ${String(numId)}, which the template does not define.`,
        ),
      );
    }
  }
  for (const block of blocks) {
    if (
      block.type !== "table" &&
      block.type !== "heading" &&
      block.type !== "list" &&
      block.style !== undefined &&
      !facts.paragraphStyles.has(block.style)
    ) {
      return Result.err(
        invalidInput(
          `A block names style ${JSON.stringify(block.style)}, which the template does not define.`,
        ),
      );
    }
  }
  return Result.ok(undefined);
};

/** The width text may take in the final section, in twentieths of a point. */
const textWidthOf = (body: XmlElement): number => {
  const section = childElements(body).findLast((child) => isW(child, "sectPr"));
  const size = childElements(section ?? {}).find((child) => isW(child, "pgSz"));
  const margins = childElements(section ?? {}).find((child) =>
    isW(child, "pgMar"),
  );
  const page = Number(wAttribute(size ?? {}, "w"));
  const left = Number(
    wAttribute(margins ?? {}, "left") ??
      wAttribute(margins ?? {}, "start") ??
      0,
  );
  const right = Number(
    wAttribute(margins ?? {}, "right") ?? wAttribute(margins ?? {}, "end") ?? 0,
  );
  const width = page - left - right;
  return Number.isFinite(width) && width > 0 ? width : DEFAULT_TEXT_WIDTH;
};

/** Fresh `w14:paraId` values, none of them already used in the package. */
const paragraphIds = (
  pkg: OoxmlPackage,
  parts: readonly string[],
): (() => string) => {
  const used = new Set<string>();
  for (const part of parts) {
    const root = pkg.root(part);
    if (root.isErr() || root.value === null) continue;
    walkElements(root.value, (element) => {
      const id = isW(element, "p") ? paraIdOf(element) : undefined;
      if (id !== undefined) used.add(id);
    });
  }
  let next = FIRST_PARAGRAPH_ID;
  return () => {
    while (used.has(next.toString(16).toUpperCase().padStart(8, "0")))
      next += 1;
    if (next > MAX_PARAGRAPH_ID) next = 1;
    const id = next.toString(16).toUpperCase().padStart(8, "0");
    used.add(id);
    next += 1;
    return id;
  };
};

/** Section breaks (`w:sectPr`) inside one body block. */
const sectionBreaksIn = (element: XmlElement | undefined): number => {
  let count = 0;
  if (element !== undefined) {
    walkElements(element, (inner) => {
      if (isW(inner, "sectPr")) count += 1;
    });
  }
  return count;
};

const FIELD_PLACEHOLDER = /\{\{\s*([^{}]+?)\s*\}\}/gu;

/** Fill `{{name}}` fields and exact-text replacements in every paragraph of a part. */
const fillText = (
  root: XmlElement,
  fields: ReadonlyMap<string, string>,
  replacements: ReadonlyMap<string, string>,
  counts: Record<string, number>,
): boolean => {
  let changed = false;
  walkElements(root, (paragraph) => {
    if (!isW(paragraph, "p")) return;
    const text = paragraphText(paragraph);
    const edits: { start: number; end: number; text: string; key: string }[] =
      [];
    for (const match of text.matchAll(FIELD_PLACEHOLDER)) {
      const key = match[1] ?? "";
      const value = fields.get(key);
      if (value === undefined) continue;
      edits.push({
        start: match.index,
        end: match.index + match[0].length,
        text: value,
        key,
      });
    }
    for (const [find, value] of replacements) {
      let from = text.indexOf(find);
      while (from !== -1) {
        edits.push({
          start: from,
          end: from + find.length,
          text: value,
          key: find,
        });
        from = text.indexOf(find, from + find.length);
      }
    }
    if (edits.length === 0) return;
    if (replaceParagraphText(paragraph, edits) > 0) changed = true;
    for (const { key } of edits) counts[key] = (counts[key] ?? 0) + 1;
  });
  return changed;
};

/** Build the document the arguments describe from the template package. */
export const fillTemplate = (
  pkg: OoxmlPackage,
  args: Readonly<Record<string, unknown>>,
  reviewer: FolioDocxReviewer,
): Result<FillReport, FolioCliError> => {
  const { markdown, blocks: blockList, keepUntil, keepFrom } = args;
  if ((markdown === undefined) === (blockList === undefined)) {
    return Result.err(
      invalidInput("Give the content as markdown or as blocks, exactly one."),
    );
  }
  if (markdown !== undefined && typeof markdown !== "string")
    return Result.err(invalidInput("markdown must be text."));
  if (
    keepUntil !== undefined &&
    (typeof keepUntil !== "string" || keepUntil === "")
  )
    return Result.err(invalidInput("keepUntil must be a block id or text."));
  if (
    keepFrom !== undefined &&
    (typeof keepFrom !== "string" || keepFrom === "")
  )
    return Result.err(invalidInput("keepFrom must be a block id or text."));
  const blocks =
    typeof markdown === "string"
      ? Result.ok(blocksFromMarkdown(markdown))
      : blocksFromList(blockList);
  if (blocks.isErr()) return Result.err(blocks.error);
  const styles = styleMap(args["styles"]);
  if (styles.isErr()) return Result.err(styles.error);
  const fields = stringMap(args["fields"], "fields");
  if (fields.isErr()) return Result.err(fields.error);
  const replacements = stringMap(args["replaceText"], "replaceText");
  if (replacements.isErr()) return Result.err(replacements.error);

  const main = mainDocumentPart(pkg);
  if (main.isErr()) return Result.err(main.error);
  if (main.value === null)
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: "The template has no main document.",
      }),
    );
  const document = pkg.root(main.value);
  if (document.isErr()) return Result.err(document.error);
  const body = childElements(document.value ?? {}).find((child) =>
    isW(child, "body"),
  );
  if (body === undefined)
    return Result.err(
      cliError({
        code: FOLIO_CLI_ERROR_CODES.invalidDocument,
        message: "The template has no body.",
      }),
    );
  const relationships = pkg.relationships(main.value);
  if (relationships.isErr()) return Result.err(relationships.error);
  const related = (type: string): XmlElement | null => {
    const part = relationships.value.find(
      (relationship) => relationship.type === type,
    )?.part;
    if (part === undefined || part === null) return null;
    const root = pkg.root(part);
    return root.isOk() ? root.value : null;
  };
  const facts = templateFacts(
    related(RELATIONSHIP_TYPES.styles),
    related(RELATIONSHIP_TYPES.numbering),
  );
  const checked = checkReferences(blocks.value, styles.value, facts);
  if (checked.isErr()) return Result.err(checked.error);

  // The body's blocks, without the final section properties.
  const children = body.elements ?? [];
  const blockIndexes = children.flatMap((child, index) =>
    child.type === "element" && !isW(child, "sectPr") ? [index] : [],
  );
  const headings = new Set(
    reviewer
      .snapshot()
      .blocks.filter(
        (block) => block.kind === "heading" && block.table === undefined,
      )
      .map(({ id }) => id.toUpperCase()),
  );
  let frontEnd: number;
  if (typeof keepUntil === "string") {
    const at = blockIndexes.findIndex((index) => {
      const child = children[index];
      return child !== undefined && matchesBoundary(child, keepUntil);
    });
    if (at === -1)
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.notFound,
          message: `No template block matches keepUntil ${JSON.stringify(keepUntil)}.`,
        }),
      );
    frontEnd = at + 1;
  } else {
    const at = blockIndexes.findIndex((index) => {
      const child = children[index];
      const id =
        child !== undefined && isW(child, "p") ? paraIdOf(child) : undefined;
      return id !== undefined && headings.has(id);
    });
    if (at === -1) {
      return Result.err(
        invalidInput(
          "The template has no heading to keep the front matter up to.",
          "Pass keepUntil with the id or text of the last block to keep.",
        ),
      );
    }
    frontEnd = at;
  }
  let tailStart = blockIndexes.length;
  if (typeof keepFrom === "string") {
    const at = blockIndexes.findIndex((index, position) => {
      const child = children[index];
      return (
        position >= frontEnd &&
        child !== undefined &&
        matchesBoundary(child, keepFrom)
      );
    });
    if (at === -1)
      return Result.err(
        cliError({
          code: FOLIO_CLI_ERROR_CODES.notFound,
          message: `No template block after the kept front matter matches keepFrom ${JSON.stringify(keepFrom)}.`,
        }),
      );
    tailStart = at;
  }

  const warnings = new Set<string>();
  const removedIndexes = new Set(blockIndexes.slice(frontEnd, tailStart));
  const removedSections = [...removedIndexes].reduce(
    (count, index) => count + sectionBreaksIn(children[index]),
    0,
  );
  if (removedSections > 0) {
    warnings.add(
      `${String(removedSections)} section break${removedSections === 1 ? "" : "s"} in the replaced body went with it.`,
    );
  }

  const stories = storyParts(pkg);
  if (stories.isErr()) return Result.err(stories.error);
  const builder: BlockBuilder = {
    body,
    w14Prefix: Object.entries(document.value?.attributes ?? {})
      .find(
        ([attribute, value]) => attribute.startsWith("xmlns:") && value === W14,
      )?.[0]
      .slice("xmlns:".length),
    nextParaId: paragraphIds(pkg, stories.value),
    facts,
    styles: styles.value,
    textWidth: textWidthOf(body),
  };
  const inserted = blocks.value.map((block) =>
    blockElements(builder, block, warnings),
  );
  // The content goes before the first block not kept in front, or before
  // the final section properties when every block is kept.
  const anchor =
    frontEnd < blockIndexes.length
      ? (blockIndexes[frontEnd] ?? children.length)
      : (blockIndexes.at(-1) ?? -1) + 1;
  const rebuilt: XmlElement[] = [];
  for (const [index, child] of children.entries()) {
    if (index === anchor) rebuilt.push(...inserted);
    if (!removedIndexes.has(index)) rebuilt.push(child);
  }
  if (anchor >= children.length) {
    const section = rebuilt.findIndex((child) => isW(child, "sectPr"));
    rebuilt.splice(section === -1 ? rebuilt.length : section, 0, ...inserted);
  }
  body.elements = rebuilt;
  pkg.touch(main.value);

  const counts: Record<string, number> = {};
  for (const part of stories.value) {
    const root = pkg.root(part);
    if (root.isErr()) return Result.err(root.error);
    if (root.value === null) continue;
    if (fillText(root.value, fields.value, replacements.value, counts))
      pkg.touch(part);
  }
  for (const key of [...fields.value.keys(), ...replacements.value.keys()]) {
    if (counts[key] === undefined) {
      counts[key] = 0;
      warnings.add(`${JSON.stringify(key)} filled nothing.`);
    }
  }
  return Result.ok({
    kept: { front: frontEnd, tail: blockIndexes.length - tailStart },
    removed: removedIndexes.size,
    inserted: inserted.length,
    fields: counts,
    warnings: [...warnings],
  });
};
