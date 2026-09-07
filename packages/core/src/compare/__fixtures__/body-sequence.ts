/**
 * A document body written from a sequence of paragraphs and tables.
 *
 * The edit-script DSL can move words and rows but cannot append a table, so a
 * pair whose difference is "a paragraph AND a table were added after the last
 * one" has to be authored as two packages. Generating both from one description
 * keeps the difference visible in the diff instead of hidden in bytes.
 */

import JSZip from "jszip";

const NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORDPROCESSING = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const MARKUP_COMPATIBILITY = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PACKAGE_RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006";
const CORE_PROPERTIES_TYPE = "application/vnd.openxmlformats-package.core-properties+xml";
const WORDML_2010 = "http://schemas.microsoft.com/office/word/2010/wordml";

/** Pinned so two builds of the fixture produce identical bytes. */
const FIXED_ZIP_DATE = new Date(Date.UTC(2000, 0, 1));

/**
 * `createFolders: false` because JSZip stamps the folder entries it
 * synthesizes with `new Date()`, which the fixed date above does not reach.
 */
const ZIP_ENTRY_OPTIONS = { date: FIXED_ZIP_DATE, createFolders: false } as const;

/**
 * One cell's content: a line of text, or a sequence of its own — which is how
 * a nested table, a blank line inside a cell, or a cell that ends with a table
 * gets written.
 */
export type CellContent = string | readonly BodyItem[];

/** `w:tcPr` children a fixture can author, each named after the element. */
export type CellProperties = {
  /** `w:gridSpan`: grid columns this cell occupies. */
  gridSpan?: number;
  /** `w:vMerge`: the cell starts a vertical merge, or continues one. */
  verticalMerge?: "restart" | "continue";
  /** `w:tcW` in twips. */
  width?: number;
  /** `w:shd` fill colour, as six hex digits. */
  shadingFill?: string;
  /** `w:vAlign`. */
  verticalAlign?: "top" | "center" | "bottom";
  /** `w:tcBorders`, single style on all four sides, in eighths of a point. */
  borderSize?: number;
  /** `w:tcMar`, the same margin on all four sides, in twips. */
  margin?: number;
};

/** A cell: its content alone, or its content and its own `w:tcPr`. */
export type Cell = CellContent | ({ content: CellContent } & CellProperties);

/** A row: its cells alone, or its cells and its own `w:trPr`. */
export type TableRow =
  | readonly Cell[]
  | {
      cells: readonly Cell[];
      /** `w:trHeight` in twips. */
      height?: number;
      /** `w:tblHeader`: the row repeats at the top of every page. */
      header?: boolean;
      /** `w:jc` on the row. */
      justification?: "left" | "center" | "right";
    };

/** `w:tblPr` children a fixture can author, each named after the element. */
export type TableProperties = {
  /** `w:tblStyle`. */
  styleId?: string;
  /** `w:tblW`. */
  width?: { value: number; type: "auto" | "dxa" | "pct" };
  /** `w:jc` on the table. */
  justification?: "left" | "center" | "right";
  /** `w:tblInd` in twips. */
  indent?: number;
  /** `w:tblBorders`, single style on every side, in eighths of a point. */
  borderSize?: number;
  /** `w:shd` fill colour on the table, as six hex digits. */
  shadingFill?: string;
  /** `w:tblLayout`. */
  layout?: "fixed" | "autofit";
  /** `w:tblCellMar`, the same margin on all four sides, in twips. */
  cellMargin?: number;
  /** `w:tblLook` value, as four hex digits. */
  look?: string;
};

type RowOptions = Exclude<TableRow, readonly Cell[]>;
type CellSpec = Exclude<Cell, CellContent>;

const isRowOptions = (row: TableRow): row is RowOptions => "cells" in row;
const isCellSpec = (cell: Cell): cell is CellSpec => typeof cell === "object" && "content" in cell;

const rowCells = (row: TableRow): readonly Cell[] => (isRowOptions(row) ? row.cells : row);

const rowOptions = (row: TableRow): Omit<RowOptions, "cells"> => (isRowOptions(row) ? row : {});

const cellContent = (cell: Cell): CellContent => (isCellSpec(cell) ? cell.content : cell);

const cellProperties = (cell: Cell): CellProperties => (isCellSpec(cell) ? cell : {});

/**
 * One inline of a paragraph: plain text, or text carrying an external
 * hyperlink. A link is the case where a revision wrapper and the linked runs
 * have to nest one inside the other, so the fixture has to be able to author
 * one.
 */
export type ParagraphInline = string | { text: string; href: string };

/** One body-level item: a paragraph, or a table given row by row. */
export type BodyItem =
  | {
      kind: "paragraph";
      text: string | readonly ParagraphInline[];
      styleId?: string;
      /**
       * An authored `w14:paraId`, for a package whose ids a producer wrote
       * without respecting the 31-bit bound the schema puts on them.
       */
      paraId?: string;
    }
  | {
      kind: "table";
      rows: readonly TableRow[];
      /** `w:tblPr` children, written in the order the schema requires. */
      properties?: TableProperties;
      /** `w:tblGrid` column widths, in twips. One per grid column. */
      columnWidths?: readonly number[];
      /**
       * Rows a package hides with `w:hidden`. The snapshot skips their whole
       * subtree, so a document that has one is the case where the snapshot
       * walk and the live walk could disagree.
       */
      hiddenRows?: readonly number[];
    };

/**
 * What the body writer carries down the tree: the relationship id each
 * authored href was written under, and the `w14:paraId` allocator.
 *
 * A real package identifies its paragraphs. A fixture that did not would make
 * every save rewrite the whole package instead of splicing the paragraphs that
 * changed, which is a different code path from the one a document exercises.
 */
type BodyContext = {
  links: ReadonlyMap<string, string>;
  paraId: (content: string) => string;
};

/**
 * A paragraph's id is derived from its own content, not from its position.
 *
 * Two packages authored from two descriptions are a base and a target, and a
 * paragraph that appears in both is the same paragraph. Numbering the ids in
 * document order would instead give the same id to the paragraph that happens
 * to sit at the same index, which is how the fixture would tell a comparison
 * that a removed paragraph was a rewrite of the one after it.
 */
const createParaIdAllocator = (): ((content: string) => string) => {
  const taken = new Set<string>();
  return (content) => {
    let hash = 0x811c_9dc5;
    for (let index = 0; index < content.length; index += 1) {
      hash = Math.imul(hash ^ content.charCodeAt(index), 0x0100_0193) >>> 0;
    }
    // A paragraph id is 31-bit, `00000000` is the reserved "no id", and an id
    // is unique per package.
    let candidate = hash % 0x7fff_fffe;
    while (taken.has((candidate + 1).toString(16).toUpperCase().padStart(8, "0"))) {
      candidate = (candidate + 1) % 0x7fff_fffe;
    }
    const paraId = (candidate + 1).toString(16).toUpperCase().padStart(8, "0");
    taken.add(paraId);
    return paraId;
  };
};

const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

const inlineXml = (inline: ParagraphInline, { links }: BodyContext): string =>
  typeof inline === "string"
    ? run(inline)
    : `<w:hyperlink r:id="${links.get(inline.href) ?? ""}">${run(inline.text)}</w:hyperlink>`;

/**
 * An empty paragraph is a `w:p` with no run at all, which is what a package
 * holds for a blank line or an empty cell. It is not the same thing as a
 * paragraph whose run carries an empty string, and both shapes occur.
 */
/** A blank line carries no run at all, so an empty string is no inline. */
const nonEmptyInlines = (text: string): readonly ParagraphInline[] =>
  text.length === 0 ? [] : [text];

type ParagraphOptions = { styleId?: string; paraId?: string };

const paragraph = (
  text: string | readonly ParagraphInline[],
  context: BodyContext,
  { styleId, paraId }: ParagraphOptions = {},
): string => {
  const properties = styleId === undefined ? "" : `<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>`;
  const inlines = typeof text === "string" ? nonEmptyInlines(text) : text;
  const content = inlines
    .map((inline) => (typeof inline === "string" ? inline : inline.text))
    .join("");
  const id = paraId ?? context.paraId(`${styleId ?? ""}|${content}`);
  return (
    `<w:p w14:paraId="${id}" w14:textId="${id}">${properties}` +
    `${inlines.map((inline) => inlineXml(inline, context)).join("")}</w:p>`
  );
};

/** Every href the body carries, in document order, so the ids are stable. */
const collectHrefs = (items: readonly BodyItem[], hrefs: string[]): void => {
  for (const item of items) {
    if (item.kind === "paragraph") {
      if (typeof item.text === "string") {
        continue;
      }
      for (const inline of item.text) {
        if (typeof inline !== "string" && !hrefs.includes(inline.href)) {
          hrefs.push(inline.href);
        }
      }
      continue;
    }
    for (const row of item.rows) {
      for (const cell of rowCells(row)) {
        const content = cellContent(cell);
        if (typeof content !== "string") {
          collectHrefs(content, hrefs);
        }
      }
    }
  }
};

const EMPTY_PARAGRAPH = { kind: "paragraph", text: "" } as const satisfies BodyItem;

/**
 * A container may not end with a table: the format requires a paragraph after
 * one, and a body's section properties do not supply it. One rule for both
 * containers, because a fixture that is well formed in a cell and malformed in
 * the body would be measuring two different things.
 */
const closedSequence = (items: readonly BodyItem[]): readonly BodyItem[] => {
  const last = items.at(-1);
  return last === undefined || last.kind === "table" ? [...items, EMPTY_PARAGRAPH] : items;
};

/** A cell must also contain a paragraph, which the empty sequence supplies. */
const cellXml = (content: CellContent, context: BodyContext): string =>
  typeof content === "string"
    ? paragraph(content, context)
    : itemsXml(closedSequence(content), context);

/** `w:tbl` is `w:tblPr, w:tblGrid, rows`: a fixture without the grid is not one. */
const tableGrid = (item: Extract<BodyItem, { kind: "table" }>): string => {
  if (item.columnWidths) {
    return `<w:tblGrid>${item.columnWidths.map((width) => `<w:gridCol w:w="${String(width)}"/>`).join("")}</w:tblGrid>`;
  }
  let columns = 0;
  for (const row of item.rows) {
    let spanned = 0;
    for (const cell of rowCells(row)) {
      spanned += cellProperties(cell).gridSpan ?? 1;
    }
    columns = Math.max(columns, spanned);
  }
  return `<w:tblGrid>${`<w:gridCol w:w="2000"/>`.repeat(columns)}</w:tblGrid>`;
};

const BORDER_SIDES = ["top", "left", "bottom", "right", "insideH", "insideV"] as const;
const CELL_BORDER_SIDES = ["top", "left", "bottom", "right"] as const;
const MARGIN_SIDES = ["top", "left", "bottom", "right"] as const;

const borders = (element: string, sides: readonly string[], size: number): string =>
  `<w:${element}>${sides
    .map(
      (side) => `<w:${side} w:val="single" w:sz="${String(size)}" w:space="0" w:color="000000"/>`,
    )
    .join("")}</w:${element}>`;

const margins = (element: string, value: number): string =>
  `<w:${element}>${MARGIN_SIDES.map(
    (side) => `<w:${side} w:w="${String(value)}" w:type="dxa"/>`,
  ).join("")}</w:${element}>`;

/**
 * `w:tblPr` children in the order CT_TblPrBase declares: tblStyle, tblW, jc,
 * tblInd, tblBorders, shd, tblLayout, tblCellMar, tblLook. `w:tblW` is always
 * written.
 */
const tableProperties = ({ properties }: Extract<BodyItem, { kind: "table" }>): string => {
  const width = properties?.width ?? { value: 0, type: "auto" };
  const parts = [
    ...(properties?.styleId === undefined ? [] : [`<w:tblStyle w:val="${properties.styleId}"/>`]),
    `<w:tblW w:w="${String(width.value)}" w:type="${width.type}"/>`,
    ...(properties?.justification === undefined
      ? []
      : [`<w:jc w:val="${properties.justification}"/>`]),
    ...(properties?.indent === undefined
      ? []
      : [`<w:tblInd w:w="${String(properties.indent)}" w:type="dxa"/>`]),
    ...(properties?.borderSize === undefined
      ? []
      : [borders("tblBorders", BORDER_SIDES, properties.borderSize)]),
    ...(properties?.shadingFill === undefined
      ? []
      : [`<w:shd w:val="clear" w:color="auto" w:fill="${properties.shadingFill}"/>`]),
    ...(properties?.layout === undefined ? [] : [`<w:tblLayout w:type="${properties.layout}"/>`]),
    ...(properties?.cellMargin === undefined ? [] : [margins("tblCellMar", properties.cellMargin)]),
    ...(properties?.look === undefined ? [] : [`<w:tblLook w:val="${properties.look}"/>`]),
  ];
  return `<w:tblPr>${parts.join("")}</w:tblPr>`;
};

/**
 * `w:trPr` children. `CT_TrPrBase` is a repeated choice rather than a
 * sequence, so the order here is the readable one rather than a required one.
 */
const rowProperties = (row: TableRow, hidden: boolean): string => {
  const { header, height, justification } = rowOptions(row);
  const parts = [
    ...(height === undefined ? [] : [`<w:trHeight w:val="${String(height)}"/>`]),
    ...(header === true ? ["<w:tblHeader/>"] : []),
    ...(justification === undefined ? [] : [`<w:jc w:val="${justification}"/>`]),
    ...(hidden ? ["<w:hidden/>"] : []),
  ];
  return parts.length === 0 ? "" : `<w:trPr>${parts.join("")}</w:trPr>`;
};

/** `w:tcPr` children in schema order; `w:tcW` is always written. */
const cellPropertiesXml = (cell: Cell): string => {
  const properties = cellProperties(cell);
  const parts = [
    `<w:tcW w:w="${String(properties.width ?? 2000)}" w:type="dxa"/>`,
    ...(properties.gridSpan === undefined
      ? []
      : [`<w:gridSpan w:val="${String(properties.gridSpan)}"/>`]),
    ...(properties.verticalMerge === undefined
      ? []
      : [properties.verticalMerge === "restart" ? `<w:vMerge w:val="restart"/>` : `<w:vMerge/>`]),
    ...(properties.borderSize === undefined
      ? []
      : [borders("tcBorders", CELL_BORDER_SIDES, properties.borderSize)]),
    ...(properties.shadingFill === undefined
      ? []
      : [`<w:shd w:val="clear" w:color="auto" w:fill="${properties.shadingFill}"/>`]),
    ...(properties.margin === undefined ? [] : [margins("tcMar", properties.margin)]),
    ...(properties.verticalAlign === undefined
      ? []
      : [`<w:vAlign w:val="${properties.verticalAlign}"/>`]),
  ];
  return `<w:tcPr>${parts.join("")}</w:tcPr>`;
};

const table = (item: Extract<BodyItem, { kind: "table" }>, context: BodyContext): string => {
  const hidden = new Set(item.hiddenRows ?? []);
  return (
    `<w:tbl>${tableProperties(item)}` +
    tableGrid(item) +
    item.rows
      .map(
        (row, rowIndex) =>
          `<w:tr>${rowProperties(row, hidden.has(rowIndex))}${rowCells(row)
            .map(
              (cell) =>
                `<w:tc>${cellPropertiesXml(cell)}${cellXml(cellContent(cell), context)}</w:tc>`,
            )
            .join("")}</w:tr>`,
      )
      .join("") +
    `</w:tbl>`
  );
};

const itemsXml = (items: readonly BodyItem[], context: BodyContext): string =>
  items
    .map((item) =>
      item.kind === "paragraph"
        ? paragraph(item.text, context, {
            ...(item.styleId === undefined ? {} : { styleId: item.styleId }),
            ...(item.paraId === undefined ? {} : { paraId: item.paraId }),
          })
        : table(item, context),
    )
    .join("");

const bodyXml = (items: readonly BodyItem[], context: BodyContext): string =>
  itemsXml(closedSequence(items), context);

/** The relationship id the default header takes when a fixture asks for one. */
const HEADER_RELATIONSHIP_ID = "rId2";

export type BodySequenceOptions = {
  /**
   * A default header part, written as its own sequence. A header is a story of
   * its own: it ends with its own paragraph, and a comparison writes it with
   * its own revision ids.
   */
  header?: readonly BodyItem[];
};

export const buildBodySequenceDocx = async (
  items: readonly BodyItem[],
  { header }: BodySequenceOptions = {},
): Promise<ArrayBuffer> => {
  const hrefs: string[] = [];
  collectHrefs(closedSequence(items), hrefs);
  collectHrefs(closedSequence(header ?? []), hrefs);
  // rId1 is the style part and rId2 the header when there is one; hyperlink
  // relationships follow them in document order.
  const firstLinkRelationship = header ? 3 : 2;
  const links = new Map(hrefs.map((href, index) => [href, `rId${index + firstLinkRelationship}`]));
  const linkRelationships = hrefs
    .map(
      (href, index) =>
        `<Relationship Id="rId${index + firstLinkRelationship}" ` +
        `Type="${OFFICE_RELATIONSHIPS}/hyperlink" ` +
        `Target="${href}" TargetMode="External"/>`,
    )
    .join("");
  const paraId = createParaIdAllocator();

  const parts: Record<string, string> = {
    "[Content_Types].xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="${WORDPROCESSING}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${WORDPROCESSING}.styles+xml"/>` +
      (header
        ? `<Override PartName="/word/header1.xml" ContentType="${WORDPROCESSING}.header+xml"/>`
        : "") +
      `<Override PartName="/docProps/core.xml" ContentType="${CORE_PROPERTIES_TYPE}"/>` +
      `</Types>`,
    "_rels/.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${RELATIONSHIPS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/>` +
      `<Relationship Id="rId2" Type="${PACKAGE_RELATIONSHIPS}/metadata/core-properties" Target="docProps/core.xml"/>` +
      `</Relationships>`,
    // A real package dates itself. Without this part nothing would prove that
    // the comparison stamps `dcterms:modified` from its own timestamp rather
    // than from the second the save happened to run.
    "docProps/core.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<cp:coreProperties xmlns:cp="${PACKAGE_RELATIONSHIPS}/metadata/core-properties" ` +
      `xmlns:dcterms="http://purl.org/dc/terms/" ` +
      `xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
      `<dcterms:created xsi:type="dcterms:W3CDTF">2000-01-01T00:00:00Z</dcterms:created>` +
      `<dcterms:modified xsi:type="dcterms:W3CDTF">2000-01-01T00:00:00Z</dcterms:modified>` +
      `</cp:coreProperties>`,
    "word/_rels/document.xml.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${RELATIONSHIPS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/styles" Target="styles.xml"/>` +
      (header
        ? `<Relationship Id="${HEADER_RELATIONSHIP_ID}" Type="${OFFICE_RELATIONSHIPS}/header" Target="header1.xml"/>`
        : "") +
      linkRelationships +
      `</Relationships>`,
    "word/styles.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles xmlns:w="${NAMESPACE}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>` +
      `</w:styles>`,
    "word/document.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS}" ` +
      `xmlns:mc="${MARKUP_COMPATIBILITY}" xmlns:w14="${WORDML_2010}" mc:Ignorable="w14">` +
      `<w:body>` +
      bodyXml(items, { links, paraId }) +
      `<w:sectPr>` +
      (header ? `<w:headerReference w:type="default" r:id="${HEADER_RELATIONSHIP_ID}"/>` : "") +
      `<w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
      `</w:body></w:document>`,
    ...(header
      ? {
          "word/header1.xml":
            `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<w:hdr xmlns:w="${NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS}" ` +
            `xmlns:mc="${MARKUP_COMPATIBILITY}" xmlns:w14="${WORDML_2010}" mc:Ignorable="w14">` +
            `${itemsXml(closedSequence(header), { links, paraId })}</w:hdr>`,
        }
      : {}),
  };

  const zip = new JSZip();
  for (const name of Object.keys(parts).toSorted()) {
    zip.file(name, parts[name] ?? "", ZIP_ENTRY_OPTIONS);
  }
  return await zip.generateAsync({ type: "arraybuffer" });
};
