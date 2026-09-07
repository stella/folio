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

/**
 * One inline of a paragraph: plain text, or text carrying an external
 * hyperlink. A link is the case where a revision wrapper and the linked runs
 * have to nest one inside the other, so the fixture has to be able to author
 * one.
 */
export type ParagraphInline = string | { text: string; href: string };

/** One body-level item: a paragraph, or a table given row by row. */
export type BodyItem =
  | { kind: "paragraph"; text: string | readonly ParagraphInline[]; styleId?: string }
  | {
      kind: "table";
      rows: readonly (readonly CellContent[])[];
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
    // `00000000` and `FFFFFFFF` are reserved, and an id is unique per package.
    let candidate = hash % 0xffff_fffe;
    while (taken.has((candidate + 1).toString(16).toUpperCase().padStart(8, "0"))) {
      candidate = (candidate + 1) % 0xffff_fffe;
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

const paragraph = (
  text: string | readonly ParagraphInline[],
  context: BodyContext,
  styleId?: string,
): string => {
  const properties = styleId === undefined ? "" : `<w:pPr><w:pStyle w:val="${styleId}"/></w:pPr>`;
  const inlines = typeof text === "string" ? nonEmptyInlines(text) : text;
  const content = inlines
    .map((inline) => (typeof inline === "string" ? inline : inline.text))
    .join("");
  return (
    `<w:p w14:paraId="${context.paraId(`${styleId ?? ""}|${content}`)}">${properties}` +
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
      for (const cell of row) {
        if (typeof cell !== "string") {
          collectHrefs(cell, hrefs);
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
const tableGrid = (rows: readonly (readonly CellContent[])[]): string => {
  let columns = 0;
  for (const cells of rows) {
    columns = Math.max(columns, cells.length);
  }
  return `<w:tblGrid>${`<w:gridCol w:w="2000"/>`.repeat(columns)}</w:tblGrid>`;
};

const table = (item: Extract<BodyItem, { kind: "table" }>, context: BodyContext): string => {
  const hidden = new Set(item.hiddenRows ?? []);
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
    tableGrid(item.rows) +
    item.rows
      .map(
        (cells, rowIndex) =>
          `<w:tr>${hidden.has(rowIndex) ? `<w:trPr><w:hidden/></w:trPr>` : ""}${cells
            .map(
              (content) =>
                `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${cellXml(content, context)}</w:tc>`,
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
        ? paragraph(item.text, context, item.styleId)
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
  const links = new Map(
    hrefs.map((href, index) => [href, `rId${index + firstLinkRelationship}`]),
  );
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
