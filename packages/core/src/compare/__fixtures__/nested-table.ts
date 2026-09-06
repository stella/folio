/**
 * A document that ends with a table whose last cell holds another table.
 *
 * The shape matters because the story's last block is then two levels deep,
 * and a paragraph appended after the whole table has no block to anchor to
 * that is not inside a cell. Authored here rather than committed as bytes so
 * the structure under test is legible in the diff.
 */

import JSZip from "jszip";

const NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORDPROCESSING = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/** Pinned so two builds of the fixture produce identical bytes. */
const FIXED_ZIP_DATE = new Date(Date.UTC(2000, 0, 1));

/**
 * `createFolders: false` because JSZip stamps the folder entries it
 * synthesizes with `new Date()`, which the fixed date above does not reach.
 */
const ZIP_ENTRY_OPTIONS = { date: FIXED_ZIP_DATE, createFolders: false } as const;

const paragraph = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const cell = (inner: string): string =>
  `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr>${inner}</w:tc>`;

const table = (rows: readonly string[]): string =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
  `<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>` +
  `${rows.join("")}</w:tbl>`;

/** The nested table sits in the outer table's last cell, so it closes the story. */
const NESTED_TABLE = table([
  `<w:tr>${cell(paragraph("The nested schedule lists the delivery dates."))}</w:tr>`,
]);

const OUTER_TABLE = table([
  `<w:tr>${cell(paragraph("Obligations of the supplier."))}${cell(paragraph("Obligations of the buyer."))}</w:tr>`,
  `<w:tr>${cell(paragraph("Deliver the goods to the named place."))}${cell(`${paragraph("See the schedule below.")}${NESTED_TABLE}`)}</w:tr>`,
]);

export type NestedTableDocxOptions = {
  /** Appended after the outer table, at body level, when set. */
  trailingParagraph?: string;
};

export const buildNestedTableDocx = async ({
  trailingParagraph,
}: NestedTableDocxOptions = {}): Promise<ArrayBuffer> => {
  const body =
    paragraph("This agreement is made between the parties named below.") +
    OUTER_TABLE +
    (trailingParagraph === undefined ? "" : paragraph(trailingParagraph));

  const parts: Record<string, string> = {
    "[Content_Types].xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="${WORDPROCESSING}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${WORDPROCESSING}.styles+xml"/>` +
      `</Types>`,
    "_rels/.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${RELATIONSHIPS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`,
    "word/_rels/document.xml.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${RELATIONSHIPS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/styles" Target="styles.xml"/>` +
      `</Relationships>`,
    "word/styles.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles xmlns:w="${NAMESPACE}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `</w:styles>`,
    "word/document.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${NAMESPACE}"><w:body>${body}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
      `</w:body></w:document>`,
  };

  const zip = new JSZip();
  for (const name of Object.keys(parts).toSorted()) {
    zip.file(name, parts[name] ?? "", ZIP_ENTRY_OPTIONS);
  }
  return await zip.generateAsync({ type: "arraybuffer" });
};
