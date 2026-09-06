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

/** Pinned so two builds of the fixture produce identical bytes. */
const FIXED_ZIP_DATE = new Date(Date.UTC(2000, 0, 1));

/**
 * `createFolders: false` because JSZip stamps the folder entries it
 * synthesizes with `new Date()`, which the fixed date above does not reach.
 */
const ZIP_ENTRY_OPTIONS = { date: FIXED_ZIP_DATE, createFolders: false } as const;

/** One body-level item: a paragraph, or a table given row by row. */
export type BodyItem =
  | { kind: "paragraph"; text: string }
  | { kind: "table"; rows: readonly (readonly string[])[] };

const paragraph = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const table = (rows: readonly (readonly string[])[]): string =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
  rows
    .map(
      (cells) =>
        `<w:tr>${cells
          .map(
            (text) =>
              `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`,
          )
          .join("")}</w:tr>`,
    )
    .join("") +
  `</w:tbl>`;

const bodyXml = (items: readonly BodyItem[]): string =>
  items
    .map((item) => (item.kind === "paragraph" ? paragraph(item.text) : table(item.rows)))
    // A table may not be the body's last element: it needs a paragraph after
    // it, which the section properties do not supply.
    .join("");

export const buildBodySequenceDocx = async (items: readonly BodyItem[]): Promise<ArrayBuffer> => {
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
      `<w:document xmlns:w="${NAMESPACE}"><w:body>` +
      bodyXml(items) +
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
