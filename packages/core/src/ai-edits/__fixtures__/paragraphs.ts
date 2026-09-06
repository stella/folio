/**
 * The smallest `.docx` that carries a list of paragraphs, generated rather
 * than committed as bytes so a test's input is legible in its own diff.
 */

import JSZip from "jszip";

const NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORDPROCESSING = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/** Pinned, with no folder entries, so two builds produce identical bytes. */
const ZIP_ENTRY_OPTIONS = { date: new Date(Date.UTC(2000, 0, 1)), createFolders: false } as const;

const escapeXml = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const paragraphXml = (text: string): string =>
  `<w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;

export type ParagraphsDocxOptions = {
  /**
   * Wrap the paragraphs in a single table cell, followed by one body
   * paragraph. The cell's last paragraph then has a sibling in the document
   * and none in its own container, which is the case a paragraph-mark
   * deletion must refuse.
   */
  inTableCell?: boolean;
};

export const buildParagraphsDocx = async (
  paragraphs: readonly string[],
  { inTableCell }: ParagraphsDocxOptions = {},
): Promise<ArrayBuffer> => {
  const paragraphsXml = paragraphs.map(paragraphXml).join("");
  const body = inTableCell
    ? `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${paragraphsXml}</w:tc></w:tr></w:tbl>` +
      paragraphXml("After the table.")
    : paragraphsXml;

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
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
      `</w:body></w:document>`,
  };

  const zip = new JSZip();
  for (const name of Object.keys(parts).toSorted()) {
    zip.file(name, parts[name] ?? "", ZIP_ENTRY_OPTIONS);
  }
  return await zip.generateAsync({ type: "arraybuffer" });
};
