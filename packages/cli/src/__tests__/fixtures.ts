/**
 * Synthetic `.docx` packages for the CLI tests, assembled part by part so a
 * test states exactly which paragraphs, identifiers and styles exist.
 */

import JSZip from "jszip";
import { mkdtemp, realpath, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const STYLES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

export type FixtureParagraph = {
  text: string;
  /** `w14:paraId`; omit for a paragraph that carries none. */
  paraId?: string;
  /** `w:pStyle`; `Heading1` is defined as an outline level 0 heading. */
  style?: string;
};

const escapeXml = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const paragraphXml = ({ text, paraId, style }: FixtureParagraph): string => {
  const id = paraId === undefined ? "" : ` w14:paraId="${paraId}" w14:textId="77777777"`;
  const properties = style === undefined ? "" : `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`;
  return `<w:p${id}>${properties}<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
};

const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:styles xmlns:w="${W_NS}">` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  "</w:styles>";

/** Build a minimal WordprocessingML package holding the given paragraphs. */
export const buildDocx = async (paragraphs: readonly FixtureParagraph[]): Promise<Uint8Array> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      "</Types>",
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${PKG_REL_NS}">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      "</Relationships>",
  );
  zip.file(
    "word/_rels/document.xml.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<Relationships xmlns="${PKG_REL_NS}">` +
      `<Relationship Id="rId1" Type="${STYLES_REL}" Target="styles.xml"/>` +
      "</Relationships>",
  );
  zip.file("word/styles.xml", STYLES_XML);
  zip.file(
    "word/document.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:r="${R_NS}"><w:body>` +
      `${paragraphs.map(paragraphXml).join("")}<w:sectPr/></w:body></w:document>`,
  );
  return await zip.generateAsync({ type: "uint8array" });
};

/** A contract-shaped fixture: a heading and three clauses, all with paraIds. */
export const CONTRACT_PARAGRAPHS: readonly FixtureParagraph[] = [
  { text: "Payment", paraId: "10000001", style: "Heading1" },
  { text: "The buyer pays $50 on signing.", paraId: "10000002" },
  { text: "Late payment accrues interest.", paraId: "10000003" },
  { text: "Either party may terminate on notice.", paraId: "10000004" },
];

/** A temporary directory removed by the returned `cleanup`. */
export const makeTempDir = async (): Promise<{ dir: string; cleanup: () => Promise<void> }> => {
  // Real path: the CLI reports resolved paths, and the system temp
  // directory is behind a symlink on some platforms.
  const dir = await realpath(await mkdtemp(path.join(tmpdir(), "folio-cli-")));
  return { dir, cleanup: () => rm(dir, { recursive: true, force: true }) };
};

/** Write a fixture into `dir` and return its path. */
export const writeDocx = async (
  dir: string,
  name: string,
  paragraphs: readonly FixtureParagraph[],
): Promise<string> => {
  const filePath = path.join(dir, name);
  await writeFile(filePath, await buildDocx(paragraphs));
  return filePath;
};
