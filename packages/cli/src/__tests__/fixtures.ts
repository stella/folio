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
  /** Runs to write instead of one plain run of `text`; `text` must be their concatenation. */
  runs?: readonly { text: string; bold?: boolean }[];
  /** `w14:paraId`; omit for a paragraph that carries none. */
  paraId?: string;
  /** `w:pStyle`; `Heading1` is defined as an outline level 0 heading. */
  style?: string;
};

const escapeXml = (value: string): string =>
  value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

/** A table: rows of cells, each cell one or more paragraphs. */
export type FixtureTable = { rows: readonly (readonly (readonly FixtureParagraph[])[])[] };

export type FixtureBlock = FixtureParagraph | FixtureTable;

const runXml = ({ text, bold }: { text: string; bold?: boolean }): string =>
  `<w:r>${bold === true ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;

const paragraphXml = ({ text, runs, paraId, style }: FixtureParagraph): string => {
  const id = paraId === undefined ? "" : ` w14:paraId="${paraId}" w14:textId="77777777"`;
  const properties = style === undefined ? "" : `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>`;
  return `<w:p${id}>${properties}${(runs ?? [{ text }]).map(runXml).join("")}</w:p>`;
};

const tableXml = ({ rows }: FixtureTable): string =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>${rows
    .map(
      (cells) =>
        `<w:tr>${cells.map((cell) => `<w:tc>${cell.map(paragraphXml).join("")}</w:tc>`).join("")}</w:tr>`,
    )
    .join("")}</w:tbl>`;

const blockXml = (block: FixtureBlock): string =>
  "rows" in block ? tableXml(block) : paragraphXml(block);

const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:styles xmlns:w="${W_NS}">` +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>' +
  '<w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
  "</w:styles>";

/** Build a minimal WordprocessingML package holding the given paragraphs and tables. */
export const buildDocx = async (paragraphs: readonly FixtureBlock[]): Promise<Uint8Array> => {
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
      `${paragraphs.map(blockXml).join("")}<w:sectPr/></w:body></w:document>`,
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

const CLAUSE_TITLES = [
  "Definitions",
  "Services",
  "Fees and Payment",
  "Term and Termination",
  "Warranties",
  "Confidentiality",
  "Liability",
  "Intellectual Property",
  "Data Protection",
  "General",
];

const SENTENCES = [
  "The Supplier shall perform the Services with reasonable skill and care and in accordance with good industry practice.",
  "The Customer shall provide the Supplier with access to its premises, systems and personnel as reasonably required.",
  "Any change to the scope of the Services shall be agreed in writing by both parties before it takes effect.",
  "The Supplier shall notify the Customer promptly of any matter that may delay or prevent the performance of the Services.",
  "Each party shall comply with all applicable laws and regulations in performing its obligations under this Agreement.",
];

const paraIdOf = (index: number): string => (0x20000000 + index).toString(16).toUpperCase();

/**
 * A contract of about four pages: ten numbered clauses of three paragraphs
 * each, a party table after the first clause, and a bold defined term.
 */
export const fourPageContract = (): FixtureBlock[] => {
  let next = 0;
  const id = (): string => paraIdOf(next++);
  const blocks: FixtureBlock[] = [
    { text: "Services Agreement", paraId: id(), style: "Heading1" },
    {
      text: 'This Agreement is made between Acme Ltd (the "Supplier") and Beta plc (the "Customer").',
      runs: [
        { text: "This Agreement is made between Acme Ltd (the " },
        { text: '"Supplier"', bold: true },
        { text: ') and Beta plc (the "Customer").' },
      ],
      paraId: id(),
    },
  ];
  for (const [clause, title] of CLAUSE_TITLES.entries()) {
    blocks.push({ text: `${clause + 1}. ${title}`, paraId: id(), style: "Heading1" });
    for (let item = 1; item <= 3; item += 1) {
      const body = [0, 1, 2].map(
        (offset) => SENTENCES[(clause + item + offset) % SENTENCES.length],
      );
      const lead =
        clause === 6 && item === 2
          ? "The Supplier's total liability is capped at the fees paid. "
          : "";
      blocks.push({ text: `${clause + 1}.${item} ${lead}${body.join(" ")}`, paraId: id() });
    }
    if (clause === 0) {
      blocks.push({
        rows: [
          [[{ text: "Party", paraId: id() }], [{ text: "Role", paraId: id() }]],
          [[{ text: "Acme Ltd", paraId: id() }], [{ text: "Supplier", paraId: id() }]],
          [[{ text: "Beta plc", paraId: id() }], [{ text: "Customer", paraId: id() }]],
        ],
      });
    }
  }
  return blocks;
};

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
  paragraphs: readonly FixtureBlock[],
): Promise<string> => {
  const filePath = path.join(dir, name);
  await writeFile(filePath, await buildDocx(paragraphs));
  return filePath;
};
