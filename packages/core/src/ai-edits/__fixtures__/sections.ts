/**
 * A `.docx` with three sections whose headers and footers use every
 * `ST_HdrFtr` role, generated rather than committed as bytes so a test's input
 * is legible in its own diff.
 *
 * - Section 1 ends at "One closes." (a portrait, `titlePg` section) and
 *   references a default and a first-page header and a default footer.
 * - Section 2 ends at "Two closes." (landscape) and references a default and
 *   an even-page header. It also holds a one-cell table.
 * - Section 3 is the body's final `w:sectPr` and references a default header
 *   and reuses section 1's first-page header part as its own first-page one.
 *
 * `w:evenAndOddHeaders` is on, so every role is live.
 */

import JSZip from "jszip";

import { escapeXmlText } from "@stll/docx-core";

const NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORDPROCESSING = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/** Pinned, with no folder entries, so two builds produce identical bytes. */
const ZIP_ENTRY_OPTIONS = { date: new Date(Date.UTC(2000, 0, 1)), createFolders: false } as const;

const run = (text: string): string =>
  `<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r>`;
const paragraph = (text: string, sectPr = ""): string =>
  `<w:p>${sectPr ? `<w:pPr>${sectPr}</w:pPr>` : ""}${run(text)}</w:p>`;

/** Every header/footer part, by relationship id: part name, role text. */
export const SECTIONS_FIXTURE_PARTS = {
  rIdHeaderOneDefault: { part: "header", file: "header1.xml", text: "Header one default." },
  rIdHeaderOneFirst: { part: "header", file: "header2.xml", text: "Header one first." },
  rIdHeaderTwoDefault: { part: "header", file: "header3.xml", text: "Header two default." },
  rIdHeaderTwoEven: { part: "header", file: "header4.xml", text: "Header two even." },
  rIdHeaderThreeDefault: { part: "header", file: "header5.xml", text: "Header three default." },
  rIdFooterOneDefault: { part: "footer", file: "footer1.xml", text: "Footer one default." },
} as const;

/** The body paragraph texts in document order; the carriers end sections 1 and 2. */
export const SECTIONS_FIXTURE_TEXT = {
  oneOpens: "One opens.",
  oneCloses: "One closes.",
  twoOpens: "Two opens.",
  cell: "Cell text.",
  twoCloses: "Two closes.",
  threeOpens: "Three opens.",
  threeCloses: "Three closes.",
} as const;

const sectionOne =
  `<w:sectPr>` +
  `<w:headerReference w:type="default" r:id="rIdHeaderOneDefault"/>` +
  `<w:headerReference w:type="first" r:id="rIdHeaderOneFirst"/>` +
  `<w:footerReference w:type="default" r:id="rIdFooterOneDefault"/>` +
  `<w:type w:val="nextPage"/>` +
  `<w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:titlePg/>` +
  `</w:sectPr>`;
const sectionTwo =
  `<w:sectPr>` +
  `<w:headerReference w:type="default" r:id="rIdHeaderTwoDefault"/>` +
  `<w:headerReference w:type="even" r:id="rIdHeaderTwoEven"/>` +
  `<w:type w:val="nextPage"/>` +
  `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>` +
  `</w:sectPr>`;
const sectionThree =
  `<w:sectPr>` +
  `<w:headerReference w:type="default" r:id="rIdHeaderThreeDefault"/>` +
  `<w:headerReference w:type="first" r:id="rIdHeaderOneFirst"/>` +
  `<w:type w:val="nextPage"/>` +
  `<w:pgSz w:w="12240" w:h="15840"/>` +
  `<w:titlePg/>` +
  `</w:sectPr>`;

const table = (text: string): string =>
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc></w:tr></w:tbl>`;

export const buildSectionsDocx = async (): Promise<ArrayBuffer> => {
  const text = SECTIONS_FIXTURE_TEXT;
  const body =
    paragraph(text.oneOpens) +
    paragraph(text.oneCloses, sectionOne) +
    paragraph(text.twoOpens) +
    table(text.cell) +
    paragraph(text.twoCloses, sectionTwo) +
    paragraph(text.threeOpens) +
    paragraph(text.threeCloses) +
    sectionThree;

  const parts: Record<string, string> = {
    "[Content_Types].xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="${WORDPROCESSING}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${WORDPROCESSING}.styles+xml"/>` +
      `<Override PartName="/word/settings.xml" ContentType="${WORDPROCESSING}.settings+xml"/>` +
      Object.values(SECTIONS_FIXTURE_PARTS)
        .map(
          ({ part, file }) =>
            `<Override PartName="/word/${file}" ContentType="${WORDPROCESSING}.${part}+xml"/>`,
        )
        .join("") +
      `</Types>`,
    "_rels/.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${RELATIONSHIPS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_RELATIONSHIPS}/officeDocument" Target="word/document.xml"/>` +
      `</Relationships>`,
    "word/_rels/document.xml.rels":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${RELATIONSHIPS}">` +
      `<Relationship Id="rIdStyles" Type="${OFFICE_RELATIONSHIPS}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdSettings" Type="${OFFICE_RELATIONSHIPS}/settings" Target="settings.xml"/>` +
      Object.entries(SECTIONS_FIXTURE_PARTS)
        .map(
          ([id, { part, file }]) =>
            `<Relationship Id="${id}" Type="${OFFICE_RELATIONSHIPS}/${part}" Target="${file}"/>`,
        )
        .join("") +
      `</Relationships>`,
    "word/styles.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles xmlns:w="${NAMESPACE}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `</w:styles>`,
    "word/settings.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:settings xmlns:w="${NAMESPACE}"><w:evenAndOddHeaders/></w:settings>`,
    "word/document.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS}"><w:body>${body}` +
      `</w:body></w:document>`,
  };
  for (const { part, file, text: partText } of Object.values(SECTIONS_FIXTURE_PARTS)) {
    const root = part === "header" ? "w:hdr" : "w:ftr";
    parts[`word/${file}`] =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<${root} xmlns:w="${NAMESPACE}" xmlns:r="${OFFICE_RELATIONSHIPS}">${paragraph(partText)}</${root}>`;
  }

  const zip = new JSZip();
  for (const name of Object.keys(parts).toSorted()) {
    zip.file(name, parts[name] ?? "", ZIP_ENTRY_OPTIONS);
  }
  return await zip.generateAsync({ type: "arraybuffer" });
};
