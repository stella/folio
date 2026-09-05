/**
 * A small multi-level numbered list, authored here rather than committed as
 * bytes.
 *
 * The corpus fixtures carry no numbering at all, so a probe that says "an added
 * list item" while editing an unnumbered paragraph is testing something else
 * and saying it is testing numbering. Generating the package keeps its content
 * legible in the diff and keeps the provenance trivial.
 */

import JSZip from "jszip";

const NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WORDPROCESSING = "application/vnd.openxmlformats-officedocument.wordprocessingml";

/** Pinned so two builds of the fixture produce identical bytes. */
const FIXED_ZIP_DATE = new Date(Date.UTC(2000, 0, 1));

const LIST_LEVELS = 3;

const abstractLevels = (): string =>
  Array.from({ length: LIST_LEVELS }, (_unused, level) => {
    const format = level === 0 ? "decimal" : "lowerLetter";
    return (
      `<w:lvl w:ilvl="${String(level)}"><w:start w:val="1"/>` +
      `<w:numFmt w:val="${format}"/><w:lvlText w:val="%${String(level + 1)}."/>` +
      `<w:lvlJc w:val="left"/></w:lvl>`
    );
  }).join("");

const listParagraph = (level: number, text: string): string =>
  `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr>` +
  `<w:ilvl w:val="${String(level)}"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
  `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

/** One list item: its indent level and its text. */
export type NumberedListItem = { level: number; text: string };

export const NUMBERED_LIST_ITEMS: readonly NumberedListItem[] = Object.freeze([
  { level: 0, text: "The supplier shall deliver the goods to the named place." },
  { level: 1, text: "Delivery is complete on unloading at that place." },
  { level: 1, text: "Risk passes to the buyer on completion of delivery." },
  { level: 0, text: "The buyer shall pay within thirty days of the invoice date." },
  { level: 1, text: "Late payment carries interest at the statutory rate." },
  { level: 0, text: "Either party may terminate for material breach." },
]);

/** The same list with one item demoted one level: a level change and no text change. */
export const withItemDemoted = (
  items: readonly NumberedListItem[],
  index: number,
): NumberedListItem[] =>
  items.map((item, at) =>
    at === index ? { level: Math.min(LIST_LEVELS - 1, item.level + 1), text: item.text } : item,
  );

export const buildNumberedListDocx = async (
  items: readonly NumberedListItem[] = NUMBERED_LIST_ITEMS,
): Promise<ArrayBuffer> => {
  const parts: Record<string, string> = {
    "[Content_Types].xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="${WORDPROCESSING}.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="${WORDPROCESSING}.styles+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="${WORDPROCESSING}.numbering+xml"/>` +
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
      `<Relationship Id="rId2" Type="${OFFICE_RELATIONSHIPS}/numbering" Target="numbering.xml"/>` +
      `</Relationships>`,
    "word/styles.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:styles xmlns:w="${NAMESPACE}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/></w:style>` +
      `</w:styles>`,
    "word/numbering.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:numbering xmlns:w="${NAMESPACE}">` +
      `<w:abstractNum w:abstractNumId="0">${abstractLevels()}</w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
      `</w:numbering>`,
    "word/document.xml":
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${NAMESPACE}"><w:body>` +
      items.map(({ level, text }) => listParagraph(level, text)).join("") +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
      `</w:body></w:document>`,
  };

  const zip = new JSZip();
  for (const name of Object.keys(parts).toSorted()) {
    zip.file(name, parts[name] ?? "", { date: FIXED_ZIP_DATE });
  }
  return await zip.generateAsync({ type: "arraybuffer" });
};
