/**
 * A short NDA-shaped document for the side-panel layout specs: four headings
 * over three pages (so the outline has a rail and a column to fill), and,
 * when asked, a tracked insertion and deletion, with an optional comment.
 *
 * Built from hand-written OOXML at test time rather than committed as a
 * binary, so what the document holds stays reviewable.
 */

import JSZip from "jszip";

export type PanelLayoutReview = "none" | "changes-only" | "comment-and-changes";
export type PanelLayoutSections = "portrait" | "landscape-then-portrait";

/** Fixed so the bytes are the same on every run. */
const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");
const REVIEW_DATE = "2026-01-01T00:00:00Z";

export const PANEL_LAYOUT_HEADINGS = [
  { text: "1. Definitions", level: 1 },
  { text: "2. Obligations", level: 1 },
  { text: "2.1 Permitted disclosures", level: 2 },
  { text: "3. Term and termination", level: 1 },
] as const;

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const CLAUSE =
  "The Receiving Party shall hold the Confidential Information in strict confidence, use it solely to evaluate the proposed transaction, and restrict access to those of its employees and advisers who need to know it for that purpose and are bound by obligations of confidentiality no less protective than these.";

const escapeXml = (text: string) =>
  text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const run = (text: string) => `<w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;

const paragraph = (content: string) => `<w:p>${content}</w:p>`;

const heading = (text: string, level: number) =>
  `<w:p><w:pPr><w:pStyle w:val="Heading${String(level)}"/></w:pPr>${run(text)}</w:p>`;

const clauses = (count: number) =>
  Array.from({ length: count }, () => paragraph(run(CLAUSE))).join("");

const reviewedParagraph = (review: PanelLayoutReview) => {
  if (review === "none") {
    return paragraph(run("Each party discloses information to the other under this Agreement."));
  }
  return paragraph(
    [
      ...(review === "comment-and-changes" ? ['<w:commentRangeStart w:id="1"/>'] : []),
      run("Each party discloses information to the other"),
      ...(review === "comment-and-changes"
        ? ['<w:commentRangeEnd w:id="1"/>', '<w:r><w:commentReference w:id="1"/></w:r>']
        : []),
      `<w:ins w:id="10" w:author="Counsel" w:date="${REVIEW_DATE}">${run(" in writing")}</w:ins>`,
      run(" under this Agreement"),
      `<w:del w:id="11" w:author="Counsel" w:date="${REVIEW_DATE}"><w:r><w:delText xml:space="preserve"> and its schedules</w:delText></w:r></w:del>`,
      run("."),
    ].join(""),
  );
};

const pageSection = (width: number, height: number) =>
  `<w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="${String(width)}" w:h="${String(height)}"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;

const documentXml = (review: PanelLayoutReview, sections: PanelLayoutSections) => {
  const [definitions, obligations, disclosures, term] = PANEL_LAYOUT_HEADINGS;
  const body = [
    paragraph(run("Mutual Non-Disclosure Agreement")),
    reviewedParagraph(review),
    heading(definitions.text, definitions.level),
    clauses(6),
    ...(sections === "landscape-then-portrait"
      ? [`<w:p><w:pPr>${pageSection(15840, 12240)}</w:pPr>${run("Section break")}</w:p>`]
      : []),
    heading(obligations.text, obligations.level),
    clauses(6),
    heading(disclosures.text, disclosures.level),
    clauses(6),
    heading(term.text, term.level),
    clauses(6),
  ].join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`;
};

const headingStyle = (level: number) =>
  `<w:style w:type="paragraph" w:styleId="Heading${String(level)}"><w:name w:val="heading ${String(level)}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="${String(level - 1)}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${String(32 - level * 4)}"/></w:rPr></w:style>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>${headingStyle(1)}${headingStyle(2)}</w:styles>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${W_NS}><w:comment w:id="1" w:author="Counsel" w:initials="C" w:date="${REVIEW_DATE}"><w:p>${run("Should this cover oral disclosures too?")}</w:p></w:comment></w:comments>`;

const contentTypes = (
  review: PanelLayoutReview,
) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${
    review !== "comment-and-changes"
      ? ""
      : `
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>`
  }
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const documentRels = (
  review: PanelLayoutReview,
) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>${
    review !== "comment-and-changes"
      ? ""
      : `
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>`
  }
</Relationships>`;

/** The document's `.docx` bytes. */
export const buildPanelLayoutDocument = (
  review: PanelLayoutReview,
  sections: PanelLayoutSections = "portrait",
): Promise<Uint8Array> => {
  const zip = new JSZip();
  const options = { date: FIXED_DATE };
  zip.file("[Content_Types].xml", contentTypes(review), options);
  zip.file("_rels/.rels", PACKAGE_RELS, options);
  zip.file("word/document.xml", documentXml(review, sections), options);
  zip.file("word/_rels/document.xml.rels", documentRels(review), options);
  zip.file("word/styles.xml", STYLES_XML, options);
  if (review === "comment-and-changes") {
    zip.file("word/comments.xml", COMMENTS_XML, options);
  }
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
};
