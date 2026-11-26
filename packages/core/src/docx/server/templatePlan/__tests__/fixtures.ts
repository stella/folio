/**
 * Synthetic packages for the template plan tests: a house-style template and
 * a source document, written as raw parts so each case controls its markup.
 */
import { ZipArchive } from "@stll/docx-core/zip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS = `xmlns:w="${W}" xmlns:r="${R}"`;
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

export type FixtureParts = Record<string, string | Uint8Array>;

export const buildDocx = async (parts: FixtureParts): Promise<Uint8Array> => {
  const zip = new ZipArchive();
  for (const [path, content] of Object.entries(parts)) {
    zip.set(path, content);
  }
  return await zip.generate();
};

export const readPart = async (
  docx: Uint8Array,
  path: string,
): Promise<string | null> =>
  (await (await ZipArchive.read(docx)).entry(path)?.text()) ?? null;

export const partPaths = async (docx: Uint8Array): Promise<string[]> =>
  (await ZipArchive.read(docx)).paths();

export const replacePart = async (
  docx: Uint8Array,
  path: string,
  content: string,
): Promise<Uint8Array> => {
  const zip = await ZipArchive.read(docx);
  zip.set(path, content);
  return await zip.generate();
};

type ContentTypeOptions = { overrides: Record<string, string> };

const contentTypes = ({ overrides }: ContentTypeOptions): string =>
  `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${Object.entries(
    overrides,
  )
    .map(
      ([part, type]) =>
        `<Override PartName="/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`,
    )
    .join("")}</Types>`;

const PACKAGE_RELS = `${XML}<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`;

type Relationship = {
  id: string;
  type: string;
  target: string;
  external?: boolean;
};

const relationships = (items: readonly Relationship[]): string =>
  `${XML}<Relationships xmlns="${REL}">${items
    .map(
      ({ id, type, target, external }) =>
        `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ""}/>`,
    )
    .join("")}</Relationships>`;

export const p = (inner: string, properties = ""): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}${inner}</w:p>`;

export const r = (text: string, properties = ""): string =>
  `<w:r>${properties === "" ? "" : `<w:rPr>${properties}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;

export const style = (id: string): string => `<w:pStyle w:val="${id}"/>`;

export const numPr = (numId: number, ilvl: number): string =>
  `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;

const fieldRuns = (instruction: string, result: string): string =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const SEPARATOR_NOTES = (kind: "footnote" | "endnote"): string =>
  `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}><w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>`;

const NOTE_STYLES = `<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="18"/></w:rPr></w:style><w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>`;

const DOC_DEFAULTS = `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>`;

export const TEMPLATE_IDS = {
  title: "p0",
  party: "p1",
  pageBreak: "p2",
  tocHeading: "p3",
  tocFirst: "p4",
  tocLast: "p5",
  fillIn: "p6",
  article: "p7",
  clause: "p8",
  signatures: "t9",
  signatureLeft: "p10",
  signatureRight: "p11",
  closing: "p12",
} as const;

const TEMPLATE_STYLES = `${XML}<w:styles ${NS}>${DOC_DEFAULTS}<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="HouseTitle"><w:name w:val="House Title"/><w:basedOn w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/><w:caps/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="HouseArticle"><w:name w:val="House Article"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:numPr><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="HouseClause"><w:name w:val="House Clause"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="HouseSubClause"><w:name w:val="House Sub Clause"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="HouseLetter"><w:name w:val="House Letter"/><w:basedOn w:val="Normal"/><w:pPr><w:numPr><w:ilvl w:val="3"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="HouseBody"><w:name w:val="House Body"/><w:basedOn w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="12pt"/></w:rPr></w:style>${NOTE_STYLES}<w:style w:type="table" w:styleId="HouseTable"><w:name w:val="House Table"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`;

const level = (
  ilvl: number,
  numFmt: string,
  lvlText: string,
  pStyle: string | null,
): string =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${numFmt}"/>${pStyle === null ? "" : `<w:pStyle w:val="${pStyle}"/>`}<w:lvlText w:val="${lvlText}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${(ilvl + 1) * 567}" w:hanging="567"/></w:pPr></w:lvl>`;

const TEMPLATE_NUMBERING = `${XML}<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${level(0, "decimal", "Art. %1", "HouseArticle")}${level(1, "decimal", "%1.%2.", "HouseClause")}${level(2, "decimal", "%1.%2.%3.", "HouseSubClause")}${level(3, "lowerLetter", "%4)", "HouseLetter")}${level(4, "lowerRoman", "(%5)", null)}</w:abstractNum><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>${level(0, "bullet", "-", null)}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`;

const TEMPLATE_SETTINGS = `${XML}<w:settings ${NS}><w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:decimalSymbol w:val=","/><w:listSeparator w:val=";"/></w:settings>`;

const TEMPLATE_HEADER = `${XML}<w:hdr ${NS}>${p(`${r("House letterhead, page ")}${fieldRuns("PAGE", r("1"))}`)}</w:hdr>`;

const TEMPLATE_FOOTER = `${XML}<w:ftr ${NS}>${p(`${r("Printed ")}${fieldRuns('DATE \\@ "d.M.yyyy"', r("1.1.2020"))}${r(" ")}<w:fldSimple w:instr=" FILENAME "><w:r><w:t>template.docx</w:t></w:r></w:fldSimple>`)}</w:ftr>`;

const TOC_ENTRY = (anchor: string, text: string, page: string): string =>
  `<w:hyperlink w:anchor="${anchor}" w:history="1">${r(text)}<w:r><w:tab/></w:r>${fieldRuns(`PAGEREF ${anchor} \\h`, r(page))}</w:hyperlink>`;

const TEMPLATE_BODY = [
  p(r("Agreement"), style("HouseTitle")),
  p(r("[party name]"), style("HouseBody")),
  p('<w:r><w:br w:type="page"/></w:r>'),
  p(r("Contents"), style("HouseTitle")),
  p(
    `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${TOC_ENTRY("_Toc100000001", "Sample article", "2")}`,
    style("TOC1"),
  ),
  p(
    `${TOC_ENTRY("_Toc100000002", "Second sample", "3")}<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
    style("TOC1"),
  ),
  p(
    `${r("To: ")}${fieldRuns('FILLIN "Recipient"', r("Recipient name"))}`,
    style("HouseBody"),
  ),
  p(
    `<w:bookmarkStart w:id="1" w:name="_Toc100000001"/>${r("Sample article")}<w:bookmarkEnd w:id="1"/>`,
    style("HouseArticle"),
  ),
  p(r("Sample clause text."), style("HouseClause")),
  `<w:tbl><w:tblPr><w:tblStyle w:val="HouseTable"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4531"/><w:gridCol w:w="4531"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4531" w:type="dxa"/></w:tcPr>${p(r("........................"))}</w:tc><w:tc><w:tcPr><w:tcW w:w="4531" w:type="dxa"/></w:tcPr>${p(r("........................"))}</w:tc></w:tr></w:tbl>`,
  p(r("Closing words of the template.")),
].join("");

const SECTION = `<w:sectPr><w:headerReference w:type="default" r:id="rId10"/><w:footerReference w:type="default" r:id="rId11"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`;

export type TemplateFixtureOptions = {
  /** Leave the footnotes part out, as a template that never held a note. */
  withoutFootnotes?: boolean;
};

/**
 * A template with styles, one multilevel list bound to them, a header and a
 * footer with fields, a table of contents field over two paragraphs, a FILLIN
 * field and a signature table.
 */
export const buildTemplate = async ({
  withoutFootnotes = false,
}: TemplateFixtureOptions = {}): Promise<Uint8Array> =>
  await buildDocx({
    "[Content_Types].xml": contentTypes({
      overrides: {
        "word/document.xml": "document.main",
        "word/styles.xml": "styles",
        "word/numbering.xml": "numbering",
        "word/settings.xml": "settings",
        "word/header1.xml": "header",
        "word/footer1.xml": "footer",
        ...(withoutFootnotes ? {} : { "word/footnotes.xml": "footnotes" }),
      },
    }),
    "_rels/.rels": PACKAGE_RELS,
    "word/_rels/document.xml.rels": relationships([
      { id: "rId1", type: "styles", target: "styles.xml" },
      { id: "rId2", type: "numbering", target: "numbering.xml" },
      { id: "rId3", type: "settings", target: "settings.xml" },
      ...(withoutFootnotes
        ? []
        : [{ id: "rId4", type: "footnotes", target: "footnotes.xml" }]),
      { id: "rId10", type: "header", target: "header1.xml" },
      { id: "rId11", type: "footer", target: "footer1.xml" },
    ]),
    "word/document.xml": `${XML}<w:document ${NS}><w:body>${TEMPLATE_BODY}${SECTION}</w:body></w:document>`,
    "word/styles.xml": TEMPLATE_STYLES,
    "word/numbering.xml": TEMPLATE_NUMBERING,
    "word/settings.xml": TEMPLATE_SETTINGS,
    "word/header1.xml": TEMPLATE_HEADER,
    "word/footer1.xml": TEMPLATE_FOOTER,
    ...(withoutFootnotes
      ? {}
      : {
          "word/footnotes.xml": `${XML}<w:footnotes ${NS}>${SEPARATOR_NOTES("footnote")}</w:footnotes>`,
        }),
  });

export const SOURCE_IDS = {
  title: "p0",
  blank: "p1",
  articleOne: "p2",
  articleOneTitle: "p3",
  clause: "p4",
  letterA: "p5",
  letterB: "p6",
  articleTwo: "p7",
  articleTwoTitle: "p8",
  manualClause: "p9",
  link: "p10",
  table: "t11",
  cellItem: "p12",
  cellPrice: "p13",
  cellService: "p14",
  cellAmount: "p15",
  mixed: "p16",
  trailingBlank: "p17",
  pageBreak: "p18",
} as const;

export const SOURCE_TEXT = {
  title: "SERVICES AGREEMENT",
  articleOne: "Article I",
  articleOneTitle: "Subject matter",
  clause: "The Provider shall deliver the services described in Article II.",
  letterA: "first item\twith a tab and [to be completed]",
  letterB: "second item￼",
  articleTwo: "Article II",
  articleTwoTitle: "Fees",
  manualClause: "1.\tThe fee follows from Article I of this agreement.",
  link: "See the public register for details.",
  mixed: "Line one\nLine two with H2O and co­operation",
  footnote: " A note with an emphasised word.",
} as const;

const SOURCE_STYLES = `${XML}<w:styles ${NS}>${DOC_DEFAULTS}<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:rPr><w:b/></w:rPr></w:style>${NOTE_STYLES}</w:styles>`;

const SOURCE_NUMBERING = `${XML}<w:numbering ${NS}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>${level(0, "decimal", "%1.", null)}${level(1, "lowerLetter", "%2)", null)}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

const FONT = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>';
const HEADING_DIRECT = `<w:jc w:val="center"/>`;

const SOURCE_BODY = [
  p(r(SOURCE_TEXT.title, `${FONT}<w:b/><w:sz w:val="28"/>`), HEADING_DIRECT),
  p(""),
  p(r(SOURCE_TEXT.articleOne), `${style("Heading1")}${HEADING_DIRECT}`),
  p(r(SOURCE_TEXT.articleOneTitle), `${style("Heading1")}${HEADING_DIRECT}`),
  p(
    `${r("The Provider shall deliver the ", FONT)}<w:commentRangeStart w:id="0"/>${r("services", `${FONT}<w:i/>`)}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>${r(" described in Article II.", '<w:color w:val="FF0000"/>')}`,
    numPr(1, 0),
  ),
  p(
    `${r("first item")}<w:r><w:tab/></w:r>${r("with a tab and ")}${r("[to be completed]", '<w:highlight w:val="yellow"/>')}`,
    numPr(1, 1),
  ),
  p(
    `${r("second item")}<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="2"/></w:r>`,
    numPr(1, 1),
  ),
  p(r(SOURCE_TEXT.articleTwo), `${style("Heading1")}${HEADING_DIRECT}`),
  p(r(SOURCE_TEXT.articleTwoTitle), `${style("Heading1")}${HEADING_DIRECT}`),
  p(
    `${r("1.")}<w:r><w:tab/></w:r>${r("The fee follows from Article I of ")}${r("this agreement", '<w:rStyle w:val="Strong"/>')}${r(".")}`,
  ),
  p(
    `${r("See the ")}<w:hyperlink r:id="rId20" w:history="1">${r("public register", '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>${r(" for details.")}`,
  ),
  `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr>${p(r("Item", "<w:b/>"))}</w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(r("Price", "<w:b/>"))}</w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(r("Service"))}</w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(r("100"))}</w:tc></w:tr></w:tbl>`,
  p(
    `${r("Line one")}<w:r><w:br/></w:r>${r("Line two with H", '<w:u w:val="single"/>')}${r("2", '<w:vertAlign w:val="subscript"/>')}${r("O and co")}<w:r><w:softHyphen/></w:r>${r("operation")}`,
  ),
  p(""),
  p('<w:r><w:br w:type="page"/></w:r>'),
].join("");

/**
 * A source with manual and automatic numbering, a split heading, a footnote,
 * italics, a highlight, a no-break space, a tab, a table, a hyperlink, a
 * comment and two internal references in its text.
 */
export const buildSource = async (): Promise<Uint8Array> =>
  await buildDocx({
    "[Content_Types].xml": contentTypes({
      overrides: {
        "word/document.xml": "document.main",
        "word/styles.xml": "styles",
        "word/numbering.xml": "numbering",
        "word/footnotes.xml": "footnotes",
        "word/comments.xml": "comments",
      },
    }),
    "_rels/.rels": PACKAGE_RELS,
    "word/_rels/document.xml.rels": relationships([
      { id: "rId1", type: "styles", target: "styles.xml" },
      { id: "rId2", type: "numbering", target: "numbering.xml" },
      { id: "rId3", type: "footnotes", target: "footnotes.xml" },
      { id: "rId4", type: "comments", target: "comments.xml" },
      {
        id: "rId20",
        type: "hyperlink",
        target: "https://example.com/register?a=1&amp;b=2",
        external: true,
      },
    ]),
    "word/document.xml": `${XML}<w:document ${NS}><w:body>${SOURCE_BODY}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    "word/styles.xml": SOURCE_STYLES,
    "word/numbering.xml": SOURCE_NUMBERING,
    "word/footnotes.xml": `${XML}<w:footnotes ${NS}>${SEPARATOR_NOTES("footnote")}<w:footnote w:id="2"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>${r(" A note with an ", FONT)}${r("emphasised", "<w:i/>")}${r(" word.")}</w:p></w:footnote></w:footnotes>`,
    "word/comments.xml": `${XML}<w:comments ${NS}><w:comment w:id="0" w:author="Reviewer" w:initials="R"><w:p>${r("Check this.")}</w:p></w:comment></w:comments>`,
  });

/** The body of a minimal document around `body`, for single-purpose cases. */
export const buildMinimalDocx = async (
  body: string,
  extra: FixtureParts = {},
): Promise<Uint8Array> =>
  await buildDocx({
    "[Content_Types].xml": contentTypes({
      overrides: { "word/document.xml": "document.main" },
    }),
    "_rels/.rels": PACKAGE_RELS,
    "word/document.xml": `${XML}<w:document ${NS}><w:body>${body}<w:sectPr/></w:body></w:document>`,
    ...extra,
  });

export { fieldRuns, NS, XML };
