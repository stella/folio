#!/usr/bin/env bun
/** Build deterministic, synthetic DOCX fixtures for broad layout parity work. */
import { mkdir } from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  buildLayoutInteractionMatrix,
  LAYOUT_INTERACTION_AXES,
  type LayoutInteractionCase,
} from "./layout-interaction-matrix";

const FIXED_DATE = new Date("2026-01-01T00:00:00.000Z");
const DEFAULT_OUTPUT_DIR = import.meta.dir;
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

const REL = {
  officeDocument:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  styles: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  settings: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
  numbering: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
  header: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
  footer: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
  image: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
  footnotes: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
} as const;

const CONTENT_TYPE = {
  document: "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
  styles: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  settings: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
  numbering: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  header: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
  footer: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
  footnotes: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
} as const;

const ROOT_NAMESPACES = [
  `xmlns:w="${W_NS}"`,
  `xmlns:r="${R_NS}"`,
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(" ");

const FIXTURE_NAMES = [
  "isolated-page-furniture.docx",
  "pairwise-layout-interactions.docx",
  "layout-kitchen-sink.docx",
] as const;
const MATRIX_MANIFEST_NAME = "layout-interaction-matrix.json";

type FixtureName = (typeof FIXTURE_NAMES)[number];

type Fixture = {
  name: FixtureName;
  body: string;
  sectionProperties: string;
  parts?: Record<string, string | Uint8Array>;
  relationships?: string;
  overrides?: string;
};

const pixelPng = Uint8Array.from(
  Buffer.from(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
    "base64",
  ),
);

const contentTypes = (
  overrides = "",
): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="${CONTENT_TYPE.document}"/>
  <Override PartName="/word/styles.xml" ContentType="${CONTENT_TYPE.styles}"/>
  <Override PartName="/word/settings.xml" ContentType="${CONTENT_TYPE.settings}"/>
  <Override PartName="/word/numbering.xml" ContentType="${CONTENT_TYPE.numbering}"/>
  ${overrides}
</Types>`;

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="${REL.officeDocument}" Target="word/document.xml"/>
</Relationships>`;

const documentRelationships = (
  extra = "",
): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${REL_NS}">
  <Relationship Id="rId1" Type="${REL.styles}" Target="styles.xml"/>
  <Relationship Id="rId2" Type="${REL.settings}" Target="settings.xml"/>
  <Relationship Id="rId3" Type="${REL.numbering}" Target="numbering.xml"/>
  ${extra}
</Relationships>`;

const SETTINGS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="${W_NS}">
  <w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>
  <w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>
</w:settings>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="30"/></w:rPr></w:style>
  <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/><w:color w:val="365F91"/></w:rPr></w:style>
</w:styles>`;

const NUMBERING_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="1440"/></w:tabs><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const pageProperties = (extra = ""): string => `
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="360" w:footer="360" w:gutter="0"/>
  ${extra}`;

const headerFooterReferences = `
  <w:headerReference w:type="default" r:id="rId10"/>
  <w:footerReference w:type="default" r:id="rId11"/>`;

type AnchoredImageOptions = {
  id: number;
  relationshipId: string;
  width: number;
  height: number;
  horizontalOffset: number;
  verticalOffset: number;
  horizontalRelativeTo?: "column" | "margin" | "page";
  verticalRelativeTo: "line" | "margin" | "paragraph" | "page";
  wrap: "square" | "topAndBottom" | "none";
};

const anchoredImage = ({
  id,
  relationshipId,
  width,
  height,
  horizontalOffset,
  verticalOffset,
  horizontalRelativeTo = "column",
  verticalRelativeTo,
  wrap,
}: AnchoredImageOptions): string => {
  let wrapXml = "<wp:wrapNone/>";
  if (wrap === "square") {
    wrapXml = '<wp:wrapSquare wrapText="bothSides"/>';
  } else if (wrap === "topAndBottom") {
    wrapXml = "<wp:wrapTopAndBottom/>";
  }
  return `<w:drawing><wp:anchor distT="0" distB="95250" distL="0" distR="0" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="${horizontalRelativeTo}"><wp:posOffset>${horizontalOffset}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="${verticalRelativeTo}"><wp:posOffset>${verticalOffset}</wp:posOffset></wp:positionV><wp:extent cx="${width}" cy="${height}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>${wrapXml}<wp:docPr id="${id}" name="Synthetic band ${id}" descr="Synthetic layout marker"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="synthetic-band.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
};

const inlineImage = (id: number): string =>
  `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="457200" cy="228600"/><wp:docPr id="${id}" name="Synthetic inline ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="synthetic-band.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId13"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="457200" cy="228600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const matrixAnchor = (scenario: LayoutInteractionCase, index: number): string => {
  if (scenario.anchorFrame === "inline") {
    return `<w:p><w:r>${inlineImage(1000 + index)}</w:r><w:r><w:t xml:space="preserve"> inline marker</w:t></w:r></w:p>`;
  }
  const frame = {
    page: { horizontalRelativeTo: "page", verticalRelativeTo: "page" },
    margin: { horizontalRelativeTo: "margin", verticalRelativeTo: "margin" },
    column: { horizontalRelativeTo: "column", verticalRelativeTo: "line" },
    paragraph: { horizontalRelativeTo: "column", verticalRelativeTo: "paragraph" },
  } as const;
  const relative = frame[scenario.anchorFrame];
  if (scenario.wrap === "inline") {
    throw new TypeError("a floating matrix anchor cannot use inline wrapping");
  }
  const wrap = scenario.wrap === "topBottom" ? "topAndBottom" : scenario.wrap;
  return `<w:p><w:r>${anchoredImage({ id: 1000 + index, relationshipId: "rId13", width: 457200, height: 228600, horizontalOffset: 182880, verticalOffset: 91440, horizontalRelativeTo: relative.horizontalRelativeTo, verticalRelativeTo: relative.verticalRelativeTo, wrap })}</w:r><w:r><w:t>Anchored marker</w:t></w:r></w:p>`;
};

const matrixFlow = ({ flow, id }: LayoutInteractionCase): string => {
  const properties = {
    normal: "",
    keepNext: "<w:keepNext/>",
    keepLines: "<w:keepLines/>",
    hardPageBreak: "",
    renderedPageBreak: "",
  }[flow];
  const prefix = {
    normal: "",
    keepNext: "",
    keepLines: "",
    hardPageBreak: '<w:br w:type="page"/>',
    renderedPageBreak: "<w:lastRenderedPageBreak/>",
  }[flow];
  return `<w:p><w:pPr>${properties}</w:pPr><w:r>${prefix}<w:t>${id} flow control with invented text.</w:t></w:r></w:p>`;
};

const matrixTypography = ({ typography, id }: LayoutInteractionCase): string => {
  switch (typography) {
    case "latin":
      return `<w:p><w:r><w:t>${id} Alpha cedar 731 uses synthetic prose.</w:t></w:r></w:p>`;
    case "rtl":
      return `<w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rtl/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>اختبار تخطيط بقيم وهمية فقط</w:t></w:r></w:p>`;
    case "cjk":
      return `<w:p><w:r><w:rPr><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>架空の値だけを使うレイアウト試験です。</w:t></w:r></w:p>`;
    case "tabs":
      return `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2160"/><w:tab w:val="decimal" w:pos="7200"/></w:tabs></w:pPr><w:r><w:t>Item ${id}</w:t><w:tab/><w:t>Fictional service</w:t><w:tab/><w:t>7,531.42</w:t></w:r></w:p>`;
    case "numbering":
      return `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>${id} numbered synthetic clause.</w:t></w:r></w:p>`;
  }
  typography satisfies never;
};

const matrixTable = ({ table, id }: LayoutInteractionCase): string => {
  if (table === "none") return "";
  const layout = table === "autofit" ? "autofit" : "fixed";
  const borders = `<w:tblBorders><w:top w:val="single" w:sz="8" w:color="365F91"/><w:left w:val="single" w:sz="8" w:color="365F91"/><w:bottom w:val="single" w:sz="8" w:color="365F91"/><w:right w:val="single" w:sz="8" w:color="365F91"/><w:insideH w:val="single" w:sz="4" w:color="A6A6A6"/><w:insideV w:val="single" w:sz="4" w:color="A6A6A6"/></w:tblBorders>`;
  const cell = (text: string, properties = ""): string =>
    `<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/>${properties}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  let rows = `<w:tr>${cell(`${id} A`)}${cell("Synthetic 842")}</w:tr>`;
  if (table === "merged") {
    rows = `<w:tr>${cell(`${id} merged`, '<w:vMerge w:val="restart"/>')}${cell("Alpha")}</w:tr><w:tr>${cell("", "<w:vMerge/>")}${cell("Beta")}</w:tr>`;
  } else if (table === "splitRow") {
    const paragraphs = Array.from(
      { length: 58 },
      (_, line) => `<w:p><w:r><w:t>${id} split-row line ${line + 1}.</w:t></w:r></w:p>`,
    ).join("");
    rows = `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>${paragraphs}</w:tc></w:tr>`;
  }
  return `<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="${layout}"/>${borders}</w:tblPr><w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid>${rows}</w:tbl>`;
};

const matrixSectionBoundary = ({ section }: LayoutInteractionCase): string => {
  if (section === "single") return "";
  const type = section === "nextPage" ? "nextPage" : "continuous";
  const columns = section === "twoColumn" ? '<w:cols w:num="2" w:space="720"/>' : "";
  return `<w:p><w:pPr><w:sectPr><w:type w:val="${type}"/>${pageProperties(columns)}</w:sectPr></w:pPr><w:r><w:t>Section boundary</w:t></w:r></w:p>`;
};

const matrixCaseContent = (scenario: LayoutInteractionCase, index: number): string => {
  const label = `${scenario.id}: ${scenario.section}, ${scenario.anchorFrame}, ${scenario.wrap}, ${scenario.flow}, ${scenario.table}, ${scenario.typography}`;
  return `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>${label}</w:t></w:r></w:p>${matrixAnchor(scenario, index)}${matrixFlow(scenario)}${matrixTypography(scenario)}${matrixTable(scenario)}`;
};

type MatrixCaseBodyOptions = {
  scenario: LayoutInteractionCase;
  index: number;
  previousScenario: LayoutInteractionCase | undefined;
};

const matrixCaseBody = ({ scenario, index, previousScenario }: MatrixCaseBodyOptions): string => {
  const pageStart =
    index === 0 || previousScenario?.section === "nextPage"
      ? ""
      : '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
  return `${pageStart}${matrixCaseContent(scenario, index)}${matrixSectionBoundary(scenario)}`;
};

const pairwiseMatrixBody = (): string => {
  const scenarios = buildLayoutInteractionMatrix();
  return scenarios
    .map((scenario, index) =>
      matrixCaseBody({ scenario, index, previousScenario: scenarios.at(index - 1) }),
    )
    .join("");
};

const indentJson = (value: unknown, spaces: number): string =>
  JSON.stringify(value, null, 2).replaceAll("\n", `\n${" ".repeat(spaces)}`);

const compactJsonArray = (values: readonly string[]): string =>
  `[${values.map((value) => JSON.stringify(value)).join(", ")}]`;

export const layoutInteractionMatrixManifest = (): string => `{
  "version": 1,
  "strength": 2,
  "axes": {
    "section": ${compactJsonArray(LAYOUT_INTERACTION_AXES.section)},
    "anchorFrame": ${compactJsonArray(LAYOUT_INTERACTION_AXES.anchorFrame)},
    "wrap": ${compactJsonArray(LAYOUT_INTERACTION_AXES.wrap)},
    "flow": ${compactJsonArray(LAYOUT_INTERACTION_AXES.flow)},
    "table": ${compactJsonArray(LAYOUT_INTERACTION_AXES.table)},
    "typography": ${compactJsonArray(LAYOUT_INTERACTION_AXES.typography)}
  },
  "cases": ${indentJson(buildLayoutInteractionMatrix(), 2)}
}\n`;

const pageTextBox = (text: string): string =>
  `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="40" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>2286000</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="40" name="Synthetic page overlay"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="BF9000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;

const HEADER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${ROOT_NAMESPACES}><w:p><w:r>${anchoredImage({ id: 20, relationshipId: "rId1", width: 5943600, height: 381000, horizontalOffset: 0, verticalOffset: 95250, verticalRelativeTo: "paragraph", wrap: "square" })}</w:r></w:p><w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="1F4E79"/></w:rPr><w:t>SYNTHETIC HEADER • ALPHA 104</w:t></w:r></w:p></w:hdr>`;

const FOOTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${ROOT_NAMESPACES}><w:p><w:r>${anchoredImage({ id: 21, relationshipId: "rId1", width: 5943600, height: 190500, horizontalOffset: 0, verticalOffset: 95250, verticalRelativeTo: "paragraph", wrap: "square" })}</w:r></w:p><w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr><w:hyperlink r:id="rId2"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>https://example.invalid/layout-204</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> • page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`;

const HEADER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${REL.image}" Target="media/synthetic-band.png"/></Relationships>`;
const FOOTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${REL.image}" Target="media/synthetic-band.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/layout-204" TargetMode="External"/></Relationships>`;

const PAGE_FURNITURE_PARTS = {
  "word/header1.xml": HEADER_XML,
  "word/footer1.xml": FOOTER_XML,
  "word/_rels/header1.xml.rels": HEADER_RELS,
  "word/_rels/footer1.xml.rels": FOOTER_RELS,
  "word/media/synthetic-band.png": pixelPng,
};

const PAGE_FURNITURE_RELATIONSHIPS = `<Relationship Id="rId10" Type="${REL.header}" Target="header1.xml"/><Relationship Id="rId11" Type="${REL.footer}" Target="footer1.xml"/>`;
const MATRIX_RELATIONSHIPS = `${PAGE_FURNITURE_RELATIONSHIPS}<Relationship Id="rId13" Type="${REL.image}" Target="media/synthetic-band.png"/>`;
const PAGE_FURNITURE_OVERRIDES = `<Override PartName="/word/header1.xml" ContentType="${CONTENT_TYPE.header}"/><Override PartName="/word/footer1.xml" ContentType="${CONTENT_TYPE.footer}"/>`;

const bodyParagraphs = (count: number, prefix: string): string =>
  Array.from(
    { length: count },
    (_, index) =>
      `<w:p><w:r><w:t>${prefix} ${index + 1}: Cedar 17 records a fictional obligation, notice period, and delivery checkpoint for layout testing.</w:t></w:r></w:p>`,
  ).join("");

const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="${W_NS}"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="2"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Synthetic note Delta 308 uses invented values only.</w:t></w:r></w:p></w:footnote></w:footnotes>`;

const fixtures = (): Fixture[] => [
  {
    name: "isolated-page-furniture.docx",
    body: `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Isolated page furniture</w:t></w:r></w:p>${bodyParagraphs(5, "Alpha")}`,
    sectionProperties: `${headerFooterReferences}${pageProperties()}`,
    parts: PAGE_FURNITURE_PARTS,
    relationships: PAGE_FURNITURE_RELATIONSHIPS,
    overrides: PAGE_FURNITURE_OVERRIDES,
  },
  {
    name: "pairwise-layout-interactions.docx",
    body: pairwiseMatrixBody(),
    sectionProperties: `${headerFooterReferences}${pageProperties()}`,
    parts: PAGE_FURNITURE_PARTS,
    relationships: MATRIX_RELATIONSHIPS,
    overrides: PAGE_FURNITURE_OVERRIDES,
  },
  {
    name: "layout-kitchen-sink.docx",
    body: `${pageTextBox("Overlay • Kappa 511")}<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Layout kitchen sink</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Mixed runs: </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t xml:space="preserve">, </w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="C00000"/></w:rPr><w:t>italic color</w:t></w:r><w:r><w:t>, and ordinary text.</w:t></w:r></w:p><w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="2160"/><w:tab w:val="decimal" w:pos="7200"/></w:tabs></w:pPr><w:r><w:t>Item Lambda</w:t><w:tab/><w:t>Fictional service</w:t><w:tab/><w:t>1,234.56</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Numbered clause with enough words to wrap and expose hanging-indent behavior near the right margin.</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Nested clause Mu 622.</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders><w:top w:val="single" w:sz="8" w:color="1F4E79"/><w:left w:val="single" w:sz="8" w:color="1F4E79"/><w:bottom w:val="single" w:sz="8" w:color="1F4E79"/><w:right w:val="single" w:sz="8" w:color="1F4E79"/><w:insideH w:val="single" w:sz="4" w:color="A6A6A6"/><w:insideV w:val="single" w:sz="4" w:color="A6A6A6"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2520"/><w:gridCol w:w="4320"/><w:gridCol w:w="2520"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2520" w:type="dxa"/><w:shd w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:b/><w:t>Code</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:shd w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:b/><w:t>Description</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2520" w:type="dxa"/><w:shd w:fill="D9EAF7"/></w:tcPr><w:p><w:r><w:b/><w:t>Amount</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="2520" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>NU-734</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Merged description spanning two rows</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2520" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>8,765.43</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="2520" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>XI-845</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcW w:w="2520" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>92.10</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p><w:r><w:t xml:space="preserve">Footnote anchor</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="2"/></w:r><w:r><w:t>.</w:t></w:r></w:p><w:p><w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:rtl/><w:lang w:bidi="ar-SA"/></w:rPr><w:t>اختبار تخطيط بقيم وهمية فقط</w:t></w:r></w:p><w:p><w:r><w:br w:type="page"/></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Second page and columns</w:t></w:r></w:p>${bodyParagraphs(8, "Omega")}<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/>${headerFooterReferences}${pageProperties('<w:cols w:num="2" w:space="720"/>')}</w:sectPr></w:pPr><w:r><w:t>Continuous two-column boundary</w:t></w:r></w:p>${bodyParagraphs(12, "Sigma")}`,
    sectionProperties: `${headerFooterReferences}${pageProperties('<w:pgBorders w:offsetFrom="page"><w:top w:val="single" w:sz="12" w:space="24" w:color="1F4E79"/><w:left w:val="single" w:sz="12" w:space="24" w:color="1F4E79"/><w:bottom w:val="single" w:sz="12" w:space="24" w:color="1F4E79"/><w:right w:val="single" w:sz="12" w:space="24" w:color="1F4E79"/></w:pgBorders>')}`,
    parts: { ...PAGE_FURNITURE_PARTS, "word/footnotes.xml": footnotesXml },
    relationships: `${PAGE_FURNITURE_RELATIONSHIPS}<Relationship Id="rId12" Type="${REL.footnotes}" Target="footnotes.xml"/>`,
    overrides: `${PAGE_FURNITURE_OVERRIDES}<Override PartName="/word/footnotes.xml" ContentType="${CONTENT_TYPE.footnotes}"/>`,
  },
];

const addPart = (zip: JSZip, filePath: string, contents: string | Uint8Array): void => {
  zip.file(filePath, contents, { createFolders: false, date: FIXED_DATE });
};

const buildFixture = ({
  body,
  overrides,
  parts,
  relationships,
  sectionProperties,
}: Fixture): Promise<Uint8Array> => {
  const zip = new JSZip();
  addPart(zip, "[Content_Types].xml", contentTypes(overrides));
  addPart(zip, "_rels/.rels", PACKAGE_RELS);
  addPart(zip, "word/_rels/document.xml.rels", documentRelationships(relationships));
  addPart(
    zip,
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${ROOT_NAMESPACES}><w:body>${body}<w:sectPr>${sectionProperties}</w:sectPr></w:body></w:document>`,
  );
  addPart(zip, "word/styles.xml", STYLES_XML);
  addPart(zip, "word/settings.xml", SETTINGS_XML);
  addPart(zip, "word/numbering.xml", NUMBERING_XML);
  for (const [partPath, contents] of Object.entries(parts ?? {})) {
    addPart(zip, partPath, contents);
  }
  return zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
};

export const buildLayoutInteractionCaseFixture = (
  scenario: LayoutInteractionCase,
): Promise<Uint8Array> =>
  buildFixture({
    name: "pairwise-layout-interactions.docx",
    body: `${matrixCaseContent(scenario, 0)}${matrixSectionBoundary(scenario)}<w:p><w:r><w:t>${scenario.id} post-boundary sentinel.</w:t></w:r></w:p>`,
    sectionProperties: `${headerFooterReferences}${pageProperties()}`,
    parts: PAGE_FURNITURE_PARTS,
    relationships: MATRIX_RELATIONSHIPS,
    overrides: PAGE_FURNITURE_OVERRIDES,
  });

export const buildLayoutCorpus = async (
  outputDirectory = DEFAULT_OUTPUT_DIR,
): Promise<ReadonlyMap<FixtureName, Uint8Array>> => {
  const result = new Map<FixtureName, Uint8Array>();
  for (const fixture of fixtures()) {
    result.set(fixture.name, await buildFixture(fixture));
  }
  await mkdir(outputDirectory, { recursive: true });
  return result;
};

const writeCorpus = async (outputDirectory: string): Promise<void> => {
  const corpus = await buildLayoutCorpus(outputDirectory);
  for (const [name, contents] of corpus) {
    await Bun.write(path.join(outputDirectory, name), contents);
    console.log(`Wrote ${name} (${contents.byteLength} bytes)`);
  }
  await Bun.write(
    path.join(outputDirectory, MATRIX_MANIFEST_NAME),
    layoutInteractionMatrixManifest(),
  );
  console.log(`Wrote ${MATRIX_MANIFEST_NAME}`);
};

const checkCorpus = async (outputDirectory: string): Promise<void> => {
  const corpus = await buildLayoutCorpus(outputDirectory);
  for (const [name, expected] of corpus) {
    const fixturePath = path.join(outputDirectory, name);
    const actual = new Uint8Array(await Bun.file(fixturePath).arrayBuffer());
    if (!Bun.deepEquals(actual, expected)) {
      throw new TypeError(`${name} is stale; rebuild the synthetic layout corpus`);
    }
  }
  const manifestPath = path.join(outputDirectory, MATRIX_MANIFEST_NAME);
  if ((await Bun.file(manifestPath).text()) !== layoutInteractionMatrixManifest()) {
    throw new TypeError(`${MATRIX_MANIFEST_NAME} is stale; rebuild the synthetic layout corpus`);
  }
};

if (import.meta.main) {
  const check = process.argv.includes("--check");
  await (check ? checkCorpus(DEFAULT_OUTPUT_DIR) : writeCorpus(DEFAULT_OUTPUT_DIR));
}
