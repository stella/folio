/**
 * A shared library of small, deterministic `.docx` document shapes.
 *
 * Editor tests tend to run against rich, well-formed fixtures. The shapes here
 * are the ordinary ones a host actually hands the editor: a letter with no
 * numbering part (or no styles part at all), a document holding one kind of
 * list, headings numbered through their paragraph style, a host that seeds
 * unused list instances, and one shape per structural feature (tables, notes,
 * comments, tracked changes, header/footer, sections, fields, images, RTL and
 * CJK text).
 *
 * Every shape is built through the same public path a host uses:
 * `fromMarkdown` → model edits → `createDocx`, or an OOXML package written the
 * way a word processor writes it, and then `ensureParaIds`. Bytes are cached per shape, so
 * a suite can build each one once and parse it per case.
 *
 * Each shape names a `focus`: the text of the paragraph a test should aim at
 * (a list item, a table cell, a paragraph holding a tracked change…).
 */

import JSZip from "jszip";

import { paragraphNumberingReference } from "@stll/docx-core/model";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document, NumberingDefinitions } from "../types/document";

/** What a shape exercises, so a suite can pick the shapes relevant to it. */
export type DocumentShapeFeature =
  | "no-numbering-part"
  | "no-styles-part"
  | "list-decimal"
  | "list-bullet"
  | "style-numbering"
  | "outline-level"
  | "unused-list-instances"
  | "table"
  | "nested-table"
  | "merged-cells"
  | "footnote"
  | "endnote"
  | "comment"
  | "tracked-insertion"
  | "tracked-deletion"
  | "paragraph-property-change"
  | "run-property-change"
  | "header-footer"
  | "sections"
  | "field"
  | "hyperlink"
  | "bookmark"
  | "image"
  | "rtl"
  | "cjk";

export type DocumentShape = {
  /** Stable identifier, used in test names. */
  id: string;
  description: string;
  features: readonly DocumentShapeFeature[];
  /** Text of the paragraph a case should aim at. It occurs once in the body. */
  focus: string;
  /**
   * The package, after `ensureParaIds`, cached per shape. Its content is
   * deterministic; a package built through `createDocx` stamps its creation
   * time in `docProps/core.xml`, so the bytes of two builds can differ.
   */
  build: () => Promise<Uint8Array>;
  /** A new build, bypassing the cache. */
  rebuild: () => Promise<Uint8Array>;
};

type DocumentShapeDefinition = Omit<DocumentShape, "rebuild">;

// ============================================================================
// PACKAGE BUILDERS
// ============================================================================

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WML_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const FIXED_DATE = "2026-01-01T00:00:00Z";

const withParaIds = async (bytes: ArrayBuffer | Uint8Array): Promise<Uint8Array> =>
  (await ensureParaIds(bytes)).docx;

const fromModel = async (document: Document): Promise<Uint8Array> =>
  withParaIds(await createDocx(document));

type RawPart = {
  /** Package path, e.g. `word/footnotes.xml`. */
  path: string;
  /** Content type override; omitted for parts covered by a default. */
  contentType?: string;
  /** Relationship type (suffix of the office relationship namespace), when related from the document. */
  relationshipType?: string;
  relationshipId?: string;
  body: string | Uint8Array;
};

type RawPackage = {
  /** Children of `w:body`, final `w:sectPr` included. */
  body: string;
  /** `word/styles.xml`, or null for a package without a styles part. */
  styles: string | null;
  parts?: RawPart[];
  /** Extra document-level relationships (external hyperlinks). */
  relationships?: string[];
};

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Serialize a package the way a word processor writes it. Timestamps are
 * fixed and JSZip is given a fixed date so the bytes are reproducible.
 */
export const buildRawPackage = async ({
  body,
  styles,
  parts = [],
  relationships = [],
}: RawPackage): Promise<Uint8Array> => {
  const zip = new JSZip();
  const date = new Date(FIXED_DATE);
  const add = (path: string, data: string | Uint8Array) => zip.file(path, data, { date });

  const allParts: RawPart[] = [
    ...(styles === null
      ? []
      : [
          {
            path: "word/styles.xml",
            contentType: `${WML_CT}.styles+xml`,
            relationshipType: "styles",
            relationshipId: "rIdStyles",
            body: styles,
          },
        ]),
    ...parts,
  ];

  const overrides = allParts
    .filter((part) => part.contentType !== undefined)
    .map((part) => `<Override PartName="/${part.path}" ContentType="${part.contentType ?? ""}"/>`)
    .join("");
  add(
    "[Content_Types].xml",
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      `<Override PartName="/word/document.xml" ContentType="${WML_CT}.document.main+xml"/>` +
      overrides +
      "</Types>",
  );
  add(
    "_rels/.rels",
    `${XML_DECL}<Relationships xmlns="${REL_NS}">` +
      `<Relationship Id="rId1" Type="${OFFICE_REL}/officeDocument" Target="word/document.xml"/>` +
      "</Relationships>",
  );

  const documentRelationships = allParts
    .filter((part) => part.relationshipType !== undefined)
    .map(
      (part) =>
        `<Relationship Id="${part.relationshipId ?? ""}" Type="${OFFICE_REL}/${part.relationshipType ?? ""}" Target="${part.path.replace(/^word\//u, "")}"/>`,
    );
  add(
    "word/_rels/document.xml.rels",
    `${XML_DECL}<Relationships xmlns="${REL_NS}">${[...documentRelationships, ...relationships].join("")}</Relationships>`,
  );
  add(
    "word/document.xml",
    `${XML_DECL}<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" ${DRAWING_NAMESPACES}><w:body>${body}</w:body></w:document>`,
  );
  for (const part of allParts) {
    add(part.path, part.body);
  }

  const bytes = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  return withParaIds(bytes);
};

const DRAWING_NAMESPACES = [
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(" ");

export const SECTION = (extra = "") =>
  `<w:sectPr>${extra}<w:pgSz w:w="12240" w:h="15840"/>` +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>' +
  "</w:sectPr>";

const escapeText = (text: string): string =>
  text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const run = (text: string, rPr = ""): string =>
  `<w:r>${rPr === "" ? "" : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${escapeText(text)}</w:t></w:r>`;

const paragraph = (inner: string, pPr = ""): string =>
  `<w:p>${pPr === "" ? "" : `<w:pPr>${pPr}</w:pPr>`}${inner}</w:p>`;

const textParagraph = (text: string, pPr = ""): string => paragraph(run(text), pPr);

/** A styles part with the styles a word-processor document typically carries. */
export const WORD_STYLES = `${XML_DECL}<w:styles xmlns:w="${W_NS}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="MS Mincho" w:cs="Arial"/><w:sz w:val="22"/><w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="EndnoteText"><w:name w:val="endnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="CommentText"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/><w:basedOn w:val="Normal"/></w:style>
<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="EndnoteReference"><w:name w:val="endnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="CommentReference"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>
</w:styles>`;

const DECIMAL_LEVELS = (count: number): NumberingDefinitions["abstractNums"][number]["levels"] =>
  Array.from({ length: count }, (_, ilvl) => ({
    ilvl,
    start: 1,
    numFmt: "decimal" as const,
    lvlText: `%${ilvl + 1}.`,
    suffix: "tab" as const,
    pPr: { indentLeft: 720 * (ilvl + 1), indentFirstLine: -360 },
  }));

const BULLET_LEVELS = (count: number): NumberingDefinitions["abstractNums"][number]["levels"] =>
  Array.from({ length: count }, (_, ilvl) => ({
    ilvl,
    start: 1,
    numFmt: "bullet" as const,
    lvlText: "•",
    suffix: "tab" as const,
    pPr: { indentLeft: 720 * (ilvl + 1), indentFirstLine: -360 },
  }));

/** Headings numbered through the `Heading 2` paragraph style (the #1092 recipe). */
const styleNumberedHeadings = (markdown: string): Document => {
  const model = fromMarkdown(markdown);
  model.package.numbering = {
    abstractNums: [
      {
        abstractNumId: 5,
        multiLevelType: "multilevel",
        levels: [
          {
            ilvl: 0,
            start: 1,
            numFmt: "decimal",
            lvlText: "%1.",
            suffix: "space",
            pPr: { indentLeft: 0, indentFirstLine: 0 },
          },
          {
            ilvl: 1,
            start: 1,
            numFmt: "decimal",
            lvlText: "%1.%2.",
            suffix: "space",
            pPr: { indentLeft: 0, indentFirstLine: 0 },
          },
        ],
      },
    ],
    nums: [{ numId: 5, abstractNumId: 5 }],
  };
  const heading2 = model.package.styles?.styles.find((style) => style.styleId === "Heading2");
  if (!heading2) {
    throw new Error("fromMarkdown produced no Heading2 style");
  }
  heading2.pPr = { ...heading2.pPr, numPr: paragraphNumberingReference({ numId: 5, ilvl: 0 }) };
  return model;
};

// A 1×1 transparent PNG.
const PNG_1X1 = Uint8Array.from(
  atob(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
  ),
  (character) => character.codePointAt(0) ?? 0,
);

const INLINE_IMAGE = (relationshipId: string) =>
  "<w:r><w:drawing>" +
  '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="Picture 1" descr="A dot"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="dot.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  `<pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
  "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

const NOTES_PART = (kind: "footnote" | "endnote", text: string) => {
  const root = kind === "footnote" ? "footnotes" : "endnotes";
  const style = kind === "footnote" ? "FootnoteText" : "EndnoteText";
  const refStyle = kind === "footnote" ? "FootnoteReference" : "EndnoteReference";
  const ref = kind === "footnote" ? "footnoteRef" : "endnoteRef";
  return (
    `${XML_DECL}<w:${root} xmlns:w="${W_NS}" xmlns:r="${R_NS}">` +
    `<w:${kind} w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:${kind}>` +
    `<w:${kind} w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:${kind}>` +
    `<w:${kind} w:id="1"><w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>` +
    `<w:r><w:rPr><w:rStyle w:val="${refStyle}"/></w:rPr><w:${ref}/></w:r>${run(` ${text}`)}</w:p></w:${kind}>` +
    `</w:${root}>`
  );
};

const noteReference = (kind: "footnote" | "endnote") =>
  `<w:r><w:rPr><w:rStyle w:val="${kind === "footnote" ? "FootnoteReference" : "EndnoteReference"}"/></w:rPr><w:${kind}Reference w:id="1"/></w:r>`;

const REVISION = (id: number) => `w:id="${id}" w:author="Reviewer" w:date="${FIXED_DATE}"`;

const TRACKED_CHANGE_FEATURES = [
  "tracked-insertion",
  "tracked-deletion",
  "paragraph-property-change",
  "run-property-change",
] as const satisfies readonly DocumentShapeFeature[];

const buildTrackedChanges = (): Promise<Uint8Array> =>
  buildRawPackage({
    styles: WORD_STYLES,
    body:
      textParagraph("Intro paragraph.") +
      paragraph(
        run("Kept text ") +
          `<w:ins ${REVISION(1)}>${run("inserted text")}</w:ins>` +
          `<w:del ${REVISION(2)}><w:r><w:delText xml:space="preserve"> deleted text</w:delText></w:r></w:del>` +
          run("."),
      ) +
      paragraph(
        run("Centred by a tracked change."),
        `<w:jc w:val="center"/><w:pPrChange ${REVISION(3)}><w:pPr/></w:pPrChange>`,
      ) +
      paragraph(
        run(
          "Bold by a tracked change.",
          `<w:b/><w:rPrChange ${REVISION(4)}><w:rPr/></w:rPrChange>`,
        ),
      ) +
      textParagraph("Tail.") +
      SECTION(),
  });

// ============================================================================
// SHAPES
// ============================================================================

const cached = (build: () => Promise<Uint8Array>): (() => Promise<Uint8Array>) => {
  let pending: Promise<Uint8Array> | null = null;
  return () => {
    pending ??= build();
    return pending;
  };
};

const shape = (definition: DocumentShapeDefinition): DocumentShape => ({
  ...definition,
  build: cached(definition.build),
  rebuild: definition.build,
});

export const DOCUMENT_SHAPES: readonly DocumentShape[] = [
  shape({
    id: "bare-package",
    description: "Three paragraphs; no styles part and no numbering part",
    features: ["no-numbering-part", "no-styles-part"],
    focus: "Second paragraph of a bare letter.",
    build: () =>
      buildRawPackage({
        styles: null,
        body:
          textParagraph("Dear reader,") +
          textParagraph("Second paragraph of a bare letter.") +
          textParagraph("Kind regards.") +
          SECTION(),
      }),
  }),
  shape({
    id: "plain-markdown",
    description: "Prose from fromMarkdown; styles part, no numbering part",
    features: ["no-numbering-part"],
    focus: "First item",
    build: () => fromModel(fromMarkdown("Intro paragraph.\n\nFirst item\n\nTail paragraph.")),
  }),
  shape({
    id: "single-decimal-list",
    description: "One decimal list (instance 1) between prose",
    features: ["list-decimal"],
    focus: "Beta",
    build: () =>
      fromModel(fromMarkdown("Intro paragraph.\n\n1. Alpha\n2. Beta\n\nPlain text.\n\nTail.")),
  }),
  shape({
    id: "single-bullet-list",
    description: "One bullet list (instance 1) between prose",
    features: ["list-bullet"],
    focus: "Beta",
    build: () =>
      fromModel(fromMarkdown("Intro paragraph.\n\n- Alpha\n- Beta\n\nPlain text.\n\nTail.")),
  }),
  shape({
    id: "mixed-lists",
    description: "A decimal list and a bullet list",
    features: ["list-decimal", "list-bullet"],
    focus: "Plain text.",
    build: () =>
      fromModel(
        fromMarkdown(
          "Intro paragraph.\n\n1. Alpha\n2. Beta\n\nPlain text.\n\n- Gamma\n- Delta\n\nTail.",
        ),
      ),
  }),
  shape({
    id: "style-numbered-headings",
    description: "Heading 2 numbered through its paragraph style (instance 5)",
    features: ["style-numbering", "list-decimal"],
    focus: "The Buyer pays on delivery.",
    build: () =>
      fromModel(
        styleNumberedHeadings(
          "# Agreement\n\n## Scope\n\nThe Supplier delivers the goods.\n\nThe Buyer pays on delivery.\n\n## Payment\n\nPayment is due in ten days (see clause 1).\n\nClosing paragraph.",
        ),
      ),
  }),
  shape({
    id: "outline-level-numbered",
    description: "A body paragraph with a direct outline level and direct numbering",
    features: ["outline-level", "list-decimal"],
    focus: "Numbered outline paragraph",
    build: () => {
      const model = fromMarkdown(
        "Intro paragraph.\n\nNumbered outline paragraph\n\nBody under it.\n\nTail.",
      );
      model.package.numbering = {
        abstractNums: [
          { abstractNumId: 3, multiLevelType: "multilevel", levels: DECIMAL_LEVELS(3) },
        ],
        nums: [{ numId: 3, abstractNumId: 3 }],
      };
      const target = model.package.document.content[1];
      if (target?.type !== "paragraph") {
        throw new Error("expected a paragraph");
      }
      target.formatting = {
        ...target.formatting,
        outlineLevel: { kind: "heading", level: 0 },
        numPr: paragraphNumberingReference({ numId: 3, ilvl: 0 }),
      };
      return fromModel(model);
    },
  }),
  shape({
    id: "host-unused-instances",
    description: "Unused decimal (901) and bullet (902) instances first in nums, plus a used list",
    features: ["unused-list-instances", "list-decimal"],
    focus: "Plain text.",
    build: () => {
      const model = fromMarkdown("Intro paragraph.\n\n1. Alpha\n2. Beta\n\nPlain text.\n\nTail.");
      const existing = model.package.numbering ?? { abstractNums: [], nums: [] };
      model.package.numbering = {
        abstractNums: [
          { abstractNumId: 901, multiLevelType: "hybridMultilevel", levels: DECIMAL_LEVELS(9) },
          { abstractNumId: 902, multiLevelType: "hybridMultilevel", levels: BULLET_LEVELS(9) },
          ...existing.abstractNums,
        ],
        nums: [
          { numId: 901, abstractNumId: 901 },
          { numId: 902, abstractNumId: 902 },
          ...existing.nums,
        ],
      };
      return fromModel(model);
    },
  }),
  shape({
    id: "tables",
    description:
      "A table with a merged block left of a vertical merge, a second vertical merge and a nested table",
    features: ["table", "nested-table", "merged-cells"],
    focus: "Cell C2",
    build: () => {
      const cell = (inner: string, tcPr = "") =>
        `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${tcPr}</w:tcPr>${inner}</w:tc>`;
      const nested =
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="1400"/><w:gridCol w:w="1400"/></w:tblGrid>' +
        `<w:tr>${cell(textParagraph("Inner 1"))}${cell(textParagraph("Inner 2"))}</w:tr></w:tbl>`;
      const table =
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/><w:gridCol w:w="2200"/></w:tblGrid>' +
        // Rows 1–2: a 2×2 merged block, then a plain column, then a vertical
        // merge in the last column, so removing row 2 has to keep counting
        // columns past the wide merged cell to reach that merge.
        `<w:tr>${cell(textParagraph("Merged A1"), '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>')}${cell(textParagraph("Cell C1"))}${cell(textParagraph("Cell D1"), '<w:vMerge w:val="restart"/>')}</w:tr>` +
        `<w:tr>${cell(paragraph(""), '<w:gridSpan w:val="2"/><w:vMerge/>')}${cell(textParagraph("Cell C2"))}${cell(paragraph(""), "<w:vMerge/>")}</w:tr>` +
        // Rows 3–4: a vertical merge in column A and a nested table.
        `<w:tr>${cell(textParagraph("Cell A3"), '<w:vMerge w:val="restart"/>')}${cell(textParagraph("Cell B3"))}${cell(textParagraph("Outer") + nested + paragraph(""))}${cell(textParagraph("Cell D3"))}</w:tr>` +
        `<w:tr>${cell(paragraph(""), "<w:vMerge/>")}${cell(textParagraph("Cell B4"))}${cell(textParagraph("Cell C4"))}${cell(textParagraph("Cell D4"))}</w:tr>` +
        "</w:tbl>";
      return buildRawPackage({
        styles: WORD_STYLES,
        body:
          textParagraph("Before the table.") +
          table +
          textParagraph("After the table.") +
          SECTION(),
      });
    },
  }),
  shape({
    id: "notes",
    description: "A footnote and an endnote reference in body text",
    features: ["footnote", "endnote"],
    focus: "A sentence with a footnote",
    build: () =>
      buildRawPackage({
        styles: WORD_STYLES,
        body:
          textParagraph("Intro paragraph.") +
          paragraph(
            run("A sentence with a footnote") +
              noteReference("footnote") +
              run(" and an endnote.") +
              noteReference("endnote"),
          ) +
          textParagraph("Tail.") +
          SECTION(),
        parts: [
          {
            path: "word/footnotes.xml",
            contentType: `${WML_CT}.footnotes+xml`,
            relationshipType: "footnotes",
            relationshipId: "rIdFootnotes",
            body: NOTES_PART("footnote", "The footnote text."),
          },
          {
            path: "word/endnotes.xml",
            contentType: `${WML_CT}.endnotes+xml`,
            relationshipType: "endnotes",
            relationshipId: "rIdEndnotes",
            body: NOTES_PART("endnote", "The endnote text."),
          },
        ],
      }),
  }),
  shape({
    id: "comments",
    description: "A comment anchored on part of a paragraph",
    features: ["comment"],
    focus: "Commented words sit in this paragraph.",
    build: () =>
      buildRawPackage({
        styles: WORD_STYLES,
        body:
          textParagraph("Intro paragraph.") +
          paragraph(
            '<w:commentRangeStart w:id="0"/>' +
              run("Commented words") +
              '<w:commentRangeEnd w:id="0"/>' +
              '<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>' +
              run(" sit in this paragraph."),
          ) +
          textParagraph("Tail.") +
          SECTION(),
        parts: [
          {
            path: "word/comments.xml",
            contentType: `${WML_CT}.comments+xml`,
            relationshipType: "comments",
            relationshipId: "rIdComments",
            body:
              `${XML_DECL}<w:comments xmlns:w="${W_NS}">` +
              `<w:comment w:id="0" w:author="Reviewer" w:date="${FIXED_DATE}" w:initials="R">` +
              '<w:p><w:pPr><w:pStyle w:val="CommentText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:annotationRef/></w:r>' +
              `${run("Please check.")}</w:p></w:comment></w:comments>`,
          },
        ],
      }),
  }),
  shape({
    id: "tracked-changes",
    description: "Tracked insertion, deletion, paragraph-property and run-property changes",
    features: TRACKED_CHANGE_FEATURES,
    focus: "Kept text inserted text",
    build: buildTrackedChanges,
  }),
  shape({
    id: "pending-property-change",
    description: "The tracked-changes package, aimed at the paragraph with a pending w:pPrChange",
    features: TRACKED_CHANGE_FEATURES,
    focus: "Centred by a tracked change.",
    build: buildTrackedChanges,
  }),
  shape({
    id: "header-footer",
    description: "Default header and footer parts",
    features: ["header-footer"],
    focus: "Body paragraph under a header.",
    build: () =>
      buildRawPackage({
        styles: WORD_STYLES,
        body:
          textParagraph("Intro paragraph.") +
          textParagraph("Body paragraph under a header.") +
          textParagraph("Tail.") +
          SECTION(
            '<w:headerReference w:type="default" r:id="rIdHeader1"/><w:footerReference w:type="default" r:id="rIdFooter1"/>',
          ),
        parts: [
          {
            path: "word/header1.xml",
            contentType: `${WML_CT}.header+xml`,
            relationshipType: "header",
            relationshipId: "rIdHeader1",
            body: `${XML_DECL}<w:hdr xmlns:w="${W_NS}" xmlns:r="${R_NS}">${textParagraph("Header text", '<w:pStyle w:val="Header"/>')}</w:hdr>`,
          },
          {
            path: "word/footer1.xml",
            contentType: `${WML_CT}.footer+xml`,
            relationshipType: "footer",
            relationshipId: "rIdFooter1",
            body: `${XML_DECL}<w:ftr xmlns:w="${W_NS}" xmlns:r="${R_NS}">${textParagraph("Footer text", '<w:pStyle w:val="Footer"/>')}</w:ftr>`,
          },
        ],
      }),
  }),
  shape({
    id: "sections",
    description: "Two sections: a paragraph-level sectPr, then a landscape final section",
    features: ["sections"],
    focus: "Last paragraph of section one.",
    build: () =>
      buildRawPackage({
        styles: WORD_STYLES,
        body:
          textParagraph("Intro paragraph.") +
          paragraph(run("Last paragraph of section one."), SECTION('<w:type w:val="nextPage"/>')) +
          textParagraph("First paragraph of section two.") +
          textParagraph("Tail.") +
          '<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>',
      }),
  }),
  shape({
    id: "fields-links-bookmarks",
    description:
      "A simple field, a complex field, an external and an internal hyperlink, a bookmark",
    features: ["field", "hyperlink", "bookmark"],
    focus: "See the site and the target.",
    build: () =>
      buildRawPackage({
        styles: WORD_STYLES,
        relationships: [
          `<Relationship Id="rIdLink" Type="${OFFICE_REL}/hyperlink" Target="https://example.com/" TargetMode="External"/>`,
        ],
        body:
          paragraph(
            '<w:bookmarkStart w:id="0" w:name="Target"/>' +
              run("Bookmarked heading text") +
              '<w:bookmarkEnd w:id="0"/>',
          ) +
          paragraph(
            run("See the ") +
              `<w:hyperlink r:id="rIdLink" w:history="1">${run("site", '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>` +
              run(" and the ") +
              `<w:hyperlink w:anchor="Target" w:history="1">${run("target", '<w:rStyle w:val="Hyperlink"/>')}</w:hyperlink>` +
              run("."),
          ) +
          paragraph(
            run("Page ") +
              '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>' +
              run(" of ") +
              '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES </w:instrText></w:r>' +
              '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
          ) +
          textParagraph("Tail.") +
          SECTION(),
      }),
  }),
  shape({
    id: "image",
    description: "An inline picture between two runs",
    features: ["image"],
    focus: "Text before the picture and after it.",
    build: () =>
      buildRawPackage({
        styles: WORD_STYLES,
        body:
          textParagraph("Intro paragraph.") +
          paragraph(
            run("Text before the picture") + INLINE_IMAGE("rIdImage1") + run(" and after it."),
          ) +
          textParagraph("Tail.") +
          SECTION(),
        parts: [
          {
            path: "word/media/image1.png",
            relationshipType: "image",
            relationshipId: "rIdImage1",
            body: PNG_1X1,
          },
        ],
      }),
  }),
  shape({
    id: "rtl-cjk",
    description: "A right-to-left paragraph (Arabic and Hebrew) and a CJK paragraph",
    features: ["rtl", "cjk"],
    focus: "مرحبا بالعالم שלום עולם",
    build: () =>
      buildRawPackage({
        styles: WORD_STYLES,
        body:
          textParagraph("Intro paragraph.") +
          paragraph(run("مرحبا بالعالم שלום עולם", "<w:rtl/>"), '<w:bidi/><w:jc w:val="right"/>') +
          paragraph(
            run(
              "日本語の段落です。中文段落。",
              '<w:rFonts w:hint="eastAsia"/><w:lang w:eastAsia="ja-JP"/>',
            ),
          ) +
          textParagraph("Tail.") +
          SECTION(),
      }),
  }),
];

export const documentShape = (id: string): DocumentShape => {
  const found = DOCUMENT_SHAPES.find((candidate) => candidate.id === id);
  if (!found) {
    throw new Error(`Unknown document shape "${id}"`);
  }
  return found;
};

/** A copy of the shape's bytes as an `ArrayBuffer`, safe to hand to a parser. */
export const shapeArrayBuffer = async (shapeOrId: DocumentShape | string): Promise<ArrayBuffer> => {
  const bytes = await (
    typeof shapeOrId === "string" ? documentShape(shapeOrId) : shapeOrId
  ).build();
  return bytes.slice().buffer;
};
