import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { extractDocumentStyleSet } from "../style-sets/extract";
import { createEmptyDocument } from "../utils/createDocument";
import type { DocumentBody, NumberingDefinitions, Style } from "../types/document";
import { createNumberingMap, parseNumbering } from "./numberingParser";
import {
  normalizeNumberingReferences,
  normalizeStyleNumberingReferences,
} from "./numberingReferenceNormalization";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { createDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** One `w:num` that resolves, and one whose `w:abstractNum` is missing. */
const NUMBERING: NumberingDefinitions = {
  abstractNums: [{ abstractNumId: 0, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] }],
  nums: [
    { numId: 1, abstractNumId: 0 },
    { numId: 7, abstractNumId: 9 },
  ],
};

describe("normalizeNumberingReferences", () => {
  test("unnumbers a paragraph whose numbering definition is missing", () => {
    const documentBody: DocumentBody = {
      content: [
        {
          type: "paragraph",
          formatting: { numPr: { numId: 2, ilvl: 0 } },
          content: [],
        },
      ],
    };

    const result = normalizeNumberingReferences({
      documentBody,
      numbering: parseNumbering(null),
    });

    expect(result).toEqual({ unnumberedDanglingReferences: 1 });
    const block = documentBody.content.at(0);
    expect(block?.type).toBe("paragraph");
    if (block?.type !== "paragraph") {
      throw new Error("Expected first block to be a paragraph");
    }
    // The sentinel, not a deletion: deleting would uncover the style's numbering.
    expect(block.formatting?.numPr).toEqual({ numId: 0 });
  });

  test("unnumbers a paragraph whose abstract numbering is missing", () => {
    const documentBody: DocumentBody = {
      content: [
        { type: "paragraph", formatting: { numPr: { numId: 7, ilvl: 0 } }, content: [] },
        { type: "paragraph", formatting: { numPr: { numId: 1, ilvl: 0 } }, content: [] },
      ],
    };

    const result = normalizeNumberingReferences({
      documentBody,
      numbering: createNumberingMap(NUMBERING),
    });

    expect(result).toEqual({ unnumberedDanglingReferences: 1 });
    expect(documentBody.content.at(0)?.formatting?.numPr).toEqual({ numId: 0 });
    expect(documentBody.content.at(1)?.formatting?.numPr).toEqual({ numId: 1, ilvl: 0 });
  });

  test("leaves the no-numbering sentinel untouched", () => {
    const documentBody: DocumentBody = {
      content: [{ type: "paragraph", formatting: { numPr: { numId: 0 } }, content: [] }],
    };

    const result = normalizeNumberingReferences({
      documentBody,
      numbering: parseNumbering(null),
    });

    expect(result).toEqual({ unnumberedDanglingReferences: 0 });
    expect(documentBody.content.at(0)?.formatting?.numPr).toEqual({ numId: 0 });
  });

  test("unnumbers missing numbering references in comment paragraphs", () => {
    const documentBody: DocumentBody = {
      content: [],
      comments: [
        {
          id: 1,
          author: "Reviewer",
          content: [
            {
              type: "paragraph",
              formatting: { numPr: { numId: 2, ilvl: 0 } },
              content: [],
            },
          ],
        },
      ],
    };

    const result = normalizeNumberingReferences({
      documentBody,
      numbering: parseNumbering(null),
    });

    expect(result).toEqual({ unnumberedDanglingReferences: 1 });
    expect(documentBody.comments?.at(0)?.content.at(0)?.formatting?.numPr).toEqual({ numId: 0 });
  });

  test("ignores negative list levels when numbering is explicitly disabled", async () => {
    const buffer = await createDocxFixture({
      bodyXml: paragraphXml({ ilvl: -1, numId: 0 }),
    });
    const doc = await parseDocx(buffer, { preloadFonts: false });
    const block = doc.package.document.content.at(0);

    expect(block?.type).toBe("paragraph");
    if (block?.type !== "paragraph") {
      throw new Error("Expected first block to be a paragraph");
    }
    expect(block.formatting?.numPr).toEqual({ numId: 0 });
  });

  test("parses and saves documents that reference a missing numbering part", async () => {
    const buffer = await createDocxFixture({ bodyXml: paragraphXml({ ilvl: 0, numId: 2 }) });
    const doc = await parseDocx(buffer, { preloadFonts: false });

    expect(doc.package.document.content.at(0)?.formatting?.numPr).toEqual({ numId: 0 });
    expect(doc.warnings).toContain(
      "Unnumbered 1 paragraph whose numbering definitions are missing.",
    );

    const repacked = await repackDocx(doc, { updateModifiedDate: false });
    const documentXml = await readPart(repacked, "word/document.xml");

    expect(documentXml).toContain('<w:numId w:val="0"/>');

    const reparsed = await parseDocx(repacked, { preloadFonts: false });
    expect(reparsed.package.document.content.at(0)?.formatting?.numPr).toEqual({ numId: 0 });
  });

  test("keeps a paragraph unnumbered when its style is numbered and its own reference dangles", async () => {
    const buffer = await createDocxFixture({
      bodyXml: paragraphXml({ ilvl: 0, numId: 42, styleId: "Numbered" }),
      stylesXml: stylesXml([{ styleId: "Numbered", numId: 1 }]),
      numberingXml: numberingXml([{ numId: 1, abstractNumId: 0 }]),
    });
    const doc = await parseDocx(buffer, { preloadFonts: false });

    expect(doc.package.document.content.at(0)?.formatting?.numPr).toEqual({ numId: 0 });
    expect(doc.package.document.content.at(0)?.listRendering).toBeUndefined();

    const reparsed = await parseDocx(await repackDocx(doc, { updateModifiedDate: false }), {
      preloadFonts: false,
    });

    // Deleting the numPr would have handed the paragraph the style's list back.
    expect(reparsed.package.document.content.at(0)?.formatting?.numPr).toEqual({ numId: 0 });
    expect(reparsed.package.document.content.at(0)?.listRendering).toBeUndefined();
  });
});

describe("normalizeStyleNumberingReferences", () => {
  test("unnumbers a style whose numbering definition is missing", () => {
    const styles: Style[] = [
      { styleId: "Dangling", type: "paragraph", pPr: { numPr: { numId: 4, ilvl: 0 } } },
    ];

    const result = normalizeStyleNumberingReferences({
      styles,
      numbering: createNumberingMap(NUMBERING),
    });

    expect(result).toEqual({ unnumberedStyleIds: ["Dangling"] });
    expect(styles.at(0)?.pPr?.numPr).toEqual({ numId: 0 });
  });

  test("unnumbers a style whose abstract numbering is missing", () => {
    const styles: Style[] = [
      { styleId: "NoAbstract", type: "paragraph", pPr: { numPr: { numId: 7, ilvl: 0 } } },
    ];

    const result = normalizeStyleNumberingReferences({
      styles,
      numbering: createNumberingMap(NUMBERING),
    });

    expect(result).toEqual({ unnumberedStyleIds: ["NoAbstract"] });
    expect(styles.at(0)?.pPr?.numPr).toEqual({ numId: 0 });
  });

  test("leaves the sentinel and resolvable references untouched", () => {
    const styles: Style[] = [
      { styleId: "Sentinel", type: "paragraph", pPr: { numPr: { numId: 0 } } },
      { styleId: "Numbered", type: "paragraph", pPr: { numPr: { numId: 1, ilvl: 0 } } },
      { styleId: "Plain", type: "paragraph" },
    ];

    const result = normalizeStyleNumberingReferences({
      styles,
      numbering: createNumberingMap(NUMBERING),
    });

    expect(result).toEqual({ unnumberedStyleIds: [] });
    expect(styles.at(0)?.pPr?.numPr).toEqual({ numId: 0 });
    expect(styles.at(1)?.pPr?.numPr).toEqual({ numId: 1, ilvl: 0 });
    expect(styles.at(2)?.pPr).toBeUndefined();
  });

  test("treats a missing numbering part as defining nothing", () => {
    const styles: Style[] = [
      { styleId: "Dangling", type: "paragraph", pPr: { numPr: { numId: 1 } } },
    ];

    const result = normalizeStyleNumberingReferences({ styles, numbering: undefined });

    expect(result).toEqual({ unnumberedStyleIds: ["Dangling"] });
    expect(styles.at(0)?.pPr?.numPr).toEqual({ numId: 0 });
  });

  test("carries the sentinel, not its parent's numbering, into a saved style set", async () => {
    const buffer = await createDocxFixture({
      bodyXml: paragraphXml({ styleId: "Dangling" }),
      stylesXml: stylesXml([
        { styleId: "Numbered", numId: 1 },
        { styleId: "Dangling", basedOn: "Numbered", numId: 42 },
      ]),
      numberingXml: numberingXml([{ numId: 1, abstractNumId: 0 }]),
    });
    const doc = await parseDocx(buffer, { preloadFonts: false });

    expect(doc.warnings).toContain(
      'Unnumbered style "Dangling" whose numbering definition is missing.',
    );
    expect(styleById(doc.package.styles?.styles, "Dangling")?.pPr?.numPr).toEqual({ numId: 0 });
    expect(styleById(doc.package.styles?.styles, "Numbered")?.pPr?.numPr?.numId).toBe(1);

    const styleSet = extractDocumentStyleSet(doc, {
      name: "Set",
      initialParagraphStyleId: "Normal",
    });
    expect(styleById(styleSet.styles.styles, "Dangling")?.pPr?.numPr).toEqual({ numId: 0 });

    // createDocx asserts its own style numbering: the seed path must not panic.
    const saved = await createDocx(createEmptyDocument({ styleSet }));
    const savedStyles = await readPart(saved, "word/styles.xml");

    expect(styleElement(savedStyles, "Dangling")).toContain('<w:numId w:val="0"/>');
    expect(styleElement(savedStyles, "Numbered")).toContain('<w:numId w:val="1"/>');
  });

  test("does not panic when a style set is extracted from an unparsed document", async () => {
    const document = createEmptyDocument();
    document.package.styles = {
      styles: [
        { styleId: "Normal", type: "paragraph", default: true },
        { styleId: "Dangling", type: "paragraph", pPr: { numPr: { numId: 42 } } },
      ],
    };

    const styleSet = extractDocumentStyleSet(document, { name: "Set" });

    expect(styleById(styleSet.styles.styles, "Dangling")?.pPr?.numPr).toEqual({ numId: 0 });
    expect(styleSet.numbering).toBeUndefined();
    // The source document keeps what the caller handed over.
    expect(styleById(document.package.styles.styles, "Dangling")?.pPr?.numPr).toEqual({
      numId: 42,
    });

    const savedStyles = await readPart(
      await createDocx(createEmptyDocument({ styleSet })),
      "word/styles.xml",
    );
    expect(styleElement(savedStyles, "Dangling")).toContain('<w:numId w:val="0"/>');
  });
});

const styleById = (styles: readonly Style[] | undefined, styleId: string): Style | undefined =>
  styles?.find((style) => style.styleId === styleId);

/** The serialized `<w:style>` element for one style id. */
const styleElement = (stylesXml: string, styleId: string): string => {
  const start = stylesXml.indexOf(`w:styleId="${styleId}">`);
  expect(start).toBeGreaterThanOrEqual(0);
  const end = stylesXml.indexOf("</w:style>", start);
  expect(end).toBeGreaterThan(start);
  return stylesXml.slice(start, end);
};

const readPart = async (docx: ArrayBuffer, path: string): Promise<string> => {
  const part = await (await JSZip.loadAsync(docx)).file(path)?.async("string");
  if (part === undefined) {
    throw new Error(`Expected ${path} in the package`);
  }
  return part;
};

type ParagraphFixtureOptions = {
  ilvl?: number;
  numId?: number;
  styleId?: string;
};

const paragraphXml = ({ ilvl, numId, styleId }: ParagraphFixtureOptions): string => {
  const numPr =
    numId === undefined
      ? ""
      : `<w:numPr>${ilvl === undefined ? "" : `<w:ilvl w:val="${String(ilvl)}"/>`}<w:numId w:val="${String(numId)}"/></w:numPr>`;
  const pStyle = styleId === undefined ? "" : `<w:pStyle w:val="${styleId}"/>`;
  return `<w:p><w:pPr>${pStyle}${numPr}</w:pPr><w:r><w:t>Body</w:t></w:r></w:p>`;
};

type StyleFixtureOptions = {
  styleId: string;
  basedOn?: string;
  numId?: number;
};

const stylesXml = (styles: readonly StyleFixtureOptions[]): string =>
  `${XML_DECLARATION}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  ${styles
    .map(
      ({ styleId, basedOn, numId }) =>
        `<w:style w:type="paragraph" w:styleId="${styleId}">${
          basedOn === undefined ? "" : `<w:basedOn w:val="${basedOn}"/>`
        }${
          numId === undefined
            ? ""
            : `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${String(numId)}"/></w:numPr></w:pPr>`
        }</w:style>`,
    )
    .join("")}
</w:styles>`;

type NumFixtureOptions = {
  numId: number;
  abstractNumId: number;
};

const numberingXml = (nums: readonly NumFixtureOptions[]): string =>
  `${XML_DECLARATION}
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
  </w:abstractNum>
  ${nums
    .map(
      ({ numId, abstractNumId }) =>
        `<w:num w:numId="${String(numId)}"><w:abstractNumId w:val="${String(abstractNumId)}"/></w:num>`,
    )
    .join("")}
</w:numbering>`;

type DocxFixtureOptions = {
  bodyXml: string;
  stylesXml?: string;
  numberingXml?: string;
};

const createDocxFixture = async ({
  bodyXml,
  stylesXml: styles,
  numberingXml: numbering,
}: DocxFixtureOptions): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  const parts = [
    ...(styles === undefined ? [] : [{ name: "styles", type: RELATIONSHIP_TYPES.styles }]),
    ...(numbering === undefined ? [] : [{ name: "numbering", type: RELATIONSHIP_TYPES.numbering }]),
  ];
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  ${parts
    .map(
      ({ name }) =>
        `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name}+xml"/>`,
    )
    .join("")}
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${parts
    .map(
      ({ name, type }, index) =>
        `<Relationship Id="rId${String(index + 1)}" Type="${type}" Target="${name}.xml"/>`,
    )
    .join("")}
</Relationships>`,
  );
  if (styles !== undefined) {
    zip.file("word/styles.xml", styles);
  }
  if (numbering !== undefined) {
    zip.file("word/numbering.xml", numbering);
  }
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>${bodyXml}</w:body>
</w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};
