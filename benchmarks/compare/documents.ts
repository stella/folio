/**
 * The benchmark's document classes, generated as raw OOXML.
 *
 * Each class isolates one thing a compare engine can get wrong: plain prose
 * says how fast the common path is, multi-level lists carry numbering that a
 * redline must not renumber silently, tables carry a three-level structure a
 * flat alignment cannot see, notes and section stories are scopes the main
 * story alignment must not reach into, indivisible atoms (images, equations,
 * fields, breaks) must not be split down the middle, bidirectional and CJK
 * text breaks word-granularity assumptions, and a base that already carries
 * revisions forces the engine to say which view it compares.
 *
 * Sizes are block counts, not byte counts: the engine's cost tracks blocks.
 */

import { escapeXml, type DocxPackage, type PackagePart } from "./package-xml";

export const DOCUMENT_CLASSES = Object.freeze([
  "prose",
  "lists",
  "tables",
  "notes",
  "graphics",
  "fields",
  "sections",
  "multiscript",
  "revised",
] as const);

export type DocumentClass = (typeof DOCUMENT_CLASSES)[number];

export const DOCUMENT_SIZES = Object.freeze(["s", "m", "l"] as const);

export type DocumentSize = (typeof DOCUMENT_SIZES)[number];

/**
 * Blocks per size. The large tier is deliberately above the 2,000-block mark
 * the performance target is stated against.
 */
const BLOCKS_BY_SIZE = {
  s: 40,
  m: 320,
  l: 2200,
} as const satisfies Record<DocumentSize, number>;

export const blockCountFor = (size: DocumentSize): number => BLOCKS_BY_SIZE[size];

/**
 * A 32-bit linear congruential generator. Deterministic and seeded per
 * document, so the corpus is reproducible without committing its bytes.
 */
const createRandom = (seed: number): (() => number) => {
  let state = seed >>> 0;
  return (): number => {
    state = (Math.imul(state, 1_664_525) + 1_013_904_223) >>> 0;
    return state / 0x1_00_00_00_00;
  };
};

const LATIN_WORDS = Object.freeze([
  "agreement",
  "clause",
  "consideration",
  "counterparty",
  "delivery",
  "effective",
  "indemnity",
  "jurisdiction",
  "liability",
  "material",
  "notice",
  "obligation",
  "provision",
  "remedy",
  "schedule",
  "termination",
  "warranty",
  "written",
] as const);

const ARABIC_WORDS = Object.freeze([
  "الاتفاق",
  "البند",
  "التسليم",
  "المسؤولية",
  "الإشعار",
  "الالتزام",
  "الضمان",
  "الإنهاء",
] as const);

const HEBREW_WORDS = Object.freeze([
  "הסכם",
  "סעיף",
  "אחריות",
  "הודעה",
  "התחייבות",
  "סיום",
] as const);

const CJK_WORDS = Object.freeze([
  "契約",
  "条項",
  "責任",
  "通知",
  "義務",
  "保証",
  "解除",
  "引渡し",
] as const);

type Sentence = { words: readonly string[]; terminator: string };

const SENTENCE_SHAPES = {
  latin: { words: LATIN_WORDS, terminator: "." },
  arabic: { words: ARABIC_WORDS, terminator: "." },
  hebrew: { words: HEBREW_WORDS, terminator: "." },
  cjk: { words: CJK_WORDS, terminator: "。" },
} as const satisfies Record<string, Sentence>;

type ScriptName = keyof typeof SENTENCE_SHAPES;

const sentence = (random: () => number, script: ScriptName, wordCount: number): string => {
  const { words, terminator } = SENTENCE_SHAPES[script];
  const picked: string[] = [];
  for (let index = 0; index < wordCount; index++) {
    // SAFETY: the modulus keeps the index inside the frozen word list.
    picked.push(words[Math.floor(random() * words.length) % words.length] ?? "");
  }
  const joiner = script === "cjk" ? "" : " ";
  return `${picked.join(joiner)}${terminator}`;
};

const run = (text: string, properties = ""): string =>
  `<w:r>${properties}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;

const paragraph = (inner: string, properties = ""): string => `<w:p>${properties}${inner}</w:p>`;

const DOCUMENT_OPEN =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  "<w:body>";

const SECTION_PROPERTIES =
  '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>';

const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/></w:style>' +
  "</w:styles>";

const LIST_NUMBER_FORMATS = Object.freeze(["decimal", "lowerLetter", "lowerRoman"] as const);

const numberingXml = (): string => {
  const levels = Array.from({ length: 3 }, (_unused, level) => {
    const format = LIST_NUMBER_FORMATS[level] ?? "decimal";
    return (
      `<w:lvl w:ilvl="${String(level)}"><w:start w:val="1"/><w:numFmt w:val="${format}"/>` +
      `<w:lvlText w:val="%${String(level + 1)}."/><w:lvlJc w:val="left"/>` +
      `<w:pPr><w:ind w:left="${String(720 * (level + 1))}" w:hanging="360"/></w:pPr></w:lvl>`
    );
  }).join("");
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:abstractNum w:abstractNumId="0">${levels}</w:abstractNum>` +
    '<w:num w:numId="1"><w:abstractNumOverride w:val="0"/><w:abstractNumId w:val="0"/></w:num>' +
    "</w:numbering>"
  );
};

/** A 1x1 opaque PNG, spelled out so the corpus needs no binary fixture. */
const ONE_PIXEL_PNG = Uint8Array.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x03, 0x01, 0x01, 0x00, 0x18, 0xdd, 0x8d, 0xb0, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82,
]);

const DRAWING =
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="190500" cy="190500"/><wp:docPr id="1" name="Picture 1"/>' +
  "<a:graphic><a:graphicData" +
  ' uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
  '<pic:nvPicPr><pic:cNvPr id="0" name="pixel.png"/><pic:cNvPicPr/></pic:nvPicPr>' +
  '<pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
  "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

const EQUATION =
  "<m:oMath><m:sSup><m:e><m:r><m:t>a</m:t></m:r></m:e>" +
  "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath>";

const SIMPLE_FIELD = '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>';

const complexField = (instruction: string, result: string): string =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
  `<w:r><w:instrText xml:space="preserve">${escapeXml(instruction)}</w:instrText></w:r>` +
  `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
  `<w:r><w:t xml:space="preserve">${escapeXml(result)}</w:t></w:r>` +
  `<w:r><w:fldChar w:fldCharType="end"/></w:r>`;

const HYPERLINK = `<w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>the schedule</w:t></w:r></w:hyperlink>`;

type ClassBuilderOptions = {
  blocks: number;
  random: () => number;
};

type ClassBuild = {
  /** Body children, excluding the trailing section properties. */
  body: string;
  /** Extra parts this class needs beyond the document, styles, and rels. */
  parts?: ReadonlyMap<string, PackagePart>;
  /** Extra document-part relationships, as `<Relationship .../>` XML. */
  relationships?: readonly string[];
  /** Extra `[Content_Types].xml` overrides. */
  overrides?: readonly string[];
  /** Extra default extensions, as `[extension, contentType]`. */
  defaults?: readonly (readonly [string, string])[];
};

const buildProse = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  for (let index = 0; index < blocks; index++) {
    if (index % 20 === 0) {
      body.push(
        paragraph(
          run(`Article ${String(index / 20 + 1)}`),
          '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>',
        ),
      );
      continue;
    }
    body.push(paragraph(run(sentence(random, "latin", 12 + Math.floor(random() * 10)))));
  }
  return { body: body.join("") };
};

/** Every seventh item is top level, every third is one deep, the rest two. */
const listLevelFor = (index: number): number => {
  if (index % 7 === 0) {
    return 0;
  }
  return index % 3 === 0 ? 1 : 2;
};

const buildLists = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  for (let index = 0; index < blocks; index++) {
    const level = listLevelFor(index);
    const properties =
      '<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr>' +
      `<w:ilvl w:val="${String(level)}"/><w:numId w:val="1"/></w:numPr></w:pPr>`;
    body.push(paragraph(run(sentence(random, "latin", 8 + level * 3)), properties));
  }
  return {
    body: body.join(""),
    parts: new Map([["word/numbering.xml", numberingXml()]]),
    relationships: [
      '<Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>',
    ],
    overrides: [
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
    ],
  };
};

type CellOptions = { inner: string; gridSpan?: number; vMerge?: "restart" | "continue" };

const tableCell = ({ inner, gridSpan, vMerge }: CellOptions): string => {
  const properties =
    '<w:tcPr><w:tcW w:w="3000" w:type="dxa"/>' +
    (gridSpan === undefined ? "" : `<w:gridSpan w:val="${String(gridSpan)}"/>`) +
    (vMerge === undefined ? "" : `<w:vMerge${vMerge === "restart" ? ' w:val="restart"' : ""}/>`) +
    "</w:tcPr>";
  return `<w:tc>${properties}${inner}</w:tc>`;
};

const TABLE_OPEN =
  '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>';

const buildTables = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  let emitted = 0;
  let tableIndex = 0;
  while (emitted < blocks) {
    body.push(paragraph(run(sentence(random, "latin", 10))));
    emitted += 1;

    const rowCount = Math.min(Math.max(3, Math.floor(blocks / 12)), 60);
    const rows: string[] = [];
    for (let rowIndex = 0; rowIndex < rowCount && emitted < blocks; rowIndex++) {
      const cells: string[] = [];
      // Row 0 of every table spans two columns; rows 1 and 2 vertically merge
      // the first column, so the class carries both merge shapes.
      if (rowIndex === 0) {
        cells.push(
          tableCell({ inner: paragraph(run(`Heading ${String(tableIndex)}`)), gridSpan: 2 }),
          tableCell({ inner: paragraph(run(sentence(random, "latin", 4))) }),
        );
        emitted += 2;
      } else {
        const nested =
          rowIndex === 1
            ? `${TABLE_OPEN}<w:tr>${tableCell({ inner: paragraph(run(sentence(random, "latin", 3))) })}</w:tr></w:tbl>${paragraph("")}`
            : paragraph(run(sentence(random, "latin", 6)));
        cells.push(
          tableCell({
            inner: paragraph(run(sentence(random, "latin", 5))),
            vMerge: rowIndex === 1 ? "restart" : "continue",
          }),
          tableCell({ inner: nested }),
          // An empty cell: it carries no text at all, which is where a
          // block-only model loses a column.
          tableCell({ inner: paragraph("") }),
        );
        emitted += rowIndex === 1 ? 4 : 3;
      }
      rows.push(`<w:tr>${cells.join("")}</w:tr>`);
    }
    body.push(`${TABLE_OPEN}${rows.join("")}</w:tbl>`);
    tableIndex += 1;
  }
  return { body: body.join("") };
};

const NOTES_PER_PARAGRAPH = 2;

const buildNotes = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  const footnotes: string[] = [
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>',
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>',
  ];
  const endnotes: string[] = [
    '<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>',
    '<w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>',
  ];
  let noteId = 1;
  for (let index = 0; index < blocks; index++) {
    const references: string[] = [];
    for (let note = 0; note < NOTES_PER_PARAGRAPH; note++) {
      const isEndnote = note % 2 === 1;
      const reference = isEndnote
        ? `<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr><w:endnoteReference w:id="${String(noteId)}"/></w:r>`
        : `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${String(noteId)}"/></w:r>`;
      references.push(reference);
      const noteBody = `<w:p>${run(sentence(random, "latin", 9))}</w:p>`;
      (isEndnote ? endnotes : footnotes).push(
        isEndnote
          ? `<w:endnote w:id="${String(noteId)}">${noteBody}</w:endnote>`
          : `<w:footnote w:id="${String(noteId)}">${noteBody}</w:footnote>`,
      );
      noteId += 1;
    }
    body.push(paragraph(`${run(sentence(random, "latin", 12))}${references.join("")}`));
  }
  const header =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
  const endnoteHeader =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
  return {
    body: body.join(""),
    parts: new Map([
      ["word/footnotes.xml", `${header}${footnotes.join("")}</w:footnotes>`],
      ["word/endnotes.xml", `${endnoteHeader}${endnotes.join("")}</w:endnotes>`],
    ]),
    relationships: [
      '<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>',
      '<Relationship Id="rIdEndnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>',
    ],
    overrides: [
      '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
      '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>',
    ],
  };
};

/** Run content a redline must move whole: splitting one is a corrupt document. */
const INDIVISIBLE_ATOMS = Object.freeze([DRAWING, EQUATION, "<w:r><w:br/></w:r>"] as const);

const buildGraphics = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  for (let index = 0; index < blocks; index++) {
    const atom = INDIVISIBLE_ATOMS[index % INDIVISIBLE_ATOMS.length] ?? DRAWING;
    body.push(
      paragraph(`${run(sentence(random, "latin", 8))}${atom}${run(sentence(random, "latin", 6))}`),
    );
  }
  return {
    body: body.join(""),
    parts: new Map<string, PackagePart>([["word/media/pixel.png", ONE_PIXEL_PNG]]),
    relationships: [
      '<Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pixel.png"/>',
    ],
    defaults: [["png", "image/png"]],
  };
};

const INLINE_REFERENCES = Object.freeze([
  (): string => SIMPLE_FIELD,
  (index: number): string =>
    complexField(` REF _Ref${String(index)} \\h `, `clause ${String(index)}`),
  (): string => HYPERLINK,
] as const);

const buildFields = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  for (let index = 0; index < blocks; index++) {
    const inline = INLINE_REFERENCES[index % INLINE_REFERENCES.length]?.(index) ?? SIMPLE_FIELD;
    body.push(
      `<w:bookmarkStart w:id="${String(index)}" w:name="_Ref${String(index)}"/>` +
        paragraph(`${run(sentence(random, "latin", 9))}${inline}${run(" applies.")}`) +
        `<w:bookmarkEnd w:id="${String(index)}"/>`,
    );
  }
  return {
    body: body.join(""),
    relationships: [
      '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/schedule" TargetMode="External"/>',
    ],
  };
};

const headerFooterXml = (element: "hdr" | "ftr", text: string): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:${element} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
  `<w:p>${run(text)}</w:p></w:${element}>`;

const SECTION_COUNT = 3;

const buildSections = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  const perSection = Math.max(1, Math.floor(blocks / SECTION_COUNT));
  for (let index = 0; index < blocks; index++) {
    body.push(paragraph(run(sentence(random, "latin", 11))));
    const isSectionBreak = (index + 1) % perSection === 0 && index + 1 < blocks;
    if (isSectionBreak) {
      body.push(
        paragraph(
          "",
          "<w:pPr><w:sectPr>" +
            '<w:headerReference w:type="default" r:id="rIdHeader"/>' +
            '<w:footerReference w:type="default" r:id="rIdFooter"/>' +
            '<w:pgSz w:w="12240" w:h="15840"/>' +
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>' +
            "</w:sectPr></w:pPr>",
        ),
      );
    }
  }
  return {
    body: body.join(""),
    parts: new Map([
      ["word/header1.xml", headerFooterXml("hdr", "Confidential draft")],
      ["word/footer1.xml", headerFooterXml("ftr", "Page of the schedule")],
    ]),
    relationships: [
      '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    ],
    overrides: [
      '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>',
      '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>',
    ],
  };
};

const buildMultiscript = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  const scripts = ["arabic", "hebrew", "cjk", "latin"] as const;
  for (let index = 0; index < blocks; index++) {
    const script = scripts[index % scripts.length] ?? "latin";
    const rightToLeft = script === "arabic" || script === "hebrew";
    const properties = rightToLeft ? '<w:pPr><w:bidi/><w:jc w:val="right"/></w:pPr>' : "";
    const runProperties = rightToLeft ? "<w:rPr><w:rtl/></w:rPr>" : "";
    body.push(paragraph(run(sentence(random, script, 9), runProperties), properties));
  }
  return { body: body.join("") };
};

const buildRevised = ({ blocks, random }: ClassBuilderOptions): ClassBuild => {
  const body: string[] = [];
  const stamp = 'w:author="prior reviewer" w:date="2000-01-01T00:00:00Z"';
  for (let index = 0; index < blocks; index++) {
    if (index % 5 === 0) {
      body.push(
        paragraph(
          `<w:ins w:id="${String(index * 2 + 1)}" ${stamp}>${run(sentence(random, "latin", 7))}</w:ins>`,
        ),
      );
      continue;
    }
    if (index % 5 === 1) {
      const deleted = escapeXml(sentence(random, "latin", 6));
      body.push(
        paragraph(
          `${run(sentence(random, "latin", 6))}` +
            `<w:del w:id="${String(index * 2 + 1)}" ${stamp}>` +
            `<w:r><w:delText xml:space="preserve">${deleted}</w:delText></w:r></w:del>`,
        ),
      );
      continue;
    }
    body.push(paragraph(run(sentence(random, "latin", 12))));
  }
  return { body: body.join("") };
};

const BUILDERS = {
  prose: buildProse,
  lists: buildLists,
  tables: buildTables,
  notes: buildNotes,
  graphics: buildGraphics,
  fields: buildFields,
  sections: buildSections,
  multiscript: buildMultiscript,
  revised: buildRevised,
} as const satisfies Record<DocumentClass, (options: ClassBuilderOptions) => ClassBuild>;

/** Distinct per class and size, so no two documents share a word stream. */
const seedFor = (documentClass: DocumentClass, size: DocumentSize): number => {
  let seed = 0x9e_37_79_b9;
  for (const character of `${documentClass}:${size}`) {
    seed = (Math.imul(seed ^ character.charCodeAt(0), 0x01_00_01_93) + 1) >>> 0;
  }
  return seed;
};

/**
 * A body may not end with a table: the format requires a paragraph after one,
 * and the section properties do not supply it. A class that ends on a table
 * gets that paragraph here, once, rather than every builder remembering.
 *
 * A document that breaks the rule is malformed INPUT, which the engine has to
 * survive but which does not belong in a corpus measuring what the engine does
 * with well-formed documents: with no paragraph at the end there is no anchor
 * to append after, so a comparison that has to add a table there can only
 * report the difference it cannot place.
 */
const wellFormedBody = (body: string): string =>
  body.endsWith("</w:tbl>") ? `${body}<w:p/>` : body;

export type BuildDocumentOptions = {
  documentClass: DocumentClass;
  size: DocumentSize;
};

export const buildDocumentPackage = ({
  documentClass,
  size,
}: BuildDocumentOptions): DocxPackage => {
  const build = BUILDERS[documentClass]({
    blocks: BLOCKS_BY_SIZE[size],
    random: createRandom(seedFor(documentClass, size)),
  });

  const relationships = [
    '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    ...(build.relationships ?? []),
  ];
  const overrides = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
    ...(build.overrides ?? []),
  ];
  const defaults = [
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ["xml", "application/xml"],
    ...(build.defaults ?? []),
  ] as const;

  const parts = new Map<string, PackagePart>([
    [
      "[Content_Types].xml",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        defaults
          .map(([extension, type]) => `<Default Extension="${extension}" ContentType="${type}"/>`)
          .join("") +
        overrides.join("") +
        "</Types>",
    ],
    [
      "_rels/.rels",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
        "</Relationships>",
    ],
    [
      "word/_rels/document.xml.rels",
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        relationships.join("") +
        "</Relationships>",
    ],
    ["word/styles.xml", STYLES_XML],
    [
      "word/document.xml",
      `${DOCUMENT_OPEN}${wellFormedBody(build.body)}${SECTION_PROPERTIES}</w:body></w:document>`,
    ],
    ...(build.parts ?? []),
  ]);
  return parts;
};
