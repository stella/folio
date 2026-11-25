/**
 * A paragraph an operation numbers has, in the canonical reviewer's live
 * model, the list rendering the saved file reopens to.
 *
 * List rendering is derived from the paragraph's numbering and the numbering
 * definitions. Each case sets, clears, changes the level of, inherits or
 * starts a new list on a paragraph through `setBlockParagraphProperties`,
 * directly or as a tracked change, on a paragraph of the body, of a table
 * cell (nested tables and merged cells too), of a content control, of a
 * header or of a note, then saves and reopens: every paragraph's numbering
 * and list rendering read the same in the live model and in the reopened one,
 * and so does the indentation a numbering level adds to the paragraph's
 * formatting.
 */

import { describe, expect, mock, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import * as canonicalListRendering from "../canonical/canonicalListRendering";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { getParagraphText } from "../docx/paragraphParser";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { type FolioDocumentStoryHandle, FolioDocxReviewer } from "./headless";
import type { FolioAIListNumbering } from "./types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml";

const paragraphXml = (text: string, properties = ""): string =>
  `<w:p>${properties === "" ? "" : `<w:pPr>${properties}</w:pPr>`}<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const numbered = (numId: number, ilvl: number): string =>
  `<w:numPr><w:ilvl w:val="${String(ilvl)}"/><w:numId w:val="${String(numId)}"/></w:numPr>`;

const cell = (content: string, properties = ""): string =>
  `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>${properties}</w:tcPr>${content}</w:tc>`;

const table = (rows: readonly string[]): string =>
  '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
  rows.map((row) => `<w:tr>${row}</w:tr>`).join("") +
  "</w:tbl>";

const control = (id: number, content: string): string =>
  `<w:sdt><w:sdtPr><w:id w:val="${String(id)}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;

/**
 * Plain and numbered paragraphs in the body, in table cells (one spanning two
 * columns, two merged down a column, one holding a nested table and one a
 * content control) and in a content control of the body.
 */
const BODY =
  paragraphXml("Body plain") +
  paragraphXml("Body numbered", numbered(5, 0)) +
  table([
    cell(paragraphXml("Cell plain")) +
      cell(paragraphXml("Cell numbered", numbered(5, 0))),
    cell(
      paragraphXml("Spanning numbered", numbered(5, 0)),
      '<w:gridSpan w:val="2"/>',
    ),
    cell(
      paragraphXml("Merged numbered", numbered(5, 0)),
      '<w:vMerge w:val="restart"/>',
    ) +
      cell(
        table([
          cell(paragraphXml("Nested numbered", numbered(5, 0))) +
            cell(paragraphXml("Nested plain")),
        ]) + paragraphXml("After nested"),
      ),
    cell("<w:p/>", "<w:vMerge/>") +
      cell(control(8, paragraphXml("Cell control numbered", numbered(5, 0)))),
  ]) +
  control(7, paragraphXml("Control plain")) +
  paragraphXml("Closing numbered", numbered(5, 0)) +
  `<w:p><w:r><w:t xml:space="preserve">Noted</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p>`;

const HEADER =
  paragraphXml("Header numbered", numbered(5, 0)) +
  table([
    cell(paragraphXml("Header cell plain")) +
      cell(paragraphXml("Header cell numbered", numbered(5, 0))),
  ]) +
  paragraphXml("Header plain");

const NOTES =
  '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
  '<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Note plain</w:t></w:r></w:p>' +
  paragraphXml("Note numbered", numbered(5, 0)) +
  "</w:footnote>";

const MAIN_STORY: FolioDocumentStoryHandle = { type: "main" };
const HEADER_STORY: FolioDocumentStoryHandle = {
  type: "header",
  relationshipId: "rIdH",
};
const NOTE_STORY: FolioDocumentStoryHandle = { type: "footnote", noteId: 1 };

const TARGETS = [
  { story: MAIN_STORY, text: "Body plain" },
  { story: MAIN_STORY, text: "Body numbered" },
  { story: MAIN_STORY, text: "Cell plain" },
  { story: MAIN_STORY, text: "Cell numbered" },
  { story: MAIN_STORY, text: "Spanning numbered" },
  { story: MAIN_STORY, text: "Merged numbered" },
  { story: MAIN_STORY, text: "Nested plain" },
  { story: MAIN_STORY, text: "Nested numbered" },
  { story: MAIN_STORY, text: "Cell control numbered" },
  { story: MAIN_STORY, text: "Control plain" },
  { story: HEADER_STORY, text: "Header cell plain" },
  { story: HEADER_STORY, text: "Header numbered" },
  { story: NOTE_STORY, text: "Note plain" },
  { story: NOTE_STORY, text: "Note numbered" },
] as const;

const level = (ilvl: number, format: string, text: string): string =>
  `<w:lvl w:ilvl="${String(ilvl)}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${String(720 * (ilvl + 1))}" w:hanging="360"/></w:pPr></w:lvl>`;

const NUMBERING =
  `${HEAD}<w:numbering xmlns:w="${W}">` +
  `<w:abstractNum w:abstractNumId="5">${level(0, "decimal", "%1.")}${level(1, "lowerLetter", "%2)")}</w:abstractNum>` +
  `<w:abstractNum w:abstractNumId="6">${level(0, "bullet", "o")}${level(1, "bullet", "-")}</w:abstractNum>` +
  '<w:num w:numId="5"><w:abstractNumId w:val="5"/></w:num>' +
  '<w:num w:numId="6"><w:abstractNumId w:val="6"/></w:num>' +
  "</w:numbering>";

const fixture = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${MAIN}.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="${MAIN}.numbering+xml"/><Override PartName="/word/header1.xml" ContentType="${MAIN}.header+xml"/><Override PartName="/word/footnotes.xml" ContentType="${MAIN}.footnotes+xml"/></Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdN" Type="${R}/numbering" Target="numbering.xml"/><Relationship Id="rIdH" Type="${R}/header" Target="header1.xml"/><Relationship Id="rIdF" Type="${R}/footnotes" Target="footnotes.xml"/></Relationships>`,
  );
  zip.set("word/numbering.xml", NUMBERING);
  zip.set(
    "word/header1.xml",
    `${HEAD}<w:hdr xmlns:w="${W}" xmlns:r="${R}">${HEADER}</w:hdr>`,
  );
  zip.set(
    "word/footnotes.xml",
    `${HEAD}<w:footnotes xmlns:w="${W}" xmlns:r="${R}">${NOTES}</w:footnotes>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${BODY}<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/></w:sectPr></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  ) as ArrayBuffer;
};

const open = (
  buffer: ArrayBuffer,
  session: HeadlessReviewerSession = "canonical",
) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), {
    author: "Reviewer",
    session,
  });

const ilvl = fc.integer({ min: 0, max: 1 });

/** Set, clear, change the level of, inherit, or start a new list. */
const numberingArbitrary: fc.Arbitrary<FolioAIListNumbering> = fc.oneof(
  fc.record({
    kind: fc.constant("reference" as const),
    numId: fc.constantFrom(5, 6),
    ilvl,
  }),
  fc.constant({ kind: "none" as const }),
  fc.record({ kind: fc.constant("levelOnly" as const), ilvl }),
  fc.constant({ kind: "inherit" as const }),
  fc.record({
    kind: fc.constant("newList" as const),
    format: fc.constantFrom("decimal" as const, "bullet" as const),
    level: ilvl,
  }),
);

type Draw = {
  target: (typeof TARGETS)[number];
  mode: "direct" | "tracked-changes";
  numbering: FolioAIListNumbering;
};

const drawArbitrary: fc.Arbitrary<Draw> = fc.record({
  target: fc.constantFrom(...TARGETS),
  mode: fc.constantFrom("direct", "tracked-changes"),
  numbering: numberingArbitrary,
});

const asJson = (value: unknown): unknown =>
  value === undefined ? null : (JSON.parse(JSON.stringify(value)) as unknown);

const paragraphRead = (paragraph: Paragraph) => ({
  numPr: asJson(paragraph.formatting?.numPr),
  numPrFromStyle: asJson(paragraph.formatting?.numPrFromStyle),
  listRendering: asJson(paragraph.listRendering),
  // The indentation a numbering level adds to the paragraph's formatting.
  indentation: asJson({
    left: paragraph.formatting?.indentLeft,
    right: paragraph.formatting?.indentRight,
    firstLine: paragraph.formatting?.indentFirstLine,
    hanging: paragraph.formatting?.hangingIndent,
  }),
});

/**
 * Every paragraph of the body, the header and the note, in tables and content
 * controls too, by its story and place.
 */
const paragraphsOf = (document: Document) => {
  const read: Record<string, ReturnType<typeof paragraphRead>> = {};
  const visit = (story: string, blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        read[`${story} #${String(Object.keys(read).length)}`] =
          paragraphRead(block);
      } else if (block.type === "table") {
        for (const row of block.rows)
          for (const tableCell of row.cells) visit(story, tableCell.content);
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml")
        visit(story, block.content);
    }
  };
  visit("body", document.package.document.content);
  for (const [id, header] of document.package.headers ?? [])
    visit(`header ${id}`, header.content);
  for (const note of document.package.footnotes ?? [])
    visit(`note ${String(note.id)}`, note.content);
  return read;
};

const operationOf = (draw: Draw, blockId: string): FolioDocumentOperation => ({
  id: "number",
  type: "setBlockParagraphProperties",
  blockId,
  properties: { numbering: draw.numbering },
});

/** Apply the draw; `null` when the reviewer refuses it. */
const liveAndReopened = async (buffer: ArrayBuffer, draw: Draw) => {
  const live = await open(buffer);
  const { story, text } = draw.target;
  const block = live
    .snapshotStory(story)
    ?.blocks.find((candidate) => candidate.text.trim() === text);
  if (block === undefined) throw new Error(`no block reads ${text}`);
  let applied = false;
  try {
    applied =
      live.applyDocumentOperationsToStory({
        story,
        batch: {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          operations: [operationOf(draw, block.id)],
          mode: draw.mode,
        },
        undefinedReferences: "refuse",
      }).applied.length === 1;
  } catch {
    applied = false;
  }
  if (!applied) return null;
  const saved = await live.toBuffer();
  const reopened = await open(saved);
  return {
    live: paragraphsOf(live.toDocument()),
    reopened: paragraphsOf(reopened.toDocument()),
  };
};

const law = (buffer: ArrayBuffer, tally: { applied: number }) =>
  fc.asyncProperty(drawArbitrary, async (draw) => {
    const read = await liveAndReopened(buffer, draw);
    if (read === null) return;
    tally.applied += 1;
    expect(read.live).toEqual(read.reopened);
  });

const RUNS = 40;

/** A tracked reference that numbers a table cell's paragraph. */
const CELL_REFERENCE: Draw = {
  target: { story: MAIN_STORY, text: "Cell plain" },
  mode: "tracked-changes",
  numbering: { kind: "reference", numId: 5, ilvl: 0 },
};

/** The marker of every numbered paragraph of `document`, by its text. */
const markersOf = (document: Document): Record<string, string> => {
  const markers: Record<string, string> = {};
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") {
        const marker = block.listRendering?.marker;
        if (marker !== undefined)
          markers[getParagraphText(block).trim()] = marker;
      } else if (block.type === "table") {
        for (const row of block.rows)
          for (const tableCell of row.cells) visit(tableCell.content);
      } else if (block.type === "blockSdt" || block.type === "blockCustomXml")
        visit(block.content);
    }
  };
  visit(document.package.document.content);
  for (const header of document.package.headers?.values() ?? [])
    visit(header.content);
  for (const note of document.package.footnotes ?? []) visit(note.content);
  return markers;
};

describe("a list's numbering", () => {
  test("continues through table cells in document order, in every story", async () => {
    const buffer = await fixture();
    // Row by row, cell by cell, through a spanning cell, merged cells, a
    // nested table and a content control in a cell; each story counts its own.
    const expected = {
      "Body numbered": "1.",
      "Cell numbered": "2.",
      "Spanning numbered": "3.",
      "Merged numbered": "4.",
      "Nested numbered": "5.",
      "Cell control numbered": "6.",
      "Closing numbered": "7.",
      "Header numbered": "1.",
      "Header cell numbered": "2.",
      "Note numbered": "1.",
    };
    for (const session of ["canonical"] as const)
      expect(markersOf((await open(buffer, session)).toDocument())).toEqual(
        expected,
      );
  });
});

describe("numbering an operation sets", () => {
  test(
    "renders the same live and after save and reopen",
    async () => {
      const buffer = await fixture();
      await assertProperty(law(buffer, { applied: 0 }), {
        id: "list rendering reads the same live and reopened",
        numRuns: RUNS,
        examples: [[CELL_REFERENCE]],
      });
      // Most draws of a fixed sample reach the law: the reviewer applies them.
      const tally = { applied: 0 };
      await checkCoverageSample(law(buffer, tally), RUNS);
      expect(tally.applied).toBeGreaterThan(RUNS / 2);
    },
    propertyTestTimeout(120_000),
  );

  test("tracked list references restore authored indentation on rejection", async () => {
    const buffer = await fixture();
    for (const target of [
      { story: MAIN_STORY, text: "Body numbered" },
      { story: NOTE_STORY, text: "Note numbered" },
    ]) {
      const read = await liveAndReopened(buffer, {
        target,
        mode: "tracked-changes",
        numbering: { kind: "reference", numId: 6, ilvl: 1 },
      });
      expect(read).not.toBeNull();
      if (read === null) continue;
      expect(read.live).toEqual(read.reopened);
    }
  });

  test(
    "the law catches an edit's normalization that leaves a table's paragraphs as they were",
    async () => {
      const buffer = await fixture();
      const draw = { ...CELL_REFERENCE, mode: "direct" } as const;
      const baseline = await liveAndReopened(buffer, draw);
      expect(baseline).not.toBeNull();
      expect(baseline?.live).toEqual(baseline?.reopened);

      // The test-only break: the normalization after an edit stops at the
      // story's own paragraphs and leaves its tables as the edit left them.
      // Reopening reads the tables through the parser, which numbers them.
      const original = { ...canonicalListRendering };
      mock.module("../canonical/canonicalListRendering", () => ({
        ...original,
        normalizeCanonicalListRendering: (
          ...args: Parameters<typeof original.normalizeCanonicalListRendering>
        ) => {
          const [document] = args;
          const content = document.package.document.content;
          const result = original.normalizeCanonicalListRendering(...args);
          const normalizedContent = result.document.package.document.content;
          if (normalizedContent.length !== content.length) return result;
          return {
            ...result,
            document: {
              ...result.document,
              package: {
                ...result.document.package,
                document: {
                  ...result.document.package.document,
                  content: normalizedContent.map((block, index) => {
                    const before = content[index];
                    return block.type === "table" && before?.type === "table"
                      ? before
                      : block;
                  }),
                },
              },
            },
          };
        },
      }));
      try {
        // The mutant also makes the tracked transaction unreviewable. Apply
        // directly so the normalization law sees the stale table.
        const read = await liveAndReopened(buffer, draw);
        expect(read).not.toBeNull();
        expect(read?.live).not.toEqual(read?.reopened);
      } finally {
        mock.module("../canonical/canonicalListRendering", () => original);
      }
    },
    propertyTestTimeout(60_000),
  );
});
