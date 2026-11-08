/**
 * The block kinds and list labels a snapshot reads do not depend on whether
 * the package's numbering is at hand: the editor's own snapshot (the document
 * alone) and the reviewer's (with the numbering) agree for every paragraph,
 * tracked mark and pending property change included. A paragraph whose mark
 * is struck counts in its original formatting in both.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { FolioDocxReviewer } from "./headless";
import { createFolioAIEditSnapshot } from "./snapshot";

setDefaultTimeout(propertyTestTimeout(120_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const numPr = (numId: number, ilvl: number) =>
  `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;

const paragraphArbitrary = fc.record({
  current: fc.constantFrom(null, 1, 2),
  level: fc.integer({ min: 0, max: 1 }),
  previous: fc.constantFrom("absent", "none", 1, 2),
  mark: fc.constantFrom("none", "ins", "del"),
});
type Generated = fc.RecordValue<typeof paragraphArbitrary>;

const paragraphXml = (generated: Generated, index: number) => {
  const id = 100 + index * 3;
  const change =
    generated.previous === "absent"
      ? ""
      : `<w:pPrChange w:id="${id}" w:author="A" w:date="2026-01-01T00:00:00Z"><w:pPr>${
          generated.previous === "none" ? "" : numPr(generated.previous, 0)
        }</w:pPr></w:pPrChange>`;
  const mark =
    generated.mark === "none"
      ? ""
      : `<w:rPr><w:${generated.mark} w:id="${id + 1}" w:author="A" w:date="2026-01-01T00:00:00Z"/></w:rPr>`;
  return (
    `<w:p w14:paraId="${(index + 1).toString(16).padStart(8, "0").toUpperCase()}"><w:pPr>` +
    (generated.current === null
      ? ""
      : numPr(generated.current, generated.level)) +
    `${mark}${change}</w:pPr><w:r><w:t>Item ${index}</w:t></w:r></w:p>`
  );
};

const level = (ilvl: number, format: string, text: string) =>
  `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${format}"/><w:lvlText w:val="${text}"/></w:lvl>`;

const packageOf = async (paragraphs: readonly Generated[]) => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>` +
      `</Types>`,
  );
  zip.set(
    "_rels/.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.set(
    "word/_rels/document.xml.rels",
    `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.numbering}" Target="numbering.xml"/></Relationships>`,
  );
  zip.set(
    "word/numbering.xml",
    `${HEAD}<w:numbering xmlns:w="${W}">` +
      `<w:abstractNum w:abstractNumId="1">${level(0, "bullet", "•")}${level(1, "bullet", "o")}</w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="2">${level(0, "decimal", "%1.")}${level(1, "lowerLetter", "%2)")}</w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>` +
      `<w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>` +
      `</w:numbering>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      paragraphs.map(paragraphXml).join("") +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const readable = (blocks: readonly { kind: string; text: string }[]) =>
  blocks.map((block) => ({
    kind: block.kind,
    text: block.text,
    label: (block as { displayLabel?: string }).displayLabel,
  }));

describe("list snapshots", () => {
  test("read the same kinds and labels with or without the numbering", async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(paragraphArbitrary, { minLength: 1, maxLength: 5 }),
        async (paragraphs) => {
          const reviewer = await FolioDocxReviewer.fromBuffer(
            await packageOf(paragraphs),
          );
          const withNumbering = readable(reviewer.snapshot().blocks);
          const documentAlone = readable(
            createFolioAIEditSnapshot(
              (
                reviewer as unknown as {
                  state: {
                    doc: Parameters<typeof createFolioAIEditSnapshot>[0];
                  };
                }
              ).state.doc,
            ).blocks,
          );
          expect(documentAlone).toEqual(withNumbering);
        },
      ),
      {
        id: "list snapshots agree with or without the numbering",
        numRuns: 150,
      },
    );
  });
});
