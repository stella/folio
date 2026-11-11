/**
 * A paragraph rewritten on save keeps its runs' toggle properties as they were
 * stated, whatever the paragraph mark, the paragraph style and the runs'
 * character styles set.
 *
 * A run's character style that set a toggle property (bold, italic, caps…)
 * the paragraph mark's character style also set read as a run that turned the
 * property off: any edit in the paragraph, or accepting or rejecting a change
 * in it, then wrote `w:val="0"` for it on every such run.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { RELATIONSHIP_TYPES } from "./relsParser";

setDefaultTimeout(propertyTestTimeout(120_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/**
 * Toggle properties, in the order `CT_RPr` lists them. `dstrike` is left out:
 * it shares its effect with `strike`, which this property does not cover.
 */
const TOGGLES = [
  "b",
  "i",
  "caps",
  "smallCaps",
  "strike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
] as const;
type Toggle = (typeof TOGGLES)[number];

const toggles = fc.subarray([...TOGGLES], { maxLength: 3 });
const rPrOf = (styleId: string | null, set: readonly Toggle[]) =>
  (styleId === null ? "" : `<w:rStyle w:val="${styleId}"/>`) +
  TOGGLES.filter((toggle) => set.includes(toggle))
    .map((toggle) => `<w:${toggle}/>`)
    .join("");

const caseArbitrary = fc.record({
  paragraphStyle: toggles,
  markStyle: toggles,
  markDirect: toggles,
  markHasStyle: fc.boolean(),
  runStyle: toggles,
  runDirect: toggles,
  runHasStyle: fc.boolean(),
  hasParagraphStyle: fc.boolean(),
  edit: fc.constantFrom("direct", "tracked-reject", "tracked-accept"),
});

const documentOf = async (
  example: fc.RecordValue<typeof caseArbitrary>,
): Promise<ArrayBuffer> => {
  const styles =
    `${HEAD}<w:styles xmlns:w="${W}">` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/><w:rPr>${rPrOf(null, example.paragraphStyle)}</w:rPr></w:style>` +
    `<w:style w:type="character" w:styleId="MarkStyle"><w:name w:val="Mark Style"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr>${rPrOf(null, example.markStyle)}</w:rPr></w:style>` +
    `<w:style w:type="character" w:styleId="RunStyle"><w:name w:val="Run Style"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr>${rPrOf(null, example.runStyle)}</w:rPr></w:style>` +
    `</w:styles>`;
  const runProperties = rPrOf(
    example.runHasStyle ? "RunStyle" : null,
    example.runDirect,
  );
  const run = (text: string) =>
    `<w:r>${runProperties === "" ? "" : `<w:rPr>${runProperties}</w:rPr>`}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const deletion =
    example.edit === "direct"
      ? ""
      : `<w:del w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r>${runProperties === "" ? "" : `<w:rPr>${runProperties}</w:rPr>`}<w:delText>gone </w:delText></w:r></w:del>`;
  const paragraph =
    `<w:p w14:paraId="22222222"><w:pPr>` +
    (example.hasParagraphStyle ? '<w:pStyle w:val="Body"/>' : "") +
    `<w:rPr>${rPrOf(example.markHasStyle ? "MarkStyle" : null, example.markDirect)}</w:rPr></w:pPr>` +
    `${run("Kept words ")}${deletion}${run("edited")}</w:p>`;
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
      `<Default Extension="xml" ContentType="application/xml"/>` +
      `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
      `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
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
      `<Relationship Id="rId2" Type="${RELATIONSHIP_TYPES.styles}" Target="styles.xml"/></Relationships>`,
  );
  zip.set("word/styles.xml", styles);
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${paragraph}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

/** Each run's text with the toggle elements and style its `w:rPr` states. */
const runsOf = async (bytes: ArrayBuffer) => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  const start = xml.indexOf('w14:paraId="22222222"');
  const paragraph = xml.slice(start, xml.indexOf("</w:p>", start));
  return [
    ...paragraph.matchAll(
      /<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t[^>]*>([^<]*)<\/w:t><\/w:r>/gu,
    ),
  ].map(([, properties = "", text = ""]) => ({
    text,
    stated: [
      ...properties.matchAll(/<w:(rStyle|\w+)(?: w:val="([^"]*)")?\/>/gu),
    ]
      .map(([, name, value]) =>
        name === "rStyle" ? `style:${value ?? ""}` : `${name}=${value ?? "1"}`,
      )
      .filter(
        (entry) =>
          entry.startsWith("style:") ||
          TOGGLES.some((toggle) => entry.startsWith(`${toggle}=`)),
      )
      .toSorted(),
  }));
};

describe("toggle properties of a paragraph written from the model", () => {
  test("stay as the runs stated them", async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (example) => {
        const source = await documentOf(example);
        const reviewer = await FolioDocxReviewer.fromBuffer(source);
        if (example.edit === "direct") {
          const result = reviewer.applyOperations(
            [
              {
                id: "edit",
                type: "replaceInBlock",
                blockId: "22222222",
                find: "edited",
                replace: "changed",
              },
            ],
            { undefinedReferences: "refuse", mode: "direct" },
          );
          expect(result.applied).toHaveLength(1);
        } else if (example.edit === "tracked-reject") reviewer.rejectAll();
        else reviewer.acceptAll();
        const before = await runsOf(source);
        const statedBefore = before.at(0)?.stated;
        const after = await runsOf(await reviewer.toBuffer());
        // Every run the save writes states what the source's runs stated:
        // merged runs share it, nothing is turned on or off.
        for (const run of after) expect(run.stated).toEqual(statedBefore);
      }),
      {
        id: "paragraph mark toggles leave run toggles as stated",
        numRuns: 200,
      },
    );
  });
});
