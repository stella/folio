/**
 * In a paragraph with a style, the paragraph mark's run properties lend the
 * paragraph's runs nothing. A save that rewrites such a paragraph, after an
 * edit, an accept or a reject through the document API, or after typing in
 * the editor, writes each run's properties as the run stated them: a property
 * the run shares with the mark is kept, and nothing is turned off.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";
import { repackDocx } from "./rezip";

setDefaultTimeout(propertyTestTimeout(120_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Run properties in `CT_RPr` order. */
const PROPERTIES = {
  b: "<w:b/>",
  i: "<w:i/>",
  color: '<w:color w:val="C00000"/>',
  position: '<w:position w:val="-2"/>',
  sz: '<w:sz w:val="40"/>',
  u: '<w:u w:val="single"/>',
} as const;
type Property = keyof typeof PROPERTIES;
const NAMES = Object.keys(PROPERTIES) as Property[];

const properties = fc.subarray(NAMES);
const rPrOf = (styleId: string | null, set: readonly Property[]) =>
  (styleId === null ? "" : `<w:rStyle w:val="${styleId}"/>`) +
  NAMES.filter((name) => set.includes(name))
    .map((name) => PROPERTIES[name])
    .join("");

const caseArbitrary = fc.record({
  paragraphStyle: properties,
  characterStyle: properties,
  mark: properties,
  markHasStyle: fc.boolean(),
  run: properties,
  runHasStyle: fc.boolean(),
  edit: fc.constantFrom("direct", "tracked-reject", "tracked-accept", "typed"),
});
type Case = fc.RecordValue<typeof caseArbitrary>;

const documentOf = async (example: Case): Promise<ArrayBuffer> => {
  const runProperties = rPrOf(
    example.runHasStyle ? "Strong" : null,
    example.run,
  );
  const rPr = runProperties === "" ? "" : `<w:rPr>${runProperties}</w:rPr>`;
  const run = (text: string) =>
    `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const deletion =
    example.edit === "tracked-reject" || example.edit === "tracked-accept"
      ? `<w:del w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r>${rPr}<w:delText>gone </w:delText></w:r></w:del>`
      : "";
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
  zip.set(
    "word/styles.xml",
    `${HEAD}<w:styles xmlns:w="${W}">` +
      `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
      `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/><w:rPr>${rPrOf(null, example.paragraphStyle)}</w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="Strong"><w:name w:val="Strong"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr>${rPrOf(null, example.characterStyle)}</w:rPr></w:style>` +
      `</w:styles>`,
  );
  zip.set(
    "word/document.xml",
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>` +
      `<w:p w14:paraId="22222222"><w:pPr><w:pStyle w:val="Body"/>` +
      `<w:rPr>${rPrOf(example.markHasStyle ? "Strong" : null, example.mark)}</w:rPr></w:pPr>` +
      `${run("Kept words ")}${deletion}${run("edited")}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

/** Each run's text and the run properties its `w:rPr` states. */
const runsOf = async (bytes: ArrayBuffer) => {
  const xml =
    (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
    "";
  const start = xml.indexOf('w14:paraId="22222222"');
  const paragraph = xml.slice(start, xml.indexOf("</w:p>", start));
  const body = paragraph.slice(paragraph.indexOf("</w:pPr>"));
  return [
    ...body.matchAll(
      /<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t[^>]*>([^<]*)<\/w:t><\/w:r>/gu,
    ),
  ].map(([, stated = "", text = ""]) => ({ text, stated }));
};

/** Type at the end of the paragraph in the editor and save the document. */
const typedAndSaved = async (source: ArrayBuffer): Promise<ArrayBuffer> => {
  const parsed = await parseDocx(source, {
    preloadFonts: false,
    detectVariables: false,
  });
  let state = EditorState.create({
    schema,
    doc: toProseDoc(parsed, { styles: parsed.package.styles }),
    plugins: [...singletonManager.getPlugins()],
  });
  const end = (state.doc.firstChild?.nodeSize ?? 2) - 1;
  state = state.apply(
    state.tr.setSelection(TextSelection.create(state.doc, end)),
  );
  state = state.apply(state.tr.insertText("X"));
  return repackDocx(
    fromProseDoc(state.doc, parsed, { stylesheetSource: { type: "package" } }),
  );
};

describe("a paragraph with a style, rewritten on save", () => {
  test("writes its runs as they were stated, whatever the mark states", async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (example) => {
        const source = await documentOf(example);
        let saved: ArrayBuffer;
        if (example.edit === "typed") saved = await typedAndSaved(source);
        else {
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
          saved = await reviewer.toBuffer();
        }
        const stated = (await runsOf(source)).at(0)?.stated ?? "";
        const runs = await runsOf(saved);
        expect(runs.length).toBeGreaterThan(0);
        for (const run of runs) expect(run.stated).toBe(stated);
      }),
      {
        id: "styled paragraph runs keep their stated properties on save",
        numRuns: 200,
      },
    );
  });
});
