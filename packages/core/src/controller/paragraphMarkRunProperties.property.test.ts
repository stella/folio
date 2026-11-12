/**
 * A paragraph mark's run properties (`w:pPr/w:rPr`) format the glyph of the
 * paragraph mark alone (ECMA-376 Part 1, 17.3.1.29): no run of the paragraph
 * takes them, whatever it states. Both engines read a run the same way, an
 * unedited save writes the package back unchanged, and an edit, accept or
 * reject in the paragraph writes each run as it was stated, with no property
 * turned off for the mark's sake. Text typed into an empty paragraph takes
 * the mark's formatting as its own.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, TextSelection } from "prosemirror-state";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { parseDocx } from "../docx/parser";
import { RELATIONSHIP_TYPES } from "../docx/relsParser";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { schema, singletonManager } from "../prosemirror/schema";
import { createCanonicalSession } from "./canonicalSession";

setDefaultTimeout(propertyTestTimeout(120_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** Run properties in `CT_RPr` order, each with the mark it reads as. */
const PROPERTIES = {
  b: { xml: "<w:b/>", mark: "bold" },
  i: { xml: "<w:i/>", mark: "italic" },
  color: { xml: '<w:color w:val="C00000"/>', mark: "textColor" },
  sz: { xml: '<w:sz w:val="40"/>', mark: "fontSize" },
  u: { xml: '<w:u w:val="single"/>', mark: "underline" },
} as const;
type Property = keyof typeof PROPERTIES;
const NAMES = Object.keys(PROPERTIES) as Property[];

const properties = fc.subarray(NAMES);
const rPrOf = (styleId: string | null, set: readonly Property[]) =>
  (styleId === null ? "" : `<w:rStyle w:val="${styleId}"/>`) +
  NAMES.filter((name) => set.includes(name))
    .map((name) => PROPERTIES[name].xml)
    .join("");

const caseArbitrary = fc.record({
  mark: properties.filter((set) => set.length > 0),
  paragraphStyle: fc.option(properties, { nil: null }),
  run: fc.option(properties, { nil: null }),
  runStyle: fc.option(properties, { nil: null }),
  edit: fc.constantFrom("none", "direct", "tracked-reject", "tracked-accept"),
});
type Case = fc.RecordValue<typeof caseArbitrary>;

const stylesXml = (example: Case) =>
  `${HEAD}<w:styles xmlns:w="${W}">` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/><w:basedOn w:val="Normal"/><w:rPr>${rPrOf(null, example.paragraphStyle ?? [])}</w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="RunStyle"><w:name w:val="Run Style"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr>${rPrOf(null, example.runStyle ?? [])}</w:rPr></w:style>` +
  `</w:styles>`;

const packageOf = async (styles: string, body: string) => {
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
    `${HEAD}<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  return (await zip.generate({ compression: "store" })).buffer;
};

const documentOf = (example: Case) => {
  const runProperties =
    example.run === null && example.runStyle === null
      ? null
      : rPrOf(example.runStyle === null ? null : "RunStyle", example.run ?? []);
  const rPr = runProperties === null ? "" : `<w:rPr>${runProperties}</w:rPr>`;
  const run = (text: string) =>
    `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;
  const deletion =
    example.edit === "tracked-reject" || example.edit === "tracked-accept"
      ? `<w:del w:id="1" w:author="A" w:date="2024-01-01T00:00:00Z"><w:r>${rPr}<w:delText>gone </w:delText></w:r></w:del>`
      : "";
  return packageOf(
    stylesXml(example),
    `<w:p w14:paraId="22222222"><w:pPr>` +
      (example.paragraphStyle === null ? "" : '<w:pStyle w:val="Body"/>') +
      `<w:rPr>${rPrOf(null, example.mark)}</w:rPr></w:pPr>` +
      `${run("Kept words ")}${deletion}${run("edited")}</w:p>`,
  );
};

const documentXml = async (bytes: ArrayBuffer) =>
  (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
  "";

/** Each run's text and the run properties its `w:rPr` states. */
const runsOf = async (bytes: ArrayBuffer) => {
  const xml = await documentXml(bytes);
  const start = xml.indexOf('w14:paraId="22222222"');
  const paragraph = xml.slice(start, xml.indexOf("</w:p>", start));
  const body = paragraph.slice(paragraph.indexOf("</w:pPr>"));
  return [
    ...body.matchAll(
      /<w:r>(?:<w:rPr>(.*?)<\/w:rPr>)?<w:t[^>]*>([^<]*)<\/w:t><\/w:r>/gu,
    ),
  ].map(([, stated = "", text = ""]) => ({ text, stated }));
};

/** The formatting marks a projected paragraph's text carries, by text. */
const visibleMarks = (doc: PMNode) => {
  const out: string[] = [];
  doc.descendants((node) => {
    if (node.isText)
      out.push(
        `${node.text ?? ""}: ${node.marks
          .map(({ type }) => type.name)
          .filter((name) =>
            NAMES.some((property) => PROPERTIES[property].mark === name),
          )
          .toSorted()
          .join(",")}`,
      );
    return true;
  });
  return out;
};

describe("paragraph mark run properties", () => {
  test("format the mark alone in both engines and on every save", async () => {
    await assertProperty(
      fc.asyncProperty(caseArbitrary, async (example) => {
        const source = await documentOf(example);
        const parsed = await parseDocx(source, {
          preloadFonts: false,
          detectVariables: false,
        });
        const projected = toProseDoc(parsed, {
          styles: parsed.package.styles,
        });
        // What each run shows comes from the run and the styles only.
        const shown = new Set(
          [
            ...(example.paragraphStyle ?? []),
            ...(example.runStyle ?? []),
            ...(example.run ?? []),
          ].map((property) => PROPERTIES[property].mark),
        );
        for (const node of visibleMarks(projected)) {
          const [, names = ""] = node.split(": ");
          for (const name of names.split(",").filter(Boolean))
            expect(shown.has(name as never)).toBe(true);
        }
        const session = createCanonicalSession(parsed).unwrap();
        expect(visibleMarks(session.projection.doc)).toEqual(
          visibleMarks(projected),
        );

        const reviewer = await FolioDocxReviewer.fromBuffer(source);
        if (example.edit === "none") {
          expect(await documentXml(await reviewer.toBuffer())).toBe(
            await documentXml(source),
          );
          return;
        }
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
        const stated = (await runsOf(source)).at(0)?.stated ?? "";
        for (const run of await runsOf(await reviewer.toBuffer()))
          expect(run.stated).toBe(stated);
      }),
      {
        id: "paragraph mark run properties format the mark alone",
        numRuns: 150,
      },
    );
  });

  test("text typed into an empty paragraph takes the mark's formatting as its own", async () => {
    await assertProperty(
      fc.asyncProperty(
        properties.filter((set) => set.length > 0),
        async (mark) => {
          const source = await packageOf(
            stylesXml({
              mark,
              paragraphStyle: null,
              run: null,
              runStyle: null,
              edit: "none",
            }),
            `<w:p w14:paraId="22222222"><w:pPr><w:rPr>${rPrOf(null, mark)}</w:rPr></w:pPr></w:p>`,
          );
          const parsed = await parseDocx(source, {
            preloadFonts: false,
            detectVariables: false,
          });
          let state = EditorState.create({
            schema,
            doc: toProseDoc(parsed, { styles: parsed.package.styles }),
            plugins: [...singletonManager.getPlugins()],
          });
          state = state.apply(
            state.tr.setSelection(TextSelection.create(state.doc, 1)),
          );
          state = state.apply(state.tr.insertText("Typed"));
          const saved = fromProseDoc(state.doc, parsed, {
            stylesheetSource: { type: "package" },
          });
          const paragraph = saved.package.document.content.at(0);
          if (paragraph?.type !== "paragraph")
            throw new Error("the paragraph is gone");
          const run = paragraph.content.find((item) => item.type === "run");
          const formatting = run?.type === "run" ? run.formatting : undefined;
          expect(formatting?.bold ?? false).toBe(mark.includes("b"));
          expect(formatting?.italic ?? false).toBe(mark.includes("i"));
          expect(formatting?.fontSize).toBe(
            mark.includes("sz") ? 40 : undefined,
          );
          expect(formatting?.color?.rgb).toBe(
            mark.includes("color") ? "C00000" : undefined,
          );
          expect(formatting?.underline?.style).toBe(
            mark.includes("u") ? "single" : undefined,
          );
        },
      ),
      {
        id: "typing into an empty paragraph takes the mark's formatting",
        numRuns: 40,
      },
    );
  });
});
