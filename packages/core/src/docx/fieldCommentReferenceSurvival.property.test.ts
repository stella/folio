import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import type {
  Document,
  InlineWrapper,
  ParagraphContent,
  SimpleField,
} from "../types/document";
import { getParagraphText } from "./paragraphParser";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";
import {
  getAttribute,
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  parseXmlDocument,
  WORDPROCESSINGML_NAMESPACE_URIS,
  type XmlElement,
} from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WRAPPERS = {
  dir: (content: ParagraphContent[]) =>
    ({
      type: "inlineWrapper",
      kind: "bidi",
      control: "embedding",
      content,
    }) as const satisfies InlineWrapper,
  bdo: (content: ParagraphContent[]) =>
    ({
      type: "inlineWrapper",
      kind: "bidi",
      control: "override",
      content,
    }) as const satisfies InlineWrapper,
  smartTag: (content: ParagraphContent[]) =>
    ({
      type: "inlineWrapper",
      kind: "smartTag",
      element: "clause",
      propertiesXml:
        '<w:smartTagPr><w:attr w:name="key" w:val="value"/></w:smartTagPr>',
      content,
    }) as const satisfies InlineWrapper,
  customXml: (content: ParagraphContent[]) =>
    ({
      type: "inlineWrapper",
      kind: "customXml",
      element: "clause",
      propertiesXml:
        '<w:customXmlPr><w:attr w:name="key" w:val="value"/></w:customXmlPr>',
      content,
    }) as const satisfies InlineWrapper,
};

const documentOf = (content: ParagraphContent[]) =>
  ({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000001", content }],
        comments: [
          {
            id: 7,
            author: "Reviewer",
            content: [
              {
                type: "paragraph",
                paraId: "7FFFFFF0",
                content: [
                  { type: "run", content: [{ type: "text", text: "Comment" }] },
                ],
              },
            ],
          },
        ],
      },
    },
  }) satisfies Document;

const xmlOf = async (buffer: ArrayBuffer) => {
  const xml = await (
    await ZipArchive.read(buffer)
  )
    .entry("word/document.xml")
    ?.text();
  if (xml === undefined) panic("The comment fixture lost its document part");
  return xml;
};

/** Read the actual authored marker positions, independently of model admission. */
const authoredFacts = (xml: string) => {
  const root = parseXmlDocument(xml);
  if (root === null) panic("The comment fixture is not XML");
  const paths: string[] = [];
  const properties: string[] = [];
  const scopedMarkup: string[] = [];
  const visit = (element: XmlElement, ancestors: string[]): void => {
    const path = [...ancestors, getLocalName(element.name)];
    if (
      WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "") &&
      getLocalName(element.name) === "commentReference"
    )
      paths.push(`${path.join("/")}#${getAttribute(element, "w", "id")}`);
    if (
      WORDPROCESSINGML_NAMESPACE_URIS.has(element.namespaceUri ?? "") &&
      ["smartTagPr", "customXmlPr", "attr"].includes(getLocalName(element.name))
    )
      properties.push(
        `${path.join("/")}#${getAttribute(element, "w", "name")}#${getAttribute(element, "w", "val")}`,
      );
    if (element.namespaceUri === "urn:example:field-markup")
      scopedMarkup.push(
        `${path.join("/")}#${getAttributeByNamespaceUri(element, new Set(["urn:example:field-markup"]), "value")}`,
      );
    for (const child of getChildElements(element)) visit(child, path);
  };
  visit(root, []);
  return { references: paths, properties, scopedMarkup };
};

const assertSurvives = async (buffer: ArrayBuffer): Promise<void> => {
  const expectedFacts = authoredFacts(await xmlOf(buffer));
  expect(expectedFacts.references).toHaveLength(1);
  let document = await parseDocx(buffer, {
    preloadFonts: false,
    detectVariables: false,
  });
  const expectedComments = document.package.document.comments;
  const texts = (source: Document) =>
    source.package.document.content.flatMap((block) =>
      block.type === "paragraph" ? [getParagraphText(block)] : [],
    );
  const expectedText = texts(document);
  expect(expectedComments?.map(({ id }) => id)).toEqual([7]);
  for (let round = 0; round < 2; round += 1) {
    // Rebuild from the parsed model, without an original package to replay.
    const saved = await createDocx({ package: document.package });
    expect(authoredFacts(await xmlOf(saved))).toEqual(expectedFacts);
    document = await parseDocx(saved, {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(document.package.document.comments).toEqual(expectedComments);
    expect(texts(document)).toEqual(expectedText);
  }
};

for (const [name, wrap] of Object.entries(WRAPPERS)) {
  for (const carrier of ["paragraph", "field"] as const) {
    test(`${name} point reference inside ${carrier} survives actual DOCX writes`, async () => {
      const wrapped = wrap([{ type: "commentReference", id: 7 }]);
      const content =
        carrier === "paragraph"
          ? [wrapped]
          : [
              {
                type: "simpleField",
                instruction: " REF target ",
                fieldType: "REF",
                content: [wrapped],
              } satisfies SimpleField,
            ];
      await assertSurvives(await createDocx(documentOf(content)));
    });
  }
}

test(
  "nested field wrappers preserve point references through repeated DOCX parsing",
  async () => {
    await assertProperty(
      fc.asyncProperty(
        fc.array(fc.constantFrom(...Object.values(WRAPPERS)), {
          minLength: 1,
          maxLength: 4,
        }),
        async (wrappers) => {
          let content: ParagraphContent[] = [
            { type: "commentReference", id: 7 },
          ];
          for (const wrap of wrappers) content = [wrap(content)];
          const outer = content.at(0);
          if (outer?.type !== "inlineWrapper")
            panic("The generated field requires a wrapper");
          await assertSurvives(
            await createDocx(
              documentOf([
                {
                  type: "simpleField",
                  instruction: " REF target ",
                  fieldType: "REF",
                  content: [outer],
                },
              ]),
            ),
          );
        },
      ),
      {
        id: "nested field wrappers preserve point references through repeated DOCX parsing",
        numRuns: 40,
      },
    );
  },
  propertyTestTimeout(15_000),
);

test.each(["", "Visible\ttext"])(
  "a direct field point-reference run retains raw markup and text %j",
  async (text) => {
    const source = await createDocx(
      documentOf([{ type: "commentReference", id: 7 }]),
    );
    const zip = await ZipArchive.read(source);
    zip.set(
      "word/document.xml",
      `<w:document xmlns:w="${W}"><w:body><w:p><w:fldSimple w:instr=" REF target "><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:t>${text.replace("\t", "</w:t><w:tab/><w:t>")}</w:t><w:commentReference w:id="7"/></w:r></w:fldSimple></w:p><w:sectPr/></w:body></w:document>`,
    );
    const bytes = await zip.generate({ compression: "store" });
    const parsed = await parseDocx(bytes.buffer, {
      preloadFonts: false,
      detectVariables: false,
    });
    const paragraph = parsed.package.document.content.at(0);
    if (paragraph?.type !== "paragraph")
      panic("The field fixture lost its paragraph");
    const field = paragraph.content.at(0);
    if (field?.type !== "simpleField")
      panic("The field fixture lost its field");
    const capture = field.content.at(0);
    expect(capture?.type).toBe("preservedInline");
    if (capture?.type !== "preservedInline")
      panic("The unsupported direct field run was not preserved");
    expect(capture.text).toBe(text);
    expect(capture.xml).toContain('<w:rStyle w:val="CommentReference"/>');
    await assertSurvives(bytes.buffer);
  },
);

test("a field wrapper keeps inherited namespace bindings on captured children", async () => {
  const zip = await ZipArchive.read(
    await createDocx(documentOf([{ type: "commentReference", id: 7 }])),
  );
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}"><w:body><w:p>
    <f:fldSimple xmlns:f="${W}" xmlns:q="urn:example:field-markup" f:instr=" REF target ">
      <f:smartTag f:element="clause"><f:smartTagPr><f:attr f:name="key" f:val="value"/></f:smartTagPr>
        <q:marker q:value="kept"/>
        <f:r><f:commentReference f:id="7"/></f:r>
      </f:smartTag>
    </f:fldSimple>
  </w:p><w:sectPr/></w:body></w:document>`,
  );
  const bytes = await zip.generate({ compression: "store" });
  expect(authoredFacts(await xmlOf(bytes.buffer)).scopedMarkup).toHaveLength(1);
  await assertSurvives(bytes.buffer);
});
