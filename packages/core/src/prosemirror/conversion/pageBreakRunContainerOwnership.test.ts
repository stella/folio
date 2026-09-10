import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  Table,
} from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  footnoteToProseDoc,
  headerFooterToProseDoc,
  UnsupportedDocxToProseMirrorConversionError,
  toProseDoc,
} from "./toProseDoc";

const REVISION = {
  id: 91,
  author: "Reviewer",
  date: "2026-09-09T00:00:00.000Z",
} as const;

type ExpectedUnsupportedConversion = Pick<
  UnsupportedDocxToProseMirrorConversionError,
  "message" | "owner" | "contentType"
>;

const expectUnsupportedConversion = (
  convert: () => unknown,
  expected: ExpectedUnsupportedConversion,
): void => {
  try {
    convert();
  } catch (error) {
    expect(error).toBeInstanceOf(UnsupportedDocxToProseMirrorConversionError);
    if (!(error instanceof UnsupportedDocxToProseMirrorConversionError)) {
      throw error;
    }
    expect(error).toMatchObject(expected);
    return;
  }
  throw new Error("Expected DOCX-to-ProseMirror conversion to be rejected");
};

const pageBreakRun = (content: readonly RunContent[] = []): Run => ({
  type: "run",
  content: [...content, { type: "break", breakType: "page" }],
});

type InlineWrapper = {
  name: string;
  wrap: (run: Run) => ParagraphContent;
};

const INLINE_WRAPPERS = [
  { name: "direct run", wrap: (run) => run },
  {
    name: "hyperlink",
    wrap: (run) => ({ type: "hyperlink", href: "https://example.test", children: [run] }),
  },
  {
    name: "simple field",
    wrap: (run) => ({
      type: "simpleField",
      instruction: "REF target",
      fieldType: "REF",
      content: [run],
    }),
  },
  {
    name: "complex field result",
    wrap: (run) => ({
      type: "complexField",
      instruction: "REF target",
      fieldType: "REF",
      fieldCode: [],
      fieldResult: [run],
    }),
  },
  {
    name: "inline content control",
    wrap: (run) => ({
      type: "inlineSdt",
      properties: { sdtType: "richText" },
      content: [run],
    }),
  },
  ...(["insertion", "deletion", "moveFrom", "moveTo"] as const).map((type) => ({
    name: `${type} wrapper`,
    wrap: (run: Run) => ({ type, info: REVISION, content: [run] }),
  })),
  {
    name: "nested control, tracking, field, and hyperlink wrappers",
    wrap: (run) => ({
      type: "inlineSdt",
      properties: { sdtType: "richText" },
      content: [
        {
          type: "insertion",
          info: REVISION,
          content: [
            {
              type: "simpleField",
              instruction: "REF target",
              fieldType: "REF",
              content: [
                {
                  type: "hyperlink",
                  href: "https://example.test/nested",
                  children: [run],
                },
              ],
            },
          ],
        },
      ],
    }),
  },
] as const satisfies readonly InlineWrapper[];

const tableWithParagraph = (paragraph: Paragraph, header = false): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      ...(header ? { formatting: { header: true } } : {}),
      cells: [{ type: "tableCell", content: [paragraph] }],
    },
  ],
});

const textBoxHost = (content: (Paragraph | Table)[]): Paragraph => ({
  type: "paragraph",
  content: [
    {
      type: "run",
      content: [
        {
          type: "shape",
          shape: {
            type: "shape",
            shapeType: "rect",
            size: { width: 914_400, height: 457_200 },
            textBody: { content },
          },
        },
      ],
    },
  ],
});

const documentWithContent = (content: BlockContent[]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = content;
  return document;
};

type ContainerContext = {
  name: string;
  owner: "table-cell" | "text-box";
  build: (paragraph: Paragraph) => Document;
};

const CONTAINER_CONTEXTS = [
  {
    name: "ordinary table cell",
    owner: "table-cell",
    build: (paragraph) => documentWithContent([tableWithParagraph(paragraph)]),
  },
  {
    name: "repeating-header table cell",
    owner: "table-cell",
    build: (paragraph) => documentWithContent([tableWithParagraph(paragraph, true)]),
  },
  {
    name: "nested table cell",
    owner: "table-cell",
    build: (paragraph) =>
      documentWithContent([
        {
          type: "table",
          rows: [
            {
              type: "tableRow",
              cells: [
                {
                  type: "tableCell",
                  content: [tableWithParagraph(paragraph)],
                },
              ],
            },
          ],
        },
      ]),
  },
  {
    name: "text box",
    owner: "text-box",
    build: (paragraph) => documentWithContent([textBoxHost([paragraph])]),
  },
  {
    name: "table nested in a text box",
    owner: "text-box",
    build: (paragraph) => documentWithContent([textBoxHost([tableWithParagraph(paragraph)])]),
  },
  {
    name: "text box nested in a table cell",
    owner: "table-cell",
    build: (paragraph) => documentWithContent([tableWithParagraph(textBoxHost([paragraph]))]),
  },
] as const satisfies readonly ContainerContext[];

const CONTAINER_MESSAGES = {
  "table-cell": "A table cell containing an explicit page break cannot be represented in the editor model",
  "text-box": "A text box containing an explicit page break cannot be represented in the editor model",
} as const satisfies Record<ContainerContext["owner"], string>;

describe("page-break run source-container ownership", () => {
  for (const context of CONTAINER_CONTEXTS) {
    test.each(INLINE_WRAPPERS)(
      `rejects a page break in a ${context.name} through $name`,
      ({ wrap }) => {
        const paragraph: Paragraph = { type: "paragraph", content: [wrap(pageBreakRun())] };
        const source = context.build(paragraph);
        const before = structuredClone(source.package.document.content);

        expectUnsupportedConversion(() => toProseDoc(source), {
          message: CONTAINER_MESSAGES[context.owner],
          owner: context.owner,
          contentType: "break",
        });
        expect(source.package.document.content).toEqual(before);
      },
    );
  }

  test.each([
    {
      name: "only",
      content: [{ type: "break", breakType: "page" }],
    },
    {
      name: "leading",
      content: [
        { type: "break", breakType: "page" },
        { type: "text", text: "after" },
      ],
    },
    {
      name: "trailing",
      content: [
        { type: "text", text: "before" },
        { type: "break", breakType: "page" },
      ],
    },
    {
      name: "interleaved with rendered layout markers",
      content: [
        { type: "renderedPageBreak" },
        { type: "text", text: "before" },
        { type: "break", breakType: "page" },
        { type: "renderedPageBreak" },
        { type: "text", text: "after" },
      ],
    },
  ] as const satisfies readonly { name: string; content: readonly RunContent[] }[])(
    "rejects a $name authored page break regardless of its run position",
    ({ content }) => {
      const source = documentWithContent([
        tableWithParagraph({ type: "paragraph", content: [{ type: "run", content: [...content] }] }),
      ]);

      expectUnsupportedConversion(() => toProseDoc(source), {
        message: CONTAINER_MESSAGES["table-cell"],
        owner: "table-cell",
        contentType: "break",
      });
    },
  );

  test.each([
    { content: [{ type: "renderedPageBreak" }] },
    { content: [{ type: "break" }] },
    { content: [{ type: "break", breakType: "textWrapping" }] },
    { content: [{ type: "break", breakType: "column" }] },
    {
      content: [
        { type: "renderedPageBreak" },
        { type: "text", text: "authored content" },
        { type: "renderedPageBreak" },
      ],
    },
  ] as const satisfies readonly { content: readonly RunContent[] }[])(
    "does not confuse rendered or non-page breaks with an authored page break",
    ({ content }) => {
      const source = documentWithContent([
        tableWithParagraph({ type: "paragraph", content: [{ type: "run", content: [...content] }] }),
      ]);
      const before = structuredClone(source.package.document.content);

      expect(() => toProseDoc(source)).not.toThrow();
      expect(source.package.document.content).toEqual(before);
    },
  );
});

const PARAGRAPH_DISPOSITIONS = [
  {
    name: "non-drop frame",
    formatting: { frame: { width: 720 } },
    owner: "paragraph-frame",
    message: "A framed paragraph containing an explicit page-break run cannot be projected",
  },
  {
    name: "outline level",
    formatting: { outlineLevel: 0 },
    owner: "paragraph-outline",
    message: "An outline paragraph containing an explicit page-break run cannot be projected",
  },
  {
    name: "border",
    formatting: { borders: { bottom: { style: "single", size: 8 } } },
    owner: "paragraph-borders",
    message: "A bordered paragraph containing an explicit page-break run cannot be projected",
  },
] as const;

const STORY_SURFACES = [
  {
    name: "document body",
    convert: (paragraph: Paragraph) => toProseDoc(documentWithContent([paragraph])),
  },
  {
    name: "header or footer",
    convert: (paragraph: Paragraph) => headerFooterToProseDoc([paragraph]),
  },
  {
    name: "footnote or endnote",
    convert: (paragraph: Paragraph) => footnoteToProseDoc([paragraph]),
  },
] as const;

describe("page-break run source-paragraph ownership", () => {
  for (const surface of STORY_SURFACES) {
    test.each(PARAGRAPH_DISPOSITIONS)(
      `rejects a $name on the ${surface.name}`,
      ({ formatting, owner, message }) => {
        const paragraph: Paragraph = {
          type: "paragraph",
          formatting,
          content: [INLINE_WRAPPERS.at(-1)!.wrap(pageBreakRun())],
        };
        const before = structuredClone(paragraph);

        expectUnsupportedConversion(() => surface.convert(paragraph), {
          message,
          owner,
          contentType: "break",
        });
        expect(paragraph).toEqual(before);
      },
    );
  }

  test("rejects a page break sharing a paragraph with a text-box anchor", () => {
    const paragraph: Paragraph = {
      type: "paragraph",
      content: [
        pageBreakRun(),
        ...textBoxHost([{ type: "paragraph", content: [] }]).content,
      ],
    };

    expectUnsupportedConversion(() => toProseDoc(documentWithContent([paragraph])), {
      message:
        "A paragraph containing both an explicit page-break run and a text-box anchor cannot be projected",
      owner: "paragraph-text-box-anchor",
      contentType: "break",
    });
  });

  test("uses resolved paragraph attrs for style-owned projection constraints", () => {
    const source = documentWithContent([
      {
        type: "paragraph",
        formatting: { styleId: "Outlined" },
        content: [pageBreakRun()],
      },
    ]);
    source.package.styles = {
      styles: [{ styleId: "Outlined", type: "paragraph", pPr: { outlineLevel: 0 } }],
    };

    expectUnsupportedConversion(() => toProseDoc(source), {
      message: "An outline paragraph containing an explicit page-break run cannot be projected",
      owner: "paragraph-outline",
      contentType: "break",
    });
  });

  test.each(["drop", "margin"] as const)(
    "keeps a page break in a supported %s-cap paragraph",
    (dropCap) => {
      const source = documentWithContent([
        {
          type: "paragraph",
          formatting: { frame: { dropCap } },
          content: [pageBreakRun([{ type: "renderedPageBreak" }])],
        },
      ]);

      expect(() => toProseDoc(source)).not.toThrow();
    },
  );

  test("does not reject a rendered page break in an otherwise unprojectable paragraph shape", () => {
    const source = documentWithContent([
      {
        type: "paragraph",
        formatting: { outlineLevel: 0, borders: { bottom: { style: "single" } } },
        content: [{ type: "run", content: [{ type: "renderedPageBreak" }] }],
      },
    ]);

    expect(() => toProseDoc(source)).not.toThrow();
  });
});
