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
import { fromProseDoc } from "./fromProseDoc";
import {
  footnoteToProseDoc,
  headerFooterToProseDoc,
  standaloneTableCellToProseMirror,
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

const pageBreakParagraph = (): Paragraph => ({ type: "paragraph", content: [pageBreakRun()] });

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

const textBoxRun = (content: (Paragraph | Table)[]): Run => ({
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
});

const textBoxHost = (content: (Paragraph | Table)[]): Paragraph => ({
  type: "paragraph",
  content: [textBoxRun(content)],
});

const documentWithContent = (content: BlockContent[]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = content;
  return document;
};

type ContainerContext = {
  name: string;
  owner: "table-cell" | "text-box";
  supportsLeadingBreak: boolean;
  build: (paragraph: Paragraph) => Document;
};

const CONTAINER_CONTEXTS = [
  {
    name: "ordinary table cell",
    owner: "table-cell",
    supportsLeadingBreak: true,
    build: (paragraph) => documentWithContent([tableWithParagraph(paragraph)]),
  },
  {
    name: "repeating-header table cell",
    owner: "table-cell",
    supportsLeadingBreak: true,
    build: (paragraph) => documentWithContent([tableWithParagraph(paragraph, true)]),
  },
  {
    name: "nested table cell",
    owner: "table-cell",
    supportsLeadingBreak: false,
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
    supportsLeadingBreak: false,
    build: (paragraph) => documentWithContent([textBoxHost([paragraph])]),
  },
  {
    name: "table nested in a text box",
    owner: "text-box",
    supportsLeadingBreak: false,
    build: (paragraph) => documentWithContent([textBoxHost([tableWithParagraph(paragraph)])]),
  },
  {
    name: "text box nested in a table cell",
    owner: "table-cell",
    supportsLeadingBreak: false,
    build: (paragraph) => documentWithContent([tableWithParagraph(textBoxHost([paragraph]))]),
  },
] as const satisfies readonly ContainerContext[];

const CONTAINER_MESSAGES = {
  "table-cell":
    "A table cell containing an explicit page break cannot be represented in the editor model",
  "text-box":
    "A text box containing an explicit page break cannot be represented in the editor model",
} as const satisfies Record<ContainerContext["owner"], string>;

describe("page-break run source-container ownership", () => {
  for (const context of CONTAINER_CONTEXTS) {
    test.each(INLINE_WRAPPERS)(
      `classifies a leading page break in a ${context.name} through $name`,
      ({ wrap }) => {
        const paragraph: Paragraph = { type: "paragraph", content: [wrap(pageBreakRun())] };
        const source = context.build(paragraph);
        const before = structuredClone(source.package.document.content);

        if (context.supportsLeadingBreak) {
          const prose = toProseDoc(source);
          let pageBreaks = 0;
          prose.descendants((node) => {
            if (node.type.name === "pageBreakRun") pageBreaks += 1;
          });
          expect(pageBreaks).toBe(1);
          expect(source.package.document.content).toEqual(before);
          return;
        }

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
      supported: true,
      content: [{ type: "break", breakType: "page" }],
    },
    {
      name: "leading",
      supported: true,
      content: [
        { type: "break", breakType: "page" },
        { type: "text", text: "after" },
      ],
    },
    {
      name: "trailing",
      supported: false,
      content: [
        { type: "text", text: "before" },
        { type: "break", breakType: "page" },
      ],
    },
    {
      name: "interleaved with rendered layout markers",
      supported: false,
      content: [
        { type: "renderedPageBreak" },
        { type: "text", text: "before" },
        { type: "break", breakType: "page" },
        { type: "renderedPageBreak" },
        { type: "text", text: "after" },
      ],
    },
  ] as const satisfies readonly {
    name: string;
    supported: boolean;
    content: readonly RunContent[];
  }[])("classifies a $name authored page break by its run position", ({ content, supported }) => {
    const source = documentWithContent([
      tableWithParagraph({
        type: "paragraph",
        content: [{ type: "run", content: [...content] }],
      }),
    ]);

    if (supported) {
      expect(() => toProseDoc(source)).not.toThrow();
      return;
    }

    expectUnsupportedConversion(() => toProseDoc(source), {
      message: CONTAINER_MESSAGES["table-cell"],
      owner: "table-cell",
      contentType: "break",
    });
  });

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
        tableWithParagraph({
          type: "paragraph",
          content: [{ type: "run", content: [...content] }],
        }),
      ]);
      const before = structuredClone(source.package.document.content);

      expect(() => toProseDoc(source)).not.toThrow();
      expect(source.package.document.content).toEqual(before);
    },
  );

  test.each([
    {
      name: "standalone table-cell conversion",
      convert: (paragraph: Paragraph) =>
        standaloneTableCellToProseMirror({ type: "tableCell", content: [paragraph] }, "tableCell"),
    },
    {
      name: "header or footer conversion",
      convert: (paragraph: Paragraph) => headerFooterToProseDoc([tableWithParagraph(paragraph)]),
    },
    {
      name: "footnote or endnote conversion",
      convert: (paragraph: Paragraph) => footnoteToProseDoc([tableWithParagraph(paragraph)]),
    },
  ])("supports a leading row boundary through $name", ({ convert }) => {
    expect(() => convert(pageBreakParagraph())).not.toThrow();
  });

  test("round-trips a leading row boundary after a zero-width bookmark", () => {
    const source = documentWithContent([
      tableWithParagraph({
        type: "paragraph",
        content: [
          { type: "bookmarkStart", id: 7, name: "boundary" },
          {
            type: "run",
            content: [
              { type: "break", breakType: "page" },
              { type: "text", text: "after" },
            ],
          },
          { type: "bookmarkEnd", id: 7 },
        ],
      }),
    ]);

    const restored = fromProseDoc(toProseDoc(source), source);

    expect(restored.package.document.content).toEqual(source.package.document.content);
  });

  test.each([
    {
      name: "table cell around a text box",
      source: () => documentWithContent([tableWithParagraph(textBoxHost([pageBreakParagraph()]))]),
      owner: "table-cell" as const,
    },
    {
      name: "text box around a table cell",
      source: () => documentWithContent([textBoxHost([tableWithParagraph(pageBreakParagraph())])]),
      owner: "text-box" as const,
    },
  ])("reports the outermost $name owner", ({ source, owner }) => {
    expectUnsupportedConversion(() => toProseDoc(source()), {
      message: CONTAINER_MESSAGES[owner],
      owner,
      contentType: "break",
    });
  });

  test("does not query synthesized cells outside the source index", () => {
    const source = documentWithContent([
      {
        type: "table",
        rows: [{ type: "tableRow", cells: [] }],
      },
    ]);

    expect(() => toProseDoc(source)).not.toThrow();
  });

  test("rejects a page break in a skipped vertical-merge continuation cell", () => {
    const source = documentWithContent([
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                formatting: { vMerge: "restart" },
                content: [{ type: "paragraph", content: [] }],
              },
            ],
          },
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                formatting: { vMerge: "continue" },
                content: [pageBreakParagraph()],
              },
            ],
          },
        ],
      },
    ]);

    expectUnsupportedConversion(() => toProseDoc(source), {
      message: CONTAINER_MESSAGES["table-cell"],
      owner: "table-cell",
      contentType: "break",
    });
  });
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
      `supports a leading page break in a $name on the ${surface.name}`,
      ({ formatting }) => {
        const paragraph: Paragraph = {
          type: "paragraph",
          formatting,
          content: [INLINE_WRAPPERS.at(-1)!.wrap(pageBreakRun())],
        };
        const before = structuredClone(paragraph);

        expect(() => surface.convert(paragraph)).not.toThrow();
        expect(paragraph).toEqual(before);
      },
    );

    test.each(PARAGRAPH_DISPOSITIONS)(
      `rejects an interior page break in a $name on the ${surface.name}`,
      ({ formatting, owner, message }) => {
        const paragraph: Paragraph = {
          type: "paragraph",
          formatting,
          content: [
            { type: "run", content: [{ type: "text", text: "before" }] },
            INLINE_WRAPPERS.at(-1)!.wrap(pageBreakRun()),
          ],
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
      content: [pageBreakRun(), ...textBoxHost([{ type: "paragraph", content: [] }]).content],
    };

    expectUnsupportedConversion(() => toProseDoc(documentWithContent([paragraph])), {
      message:
        "A paragraph containing both an explicit page-break run and a text-box anchor cannot be projected",
      owner: "paragraph-text-box-anchor",
      contentType: "break",
    });
  });

  test.each(INLINE_WRAPPERS)(
    "finds a text-box shape through $name from source truth",
    ({ wrap }) => {
      const paragraph: Paragraph = {
        type: "paragraph",
        content: [pageBreakRun(), wrap(textBoxRun([{ type: "paragraph", content: [] }]))],
      };

      expectUnsupportedConversion(() => toProseDoc(documentWithContent([paragraph])), {
        message:
          "A paragraph containing both an explicit page-break run and a text-box anchor cannot be projected",
        owner: "paragraph-text-box-anchor",
        contentType: "break",
      });
    },
  );

  test("supports a leading page break with a style-owned outline level", () => {
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

    expect(() => toProseDoc(source)).not.toThrow();
  });

  test("supports a leading page break in a paragraph-style-owned frame", () => {
    const source = documentWithContent([
      {
        type: "paragraph",
        formatting: { styleId: "Framed" },
        content: [pageBreakRun()],
      },
    ]);
    source.package.styles = {
      styles: [{ styleId: "Framed", type: "paragraph", pPr: { frame: { width: 720 } } }],
    };

    expect(() => toProseDoc(source)).not.toThrow();
  });

  test("supports a leading table-row boundary with a table-style-owned frame", () => {
    const table = tableWithParagraph(pageBreakParagraph());
    table.formatting = { styleId: "FramedTable" };
    const source = documentWithContent([table]);
    source.package.styles = {
      styles: [{ styleId: "FramedTable", type: "table", pPr: { frame: { width: 720 } } }],
    };

    expect(() => toProseDoc(source)).not.toThrow();
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
