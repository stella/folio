import { escapeXmlText } from "@stll/docx-core";
import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
  RunContent,
  NoteReferenceContent,
} from "../src/types/document";

const textRun = (text: string) =>
  ({ type: "run", content: [{ type: "text", text }] }) satisfies Run;

export type CanonicalInlineShapeFixture<Item> = {
  variant: string;
  item: Item;
  content: ParagraphContent[];
  story: "main" | "footnote" | "endnote";
};

type ParagraphFixtureFactories = {
  [Kind in ParagraphContent["type"]]: (
    text?: string,
  ) => CanonicalInlineShapeFixture<Extract<ParagraphContent, { type: Kind }>>[];
};

type RunFixtureFactories = {
  [Kind in RunContent["type"]]: (
    text?: string,
  ) => CanonicalInlineShapeFixture<
    Kind extends NoteReferenceContent["type"]
      ? NoteReferenceContent
      : Extract<RunContent, { type: Kind }>
  >[];
};

const paragraphFixture = <Item extends ParagraphContent>(item: Item, variant = "plain") =>
  ({ item, content: [item], variant, story: "main" }) satisfies CanonicalInlineShapeFixture<Item>;

const runFixture = <Item extends RunContent>(item: Item, variant = "plain") =>
  ({
    item,
    variant,
    content: [{ type: "run", content: [item] } satisfies Run],
    story: "main",
  }) satisfies CanonicalInlineShapeFixture<Item>;

type RangeFixturesOptions<Item extends ParagraphContent> = {
  item: Item;
  start: ParagraphContent;
  end: ParagraphContent;
  text: string;
};

const rangeFixtures = <Item extends ParagraphContent>({
  item,
  start,
  end,
  text,
}: RangeFixturesOptions<Item>) => [
  {
    item,
    content: [start, end],
    variant: "empty",
    story: "main",
  } satisfies CanonicalInlineShapeFixture<Item>,
  {
    item,
    content: [start, textRun(text), end],
    variant: "filled",
    story: "main",
  } satisfies CanonicalInlineShapeFixture<Item>,
];

const bookmarkStart = { type: "bookmarkStart", id: 11, name: "target" } satisfies ParagraphContent;
const bookmarkEnd = { type: "bookmarkEnd", id: 11 } satisfies ParagraphContent;
const commentStart = { type: "commentRangeStart", id: 12 } satisfies ParagraphContent;
const commentEnd = { type: "commentRangeEnd", id: 12 } satisfies ParagraphContent;
const moveFromStart = {
  type: "moveFromRangeStart",
  id: 13,
  name: "source",
  author: "Reviewer",
} satisfies ParagraphContent;
const moveFromEnd = { type: "moveFromRangeEnd", id: 13 } satisfies ParagraphContent;
const moveToStart = {
  type: "moveToRangeStart",
  id: 14,
  name: "destination",
  author: "Reviewer",
} satisfies ParagraphContent;
const moveToEnd = { type: "moveToRangeEnd", id: 14 } satisfies ParagraphContent;
const revision = { id: 21, author: "Reviewer", date: "2026-01-01T00:00:00Z" };

export const CANONICAL_PARAGRAPH_SHAPE_FACTORIES = {
  run: (text = "x") => [paragraphFixture(textRun(text))],
  hyperlink: (text = "x") => [
    paragraphFixture({
      type: "hyperlink",
      href: "https://example.test/",
      children: [textRun(text)],
    }),
    paragraphFixture({ type: "hyperlink", href: "https://empty.example/", children: [] }, "empty"),
  ],
  bookmarkStart: (text = "x") => [
    ...rangeFixtures({ item: bookmarkStart, start: bookmarkStart, end: bookmarkEnd, text }),
    paragraphFixture(bookmarkStart, "unpaired"),
  ],
  bookmarkEnd: (text = "x") => [
    ...rangeFixtures({ item: bookmarkEnd, start: bookmarkStart, end: bookmarkEnd, text }),
    paragraphFixture(bookmarkEnd, "unpaired"),
  ],
  commentRangeStart: (text = "x") =>
    rangeFixtures({ item: commentStart, start: commentStart, end: commentEnd, text }),
  commentRangeEnd: (text = "x") =>
    rangeFixtures({ item: commentEnd, start: commentStart, end: commentEnd, text }),
  commentReference: () => [paragraphFixture({ type: "commentReference", id: 12 })],
  moveFromRangeStart: (text = "x") =>
    rangeFixtures({ item: moveFromStart, start: moveFromStart, end: moveFromEnd, text }),
  moveFromRangeEnd: (text = "x") =>
    rangeFixtures({ item: moveFromEnd, start: moveFromStart, end: moveFromEnd, text }),
  moveToRangeStart: (text = "x") =>
    rangeFixtures({ item: moveToStart, start: moveToStart, end: moveToEnd, text }),
  moveToRangeEnd: (text = "x") =>
    rangeFixtures({ item: moveToEnd, start: moveToStart, end: moveToEnd, text }),
  simpleField: (text = "x") => [
    paragraphFixture({
      type: "simpleField",
      instruction: "REF target",
      fieldType: "REF",
      content: [textRun(text)],
    }),
    paragraphFixture(
      {
        type: "simpleField",
        instruction: "REF target",
        fieldType: "REF",
        content: [{ type: "hyperlink", href: "https://example.test/", children: [textRun(text)] }],
      },
      "structured",
    ),
  ],
  complexField: (text = "x") => [
    paragraphFixture({
      type: "complexField",
      instruction: "REF target",
      fieldType: "REF",
      fieldCode: [{ type: "run", content: [{ type: "instrText", text: "REF target" }] }],
      fieldResult: [textRun(text)],
    }),
  ],
  insertion: (text = "x") => [
    paragraphFixture({ type: "insertion", info: revision, content: [textRun(text)] }),
  ],
  deletion: (text = "x") => [
    paragraphFixture({ type: "deletion", info: revision, content: [textRun(text)] }),
  ],
  moveFrom: (text = "x") => [
    paragraphFixture({ type: "moveFrom", info: revision, content: [textRun(text)] }),
  ],
  moveTo: (text = "x") => [
    paragraphFixture({ type: "moveTo", info: revision, content: [textRun(text)] }),
  ],
  inlineWrapper: (text = "x") => [
    paragraphFixture(
      {
        type: "inlineWrapper",
        kind: "bidi",
        control: "override",
        direction: "rtl",
        content: [textRun(text)],
      },
      "bidi",
    ),
    paragraphFixture(
      {
        type: "inlineWrapper",
        kind: "smartTag",
        element: "place",
        content: [bookmarkStart, textRun(text), bookmarkEnd],
      },
      "smartTag",
    ),
    paragraphFixture(
      {
        type: "inlineWrapper",
        kind: "customXml",
        element: "value",
        content: [{ type: "insertion", info: revision, content: [textRun(text)] }],
      },
      "customXml",
    ),
    paragraphFixture(
      { type: "inlineWrapper", kind: "bidi", control: "override", direction: "rtl", content: [] },
      "emptyBidi",
    ),
    paragraphFixture(
      { type: "inlineWrapper", kind: "smartTag", element: "place", content: [] },
      "emptySmartTag",
    ),
    paragraphFixture(
      { type: "inlineWrapper", kind: "customXml", element: "value", content: [] },
      "emptyCustomXml",
    ),
  ],
  inlineSdt: (text = "x") => [
    paragraphFixture({
      type: "inlineSdt",
      properties: { sdtType: "richText", id: 31 },
      content: [textRun(text)],
    }),
  ],
  mathEquation: () => [
    paragraphFixture({
      type: "mathEquation",
      display: "inline",
      ommlXml:
        '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>',
      plainText: "x",
    }),
  ],
  preservedInline: (text = "x") => [
    paragraphFixture(
      {
        type: "preservedInline",
        xml: '<w:proofErr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:type="spellStart"/>',
        text: "",
      },
      "empty",
    ),
    paragraphFixture({
      type: "preservedInline",
      xml: `<w:customXml xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>${escapeXmlText(text)}</w:t></w:r></w:customXml>`,
      text,
    }),
  ],
} satisfies ParagraphFixtureFactories;

const rawFieldBegin = { type: "fieldChar", charType: "begin" } satisfies RunContent;
const rawFieldInstruction = { type: "instrText", text: "REF target" } satisfies RunContent;
const rawFieldContent = (text: string) => [
  rawFieldBegin,
  rawFieldInstruction,
  { type: "fieldChar", charType: "separate" } satisfies RunContent,
  { type: "text", text } satisfies RunContent,
  { type: "fieldChar", charType: "end" } satisfies RunContent,
];

const adjacentReferenceFixtures = (item: NoteReferenceContent) => [
  runFixture(item),
  ...[2, 3].map(
    (count) =>
      ({
        item,
        story: "main",
        variant: `adjacentIdenticalMarks${count}`,
        content: [
          { type: "run", content: Array.from({ length: count }, () => item) } satisfies Run,
        ],
      }) satisfies CanonicalInlineShapeFixture<NoteReferenceContent>,
  ),
];

export const CANONICAL_RUN_SHAPE_FACTORIES = {
  text: (text = "x") => [runFixture({ type: "text", text })],
  tab: () => [runFixture({ type: "tab" })],
  break: () => [
    runFixture({ type: "break", breakType: "textWrapping" }),
    runFixture({ type: "break", breakType: "column" }, "column"),
    runFixture({ type: "break", breakType: "page" }, "page"),
  ],
  symbol: () => [runFixture({ type: "symbol", font: "Symbol", char: "F061" })],
  footnoteRef: () => adjacentReferenceFixtures({ type: "footnoteRef", id: 1 }),
  endnoteRef: () => adjacentReferenceFixtures({ type: "endnoteRef", id: 2 }),
  noteMarker: () => [
    { ...runFixture({ type: "noteMarker", kind: "footnote" }), story: "footnote" },
    { ...runFixture({ type: "noteMarker", kind: "endnote" }, "endnote"), story: "endnote" },
  ],
  fieldChar: (text = "x") => [
    {
      item: rawFieldBegin,
      variant: "balanced",
      content: [{ type: "run", content: rawFieldContent(text) } satisfies Run],
      story: "main",
    },
  ],
  instrText: (text = "x") => [
    {
      item: rawFieldInstruction,
      variant: "balanced",
      content: [{ type: "run", content: rawFieldContent(text) } satisfies Run],
      story: "main",
    },
  ],
  softHyphen: () => [runFixture({ type: "softHyphen" })],
  noBreakHyphen: () => [runFixture({ type: "noBreakHyphen" })],
  renderedPageBreak: () => [runFixture({ type: "renderedPageBreak" })],
  preservedXml: (text = "x") => [
    runFixture({
      type: "preservedXml",
      xml: `<w:ruby xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rubyBase><w:r><w:t>${escapeXmlText(text)}</w:t></w:r></w:rubyBase></w:ruby>`,
      text,
    }),
    runFixture(
      {
        type: "preservedXml",
        xml: '<w:separator xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
        text: "",
      },
      "empty",
    ),
  ],
  drawing: () => [
    runFixture({
      type: "drawing",
      image: { type: "image", size: { width: 914400, height: 914400 }, wrap: { type: "inline" } },
    }),
  ],
  shape: () => [
    runFixture({
      type: "shape",
      shape: { type: "shape", shapeType: "rect", size: { width: 914400, height: 914400 } },
    }),
  ],
} satisfies RunFixtureFactories;

type ShapeDocumentOptions = {
  content: ParagraphContent[];
  story: CanonicalInlineShapeFixture<ParagraphContent>["story"];
};

const paragraph = (paraId: string, content: ParagraphContent[]) =>
  ({ type: "paragraph", paraId, content }) satisfies Paragraph;

export const canonicalInlineShapeDocument = ({ content, story }: ShapeDocumentOptions) => {
  const noteParagraph = (kind: "footnote" | "endnote", paraId: string) =>
    paragraph(
      paraId,
      story === kind
        ? [textRun("L"), ...content, textRun("R")]
        : [
            {
              type: "run",
              content: [
                { type: "noteMarker", kind },
                { type: "text", text: "Note" },
              ],
            },
          ],
    );
  return {
    package: {
      document: {
        content: [
          paragraph("12345678", [textRun("L"), ...(story === "main" ? content : []), textRun("R")]),
        ],
        comments: [
          { id: 12, author: "Reviewer", content: [paragraph("22345678", [textRun("Comment")])] },
        ],
      },
      footnotes: [{ type: "footnote", id: 1, content: [noteParagraph("footnote", "32345678")] }],
      endnotes: [{ type: "endnote", id: 2, content: [noteParagraph("endnote", "42345678")] }],
    },
  } satisfies Document;
};
