/**
 * Every value folio writes into a package survives the package.
 *
 * The escaper's own property test (`xmlEscape.property.test.ts`) pins one
 * string against a parser. This one pins the whole path: a hostile string set
 * as a drawing title and alt text, a comment author and initials, a bookmark
 * name, a hyperlink tooltip, a style name, the core-properties title, a
 * content control's tag, alias and placeholder, and run text; saved through
 * the real save path; reopened with `parseDocx`; and compared to what the
 * boundary rule says the value should be.
 *
 * Before the escaper had one owner, six of those fields were written by four
 * different escapers, and none of them survived a carriage return.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { serializeDocumentToDocx } from "@stll/docx-core";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import { buildParagraphsDocx } from "../../ai-edits/__fixtures__/paragraphs";
import type { Document, Image, Paragraph, Run } from "../../types/document";
import { parseDocx } from "../parser";
import { repackDocx } from "../rezip";

/**
 * XML 1.0 §2.2 restated: what a value is once the package holds it. The
 * escaper may keep nothing else, and must keep all of this.
 */
const legalCharactersOf = (value: string): string =>
  value.replace(
    // eslint-disable-next-line no-control-regex -- the control range is the subject.
    /[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu,
    "",
  );

const hazards = [
  "]]>",
  "&",
  "<",
  ">",
  '"',
  "'",
  "&amp;",
  "\t",
  "\n",
  "\r",
  "\r\n",
  "\u0000",
  "\u000B",
  "\uFFFF",
  "\uD800",
  "\u{1F4C4}",
  "\u00A0",
  "x",
];

const hostileString = fc
  .array(fc.constantFrom(...hazards), { minLength: 1, maxLength: 8 })
  .map((pieces) => pieces.join(""));

const image = (value: string): Image => ({
  type: "image",
  size: { width: 914_400, height: 457_200 },
  wrap: { type: "inline" },
  alt: value,
  title: value,
});

/** Put `value` into every field this test claims to cover. */
const documentCarrying = (parsed: Document, value: string): Document => {
  const paragraph: Paragraph = {
    type: "paragraph",
    content: [
      { type: "bookmarkStart", id: 1, name: value },
      {
        type: "hyperlink",
        anchor: "top",
        tooltip: value,
        children: [{ type: "run", content: [{ type: "text", text: value }] } satisfies Run],
      },
      { type: "bookmarkEnd", id: 1 },
      { type: "commentRangeStart", id: 1 },
      { type: "run", content: [{ type: "drawing", image: image(value) }] },
      { type: "commentRangeEnd", id: 1 },
    ],
  };

  return {
    ...parsed,
    package: {
      ...parsed.package,
      styles: {
        ...parsed.package.styles,
        styles: [
          ...(parsed.package.styles?.styles ?? []).filter((style) => style.styleId !== "Marked"),
          { styleId: "Marked", type: "paragraph", name: value },
        ],
      },
      document: {
        ...parsed.package.document,
        content: [
          paragraph,
          {
            type: "blockSdt",
            properties: { sdtType: "richText", tag: value, alias: value, placeholder: value },
            content: [{ type: "paragraph", content: [{ type: "run", content: [] }] }],
          },
        ],
        comments: [
          {
            id: 1,
            // The model validator refuses a whitespace-only comment author, a rule
            // of its own; the letter keeps the value inside it.
            author: `A${value}`,
            initials: `A${value}`,
            content: [{ type: "paragraph", content: [{ type: "run", content: [] }] }],
          },
        ],
      },
    },
  };
};

type CarriedValues = Record<string, string | undefined>;

const valuesIn = (parsed: Document): CarriedValues => {
  const [paragraph, sdt] = parsed.package.document.content;
  const bookmark =
    paragraph?.type === "paragraph"
      ? paragraph.content.find((child) => child.type === "bookmarkStart")
      : undefined;
  const hyperlink =
    paragraph?.type === "paragraph"
      ? paragraph.content.find((child) => child.type === "hyperlink")
      : undefined;
  const drawingRun =
    paragraph?.type === "paragraph"
      ? paragraph.content.find(
          (child) => child.type === "run" && child.content.some((c) => c.type === "drawing"),
        )
      : undefined;
  const drawing =
    drawingRun?.type === "run"
      ? drawingRun.content.find((child) => child.type === "drawing")
      : undefined;
  const runText =
    hyperlink?.type === "hyperlink"
      ? hyperlink.children
          .flatMap((child) => (child.type === "run" ? child.content : []))
          .flatMap((child) => (child.type === "text" ? [child.text] : []))
          .join("")
      : undefined;
  const comment = parsed.package.document.comments?.at(0);

  return {
    bookmarkName: bookmark?.type === "bookmarkStart" ? bookmark.name : undefined,
    hyperlinkTooltip: hyperlink?.type === "hyperlink" ? hyperlink.tooltip : undefined,
    drawingAlt: drawing?.type === "drawing" ? drawing.image.alt : undefined,
    drawingTitle: drawing?.type === "drawing" ? drawing.image.title : undefined,
    runText,
    commentAuthor: comment?.author,
    commentInitials: comment?.initials,
    styleName: parsed.package.styles?.styles?.find((style) => style.styleId === "Marked")?.name,
    sdtTag: sdt?.type === "blockSdt" ? sdt.properties?.tag : undefined,
    sdtAlias: sdt?.type === "blockSdt" ? sdt.properties?.alias : undefined,
    sdtPlaceholder: sdt?.type === "blockSdt" ? sdt.properties?.placeholder : undefined,
  };
};

describe("a value written into a package comes back out of it", () => {
  test(
    "every field folio writes preserves the value the boundary rule allows",
    async () => {
      const base = await buildParagraphsDocx(["Original."]);
      const parsed = await parseDocx(base);

      await fc.assert(
        fc.asyncProperty(hostileString, async (value) => {
          // A value that sanitises to nothing asks a different question — does
          // an empty attribute mean "empty" or "absent"? — that several of
          // these fields answer "absent" on purpose. Out of scope here.
          fc.pre(legalCharactersOf(value).length > 0);
          const saved = await repackDocx(documentCarrying(parsed, value), {
            updateModifiedDate: false,
          });
          const reopened = await parseDocx(saved);
          const expected = legalCharactersOf(value);
          const { commentAuthor, commentInitials, ...fields } = valuesIn(reopened);
          for (const [field, actual] of Object.entries(fields)) {
            expect({ field, value: actual }).toEqual({ field, value: expected });
          }
          expect(commentInitials).toBe(`A${expected}`);
          // `parseCommentAuthor` trims the author, a normalisation that predates
          // this test and is not about escaping. Everything between the ends
          // survives, which is what the escaper is answerable for.
          expect(commentAuthor).toBe(`A${expected}`.trim());
        }),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(60_000),
  );

  // `docProps/core.xml` is replayed verbatim on the repack path — folio-core
  // never rewrites it — so the title's writer is `serializeDocumentToDocx`,
  // the build-from-scratch path, and that is where it is checked.
  test(
    "the core-properties title survives a package built from a model",
    async () => {
      await fc.assert(
        fc.asyncProperty(hostileString, async (value) => {
          const saved = await serializeDocumentToDocx({
            package: {
              properties: { title: value },
              document: {
                content: [{ type: "paragraph", content: [{ type: "run", content: [] }] }],
              },
            },
          });
          const reopened = await parseDocx(saved);
          expect(reopened.package.properties?.title).toBe(legalCharactersOf(value));
        }),
        propertyConfig({ numRuns: 40 }),
      );
    },
    propertyTestTimeout(60_000),
  );
});
