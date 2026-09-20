/**
 * A parse-time removal leaves the paragraph in the shape a parse produces.
 *
 * `parseParagraph` consolidates a paragraph's runs, and every inline item that
 * is not a run is a merge boundary. Three normalisers then remove items from
 * that already-consolidated array: a comment marker naming a comment the
 * package does not define, a move-range marker with no other half, and the
 * per-paragraph range markers a multi-paragraph comment is cut into. Each
 * removal can leave two mergeable runs adjacent, which the parse that
 * consolidated them would never have produced.
 *
 * That is not a loss but an oscillation. Save 1 writes the pair, the next parse
 * merges it, save 2 writes one run: the second save differs from the first,
 * which is the `save-idempotence` family's whole subject.
 *
 * The property generates the marker, not the defect: the paragraph is a random
 * sequence of same-formatted runs and markers, some resolvable and some not, so
 * a removal lands between runs, at an edge, beside another removal and not at
 * all. Both fixed points are asserted, because the bytes settling is worth
 * nothing if the text moved.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { getParagraphText } from "./paragraphParser";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const COMMENTS_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

/** The comment `word/comments.xml` defines. Every other id resolves to nothing. */
const DEFINED_COMMENT_ID = 1;
const DANGLING_COMMENT_ID = 7;

/**
 * One inline item, by what the parse does with it.
 *
 * `run` is the material a removal can merge; the markers are the boundaries.
 * `definedRange*` survive the parse, `dangling*` are removed by
 * `commentReferenceNormalization`, and the move-range halves are removed by
 * `trackedMoveRangeNormalization` because each is generated without its twin.
 */
const ITEMS = {
  run: "run",
  definedRangeStart: "definedRangeStart",
  definedRangeEnd: "definedRangeEnd",
  danglingRangeStart: "danglingRangeStart",
  danglingRangeEnd: "danglingRangeEnd",
  danglingReference: "danglingReference",
  moveFromRangeStart: "moveFromRangeStart",
  moveToRangeEnd: "moveToRangeEnd",
  bookmarkStart: "bookmarkStart",
} as const;

type Item = (typeof ITEMS)[keyof typeof ITEMS];

const itemMarkup = (item: Item, index: number): string => {
  switch (item) {
    case ITEMS.run:
      return `<w:r><w:rPr><w:b/></w:rPr><w:t>t${index}</w:t></w:r>`;
    case ITEMS.definedRangeStart:
      return `<w:commentRangeStart w:id="${DEFINED_COMMENT_ID}"/>`;
    case ITEMS.definedRangeEnd:
      return `<w:commentRangeEnd w:id="${DEFINED_COMMENT_ID}"/>`;
    case ITEMS.danglingRangeStart:
      return `<w:commentRangeStart w:id="${DANGLING_COMMENT_ID}"/>`;
    case ITEMS.danglingRangeEnd:
      return `<w:commentRangeEnd w:id="${DANGLING_COMMENT_ID}"/>`;
    case ITEMS.danglingReference:
      return `<w:r><w:commentReference w:id="${DANGLING_COMMENT_ID}"/></w:r>`;
    case ITEMS.moveFromRangeStart:
      return '<w:moveFromRangeStart w:id="11" w:name="m1"/>';
    case ITEMS.moveToRangeEnd:
      return '<w:moveToRangeEnd w:id="12"/>';
    case ITEMS.bookmarkStart:
      return `<w:bookmarkStart w:id="${index}" w:name="b${index}"/>`;
    default:
      return item satisfies never;
  }
};

const COMMENTS_PART =
  `${XML_DECLARATION}<w:comments xmlns:w="${W_NAMESPACE}">` +
  `<w:comment w:id="${DEFINED_COMMENT_ID}" w:author="A" w:initials="A">` +
  "<w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>";

type Package = { items: readonly Item[]; withComments: boolean };

const buildDocx = async ({ items, withComments }: Package): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:r="${R_NAMESPACE}"><w:body>` +
      `<w:p>${items.map(itemMarkup).join("")}</w:p>` +
      "<w:sectPr/></w:body></w:document>",
  );
  if (!withComments) {
    return zip.generateAsync({ type: "arraybuffer" });
  }
  zip.file("word/comments.xml", COMMENTS_PART);
  const types = await zip.file("[Content_Types].xml")!.async("text");
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      `<Override PartName="/word/comments.xml" ContentType="${COMMENTS_CONTENT_TYPE}"/></Types>`,
    ),
  );
  const rels = await zip.file("word/_rels/document.xml.rels")!.async("text");
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      `<Relationship Id="rIdComments" Type="${R_NAMESPACE}/comments" Target="comments.xml"/></Relationships>`,
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const documentPart = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file("word/document.xml")?.async("text")) ?? "";

const bodyText = ({ package: { document } }: Awaited<ReturnType<typeof parseDocx>>): string =>
  document.content
    .filter((block) => block.type === "paragraph")
    .map((paragraph) => getParagraphText(paragraph))
    .join("\n");

const generated = fc.record({
  items: fc.array(fc.constantFrom(...Object.values(ITEMS)), { minLength: 1, maxLength: 7 }),
  withComments: fc.boolean(),
});

describe("a removal restores the consolidation the parse guarantees", () => {
  test(
    "the second save writes what the first save wrote, with the same text",
    async () => {
      await fc.assert(
        fc.asyncProperty(generated, async (generatedPackage) => {
          const source = await buildDocx(generatedPackage);

          const first = await parseDocx(source, { preloadFonts: false });
          const firstSave = await repackDocx(first, { updateModifiedDate: false });
          const second = await parseDocx(firstSave, { preloadFonts: false });
          const secondSave = await repackDocx(second, { updateModifiedDate: false });

          expect({ generatedPackage, text: bodyText(second) }).toEqual({
            generatedPackage,
            text: bodyText(first),
          });
          expect({ generatedPackage, part: await documentPart(secondSave) }).toEqual({
            generatedPackage,
            part: await documentPart(firstSave),
          });
        }),
        propertyConfig({ numRuns: 150 }),
      );
    },
    propertyTestTimeout(120_000),
  );
});
