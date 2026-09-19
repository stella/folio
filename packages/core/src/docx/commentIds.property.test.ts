/**
 * Property tests for duplicate `w:id`s in `word/comments.xml`.
 *
 * `w:id` is how the body addresses a comment, so a package that defines two
 * comments under one id makes every marker naming it ambiguous. Word opens
 * such a package (Microsoft's own conformance corpus ships one) and resolves
 * each marker to the first definition. Over arbitrary multisets of comment ids:
 *
 *   1. The package parses: no throw, whatever the ids.
 *   2. The parsed comments carry distinct ids.
 *   3. Every marker that was unambiguous keeps its comment: an id the source
 *      defined once still resolves to that same comment's text.
 *   4. A duplicate is reported, never dropped silently.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { parseDocx } from "./parser";
import { RELATIONSHIP_TYPES } from "./relsParser";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The text the comment defined at this position in comments.xml carries. */
const commentBody = (position: number): string => `Comment body ${String(position)}`;

const commentsXmlFor = (ids: readonly number[]): string =>
  `${XML_DECLARATION}
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  ${ids
    .map(
      (id, position) =>
        `<w:comment w:id="${String(id)}" w:author="Reviewer" w:initials="R"><w:p><w:r><w:t>${commentBody(position)}</w:t></w:r></w:p></w:comment>`,
    )
    .join("")}
</w:comments>`;

/** One anchored range plus a reference per declared id, markers in id order. */
const documentXmlFor = (ids: readonly number[]): string =>
  `${XML_DECLARATION}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    ${[...new Set(ids)]
      .map(
        (id) =>
          `<w:p><w:commentRangeStart w:id="${String(id)}"/><w:r><w:t>Anchored ${String(id)}</w:t></w:r><w:commentRangeEnd w:id="${String(id)}"/><w:r><w:commentReference w:id="${String(id)}"/></w:r></w:p>`,
      )
      .join("")}
  </w:body>
</w:document>`;

const docxFor = async (ids: readonly number[]): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML_DECLARATION}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML_DECLARATION}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.comments}" Target="comments.xml"/>
</Relationships>`,
  );
  zip.file("word/comments.xml", commentsXmlFor(ids));
  zip.file("word/document.xml", documentXmlFor(ids));
  return zip.generateAsync({ type: "arraybuffer" });
};

/** Ids drawn from a small pool, so duplicates are generated often. */
const commentIdsArbitrary = fc.array(fc.integer({ min: 0, max: 4 }), {
  minLength: 1,
  maxLength: 8,
});

const plainText = (content: { content: { content: { text?: string }[] }[] }): string =>
  content.content.flatMap((run) => run.content.map((item) => item.text ?? "")).join("");

describe("duplicate comment ids (property)", () => {
  test(
    "parse keeps one comment per id and every unambiguous reference",
    async () => {
      await fc.assert(
        fc.asyncProperty(commentIdsArbitrary, async (ids) => {
          const document = await parseDocx(await docxFor(ids), { preloadFonts: false });
          const comments = document.package.document.comments ?? [];

          const parsedIds = comments.map((comment) => comment.id);
          expect(parsedIds).toEqual([...new Set(parsedIds)]);
          expect(new Set(parsedIds)).toEqual(new Set(ids));

          // A marker naming an id the source defined once is unambiguous, and
          // the comment it named must still be the one carrying that id.
          for (const [position, id] of ids.entries()) {
            if (ids.indexOf(id) !== position) {
              continue;
            }
            const comment = comments.find((candidate) => candidate.id === id);
            expect(comment?.content.map((paragraph) => plainText(paragraph)).join("")).toBe(
              commentBody(position),
            );
          }

          const duplicates = ids.length - new Set(ids).size;
          const reported = (document.warnings ?? []).filter((warning) =>
            warning.startsWith("Dropped "),
          );
          expect(reported.length).toBe(duplicates > 0 ? 1 : 0);
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(60_000),
  );
});
