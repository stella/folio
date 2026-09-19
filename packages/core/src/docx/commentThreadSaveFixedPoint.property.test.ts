/**
 * Saving a threaded comment set must be a fixed point after the first save.
 *
 * The first save is allowed to normalise: it rewrites `word/comments.xml`, and
 * it mints the `w14:paraId` a thread needs when the model has none. What it may
 * not do is write a part the next save cannot reproduce. The public corpus
 * found it doing exactly that on `word/commentsExtended.xml`: comments.xml is
 * written top-level-first, so the reparse hands the comments back in that
 * order, while the extended part was built by walking the model's own order.
 *
 * The property generates comment forests rather than examples because the
 * defect is about order and about which comments carry an id: a fixture pins
 * one arrangement, and the arrangement is the variable.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import type { Comment, Document } from "../types/document";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";

type GeneratedComment = {
  /** Index of the comment this one replies to, or `null` for a thread root. */
  parent: number | null;
  done: boolean;
  text: string;
  /** Whether the model hands the save a comment with no paraId to thread by. */
  keepsParaId: boolean;
};

/**
 * A forest: every reply points at an earlier comment, so the generated shape is
 * always one Word would accept, and the order of roots and replies varies.
 */
const commentForest = fc
  .array(
    fc.record({
      parentOffset: fc.option(fc.nat({ max: 5 }), { nil: null }),
      done: fc.boolean(),
      text: fc.stringMatching(/^[A-Za-z0-9 ]{1,12}$/u),
      keepsParaId: fc.boolean(),
    }),
    { minLength: 1, maxLength: 6 },
  )
  .map((rows): GeneratedComment[] =>
    rows.map(({ parentOffset, done, text, keepsParaId }, index) => ({
      parent: parentOffset === null || index === 0 ? null : parentOffset % index,
      done,
      text,
      keepsParaId,
    })),
  );

const commentId = (index: number): number => index + 1;
const paraId = (index: number): string =>
  (0x0000_1000 + index).toString(16).toUpperCase().padStart(8, "0");

const buildDocx = async (forest: readonly GeneratedComment[]): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  const anchors = forest
    .map(
      (_, index) =>
        `<w:p><w:commentRangeStart w:id="${commentId(index)}"/><w:r><w:t>anchor</w:t></w:r>` +
        `<w:commentRangeEnd w:id="${commentId(index)}"/>` +
        `<w:r><w:commentReference w:id="${commentId(index)}"/></w:r></w:p>`,
    )
    .join("");
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}"><w:body>${anchors}<w:sectPr/></w:body></w:document>`,
  );
  zip.file(
    "word/comments.xml",
    `${XML_DECLARATION}<w:comments xmlns:w="${W_NAMESPACE}" xmlns:w14="${W14_NAMESPACE}">` +
      forest
        .map(
          ({ text }, index) =>
            `<w:comment w:id="${commentId(index)}" w:author="Author" w:date="2024-01-01T00:00:00Z">` +
            `<w:p w14:paraId="${paraId(index)}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`,
        )
        .join("") +
      "</w:comments>",
  );
  zip.file(
    "word/commentsExtended.xml",
    `${XML_DECLARATION}<w15:commentsEx xmlns:w="${W_NAMESPACE}" xmlns:w15="${W15_NAMESPACE}">` +
      forest
        .map(({ parent, done }, index) => {
          const parentAttribute = parent === null ? "" : ` w15:paraIdParent="${paraId(parent)}"`;
          return `<w15:commentEx w15:paraId="${paraId(index)}"${parentAttribute} w15:done="${done ? 1 : 0}"/>`;
        })
        .join("") +
      "</w15:commentsEx>",
  );
  const types = await zip.file("[Content_Types].xml")?.async("text");
  const rels = await zip.file("word/_rels/document.xml.rels")?.async("text");
  if (types === undefined || rels === undefined) {
    throw new Error("the empty package lost its packaging parts");
  }
  zip.file(
    "[Content_Types].xml",
    types.replace(
      "</Types>",
      '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>' +
        '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>',
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      '<Relationship Id="rIdPropComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
        '<Relationship Id="rIdPropCommentsEx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>',
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/**
 * Take the paraIds away that the generated case says the model does not have.
 *
 * This is the state the editor hands the save when a reply was written in the
 * app: a comment with a parent and no id to thread by, which the save has to
 * mint. Minting is what has to land in the save's own output.
 */
const stripParaIds = (comments: readonly Comment[], forest: readonly GeneratedComment[]): void => {
  for (const [index, comment] of comments.entries()) {
    if (forest[index]?.keepsParaId !== false) {
      continue;
    }
    for (const paragraph of comment.content ?? []) {
      paragraph.paraId = undefined;
    }
  }
};

/** Every part of a package, by path, as text or bytes. */
const packageParts = async (buffer: ArrayBuffer): Promise<Map<string, string>> => {
  const zip = await JSZip.loadAsync(buffer);
  const entries = Object.entries(zip.files).filter(([, file]) => !file.dir);
  return new Map(
    await Promise.all(
      entries.map(
        async ([path, file]): Promise<readonly [string, string]> => [
          path,
          await file.async("base64"),
        ],
      ),
    ),
  );
};

/** Which comment each comment replies to, keyed by comment id. */
const threading = ({ package: { document } }: Document): Record<number, number | null> =>
  Object.fromEntries(
    (document.comments ?? []).map((comment) => [comment.id, comment.parentId ?? null]),
  );

describe("saving threaded comments is a fixed point after the first save", () => {
  test("every part is byte-stable from the second save on", async () => {
    await fc.assert(
      fc.asyncProperty(commentForest, async (forest) => {
        const parsed = await parseDocx(await buildDocx(forest), { preloadFonts: false });
        stripParaIds(parsed.package.document.comments ?? [], forest);
        const expectedThreading = threading(parsed);

        const first = await repackDocx(parsed, { updateModifiedDate: false });
        const reparsed = await parseDocx(first, { preloadFonts: false });
        const second = await repackDocx(reparsed, { updateModifiedDate: false });

        const firstParts = await packageParts(first);
        const secondParts = await packageParts(second);
        expect([...secondParts.keys()].sort()).toEqual([...firstParts.keys()].sort());
        for (const [path, content] of firstParts) {
          expect({ path, content }).toEqual({ path, content: secondParts.get(path) });
        }

        // A stable byte sequence that lost the threading would be a fixed point
        // and a data loss, so the reply links are asserted separately.
        expect(threading(reparsed)).toEqual(expectedThreading);
      }),
      propertyConfig({ numRuns: 40 }),
    );
  });
});
