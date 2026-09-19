/**
 * A comment keeps its own identity across a round trip.
 *
 * The `w:id` is the only name the package has for a comment: every
 * `w:commentRangeStart`, `w:commentRangeEnd` and `w:commentReference` in the
 * body addresses one, and nothing else does. So the facts a comment carries —
 * who wrote it, what it says, when, whether it is resolved, which comment it
 * replies to — must still belong to that same `w:id` after a save, whatever
 * order the parts were written in.
 *
 * The public corpus found them moving. `word/comments.xml` was written
 * top-level comments first and replies after, so a document whose comments.xml
 * interleaves a reply with a later thread root came back from the next parse in
 * a different `comments[]` order than it went in, and everything reading that
 * array by position — folio's own comparison engine among them — read one
 * person's words under another's name.
 *
 * The property generates comment forests rather than examples because the
 * defect is about order: a fixture pins one arrangement, and the arrangement is
 * the variable. Identity is asserted per `w:id`, not as a multiset, because a
 * swap of two comments' bodies preserves the multiset exactly.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";
import type { Comment, Document } from "../types/document";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const W15_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";

/** Where a comment's range is anchored: the three surfaces comments.xml serves. */
const ANCHOR_SURFACES = ["body", "table", "header"] as const;
type AnchorSurface = (typeof ANCHOR_SURFACES)[number];

/** What a comment does with the `w14:paraId` its thread key comes from. */
const PARA_ID_STATES = ["own", "none", "shared"] as const;
type ParaIdState = (typeof PARA_ID_STATES)[number];

type GeneratedComment = {
  /** Index of the comment this one replies to, or `null` for a thread root. */
  parent: number | null;
  done: boolean;
  author: string;
  initials: string;
  /** One entry per paragraph: a multi-paragraph comment keys on its last. */
  paragraphs: readonly string[];
  paraIdState: ParaIdState;
  anchor: AnchorSurface;
};

/**
 * A forest whose replies always point at an earlier comment, so the shape is
 * one Word would accept, while the order roots and replies appear in, the
 * number of paragraphs each carries and which of them have ids all vary.
 */
const commentForest = fc
  .array(
    fc.record({
      parentOffset: fc.option(fc.nat({ max: 5 }), { nil: null }),
      done: fc.boolean(),
      author: fc.constantFrom("Author A", "Author B", "Author C"),
      initials: fc.constantFrom("AA", "AB", "AC"),
      paragraphs: fc.array(fc.stringMatching(/^[A-Za-z0-9 ]{1,10}$/u), {
        minLength: 1,
        maxLength: 3,
      }),
      paraIdState: fc.constantFrom(...PARA_ID_STATES),
      anchor: fc.constantFrom(...ANCHOR_SURFACES),
    }),
    { minLength: 1, maxLength: 6 },
  )
  .map((rows): GeneratedComment[] =>
    rows.map((row, index) => ({
      ...row,
      parent: row.parentOffset === null || index === 0 ? null : row.parentOffset % index,
      // A thread root cannot go id-less or share an id: its key is what its
      // replies name, and a package that loses it is malformed input, not a
      // round trip to hold folio to.
      paraIdState: rows.some(
        (other, otherIndex) =>
          otherIndex > index &&
          other.parentOffset !== null &&
          otherIndex !== 0 &&
          other.parentOffset % otherIndex === index,
      )
        ? "own"
        : row.paraIdState,
    })),
  );

const commentId = (index: number): number => index + 1;

const paraIdFor = (forest: readonly GeneratedComment[], index: number): string | null => {
  const state = forest[index]?.paraIdState;
  if (state === "none") {
    return null;
  }
  // "shared" hands two comments the same `w14:paraId`, which Word-written
  // files do: the reading has to be deterministic rather than last-wins.
  const owner = state === "shared" ? Math.max(0, index - 1) : index;
  return (0x0000_1000 + owner).toString(16).toUpperCase().padStart(8, "0");
};

const anchorMarkers = (index: number): string =>
  `<w:commentRangeStart w:id="${commentId(index)}"/><w:r><w:t>anchor</w:t></w:r>` +
  `<w:commentRangeEnd w:id="${commentId(index)}"/>` +
  `<w:r><w:commentReference w:id="${commentId(index)}"/></w:r>`;

const anchorsFor = (forest: readonly GeneratedComment[], surface: AnchorSurface): string => {
  const markers = forest
    .map((comment, index) => (comment.anchor === surface ? anchorMarkers(index) : ""))
    .join("");
  if (markers === "") {
    return "";
  }
  if (surface !== "table") {
    return `<w:p>${markers}</w:p>`;
  }
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4675"/></w:tblGrid>' +
    `<w:tr><w:tc><w:tcPr><w:tcW w:w="4675" w:type="dxa"/></w:tcPr><w:p>${markers}</w:p></w:tc></w:tr></w:tbl>`
  );
};

const HEADER_RELATIONSHIP_ID = "rIdPropHeader";

const buildDocx = async (forest: readonly GeneratedComment[]): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());

  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>` +
      anchorsFor(forest, "body") +
      anchorsFor(forest, "table") +
      `<w:sectPr><w:headerReference w:type="default" r:id="${HEADER_RELATIONSHIP_ID}"/></w:sectPr>` +
      "</w:body></w:document>",
  );
  zip.file(
    "word/header1.xml",
    `${XML_DECLARATION}<w:hdr xmlns:w="${W_NAMESPACE}">${anchorsFor(forest, "header") || "<w:p/>"}</w:hdr>`,
  );

  zip.file(
    "word/comments.xml",
    `${XML_DECLARATION}<w:comments xmlns:w="${W_NAMESPACE}" xmlns:w14="${W14_NAMESPACE}">` +
      forest
        .map(({ author, initials, paragraphs }, index) => {
          const paraId = paraIdFor(forest, index);
          const body = paragraphs
            .map((text, paragraphIndex) => {
              // Only the LAST paragraph carries the thread key, the way Word
              // writes it; the earlier ones get ids of their own.
              const isLast = paragraphIndex === paragraphs.length - 1;
              const id = isLast
                ? paraId
                : `${(0x0020_0000 + index * 8 + paragraphIndex).toString(16).toUpperCase()}`;
              const attribute = id === null ? "" : ` w14:paraId="${id}"`;
              return `<w:p${attribute}><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
            })
            .join("");
          return (
            `<w:comment w:id="${commentId(index)}" w:author="${author}" w:initials="${initials}"` +
            ` w:date="2024-01-0${(index % 9) + 1}T00:00:00Z">${body}</w:comment>`
          );
        })
        .join("") +
      "</w:comments>",
  );

  const extendedEntries = forest
    .map(({ parent, done }, index) => {
      const paraId = paraIdFor(forest, index);
      if (paraId === null) {
        return "";
      }
      const parentParaId = parent === null ? null : paraIdFor(forest, parent);
      const parentAttribute = parentParaId === null ? "" : ` w15:paraIdParent="${parentParaId}"`;
      return `<w15:commentEx w15:paraId="${paraId}"${parentAttribute} w15:done="${done ? 1 : 0}"/>`;
    })
    .join("");
  zip.file(
    "word/commentsExtended.xml",
    `${XML_DECLARATION}<w15:commentsEx xmlns:w="${W_NAMESPACE}" xmlns:w15="${W15_NAMESPACE}">${extendedEntries}</w15:commentsEx>`,
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
        '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>' +
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>',
    ),
  );
  zip.file(
    "word/_rels/document.xml.rels",
    rels.replace(
      "</Relationships>",
      '<Relationship Id="rIdPropComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>' +
        '<Relationship Id="rIdPropCommentsEx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>' +
        `<Relationship Id="${HEADER_RELATIONSHIP_ID}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`,
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

/** The plain text of a comment, paragraphs joined so a re-split is not a change. */
const commentText = ({ content }: Comment): string =>
  (content ?? [])
    .map((paragraph) =>
      (paragraph.content ?? [])
        .map((item) =>
          item.type === "run"
            ? (item.content ?? []).map((run) => (run.type === "text" ? run.text : "")).join("")
            : "",
        )
        .join(""),
    )
    .join("\n");

/** Every comment id a range marker or reference in the package addresses. */
const anchoredCommentIds = (document: Document): Record<number, number> => {
  const counts: Record<number, number> = {};
  const walk = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) {
        walk(item);
      }
      return;
    }
    if (typeof value !== "object" || value === null) {
      return;
    }
    const record: Record<string, unknown> = value;
    const { type, id } = record;
    if (
      typeof id === "number" &&
      (type === "commentRangeStart" || type === "commentRangeEnd" || type === "commentReference")
    ) {
      counts[id] = (counts[id] ?? 0) + 1;
    }
    for (const nested of Object.values(record)) {
      walk(nested);
    }
  };
  walk(document.package.document.content);
  walk([...(document.package.headers?.values() ?? [])]);
  walk([...(document.package.footers?.values() ?? [])]);
  return counts;
};

/** Everything a comment is, keyed by the only name the package has for it. */
const identities = (document: Document): Record<number, unknown> =>
  Object.fromEntries(
    (document.package.document.comments ?? []).map((comment) => [
      comment.id,
      {
        author: comment.author,
        initials: comment.initials ?? null,
        date: comment.date ?? null,
        done: comment.done ?? false,
        parentId: comment.parentId ?? null,
        text: commentText(comment),
      },
    ]),
  );

const commentOrder = (document: Document): number[] =>
  (document.package.document.comments ?? []).map(({ id }) => id);

describe("a comment keeps its own identity across a round trip", () => {
  test("parse → save → parse moves no comment's facts onto another id", async () => {
    await fc.assert(
      fc.asyncProperty(commentForest, async (forest) => {
        const parsed = await parseDocx(await buildDocx(forest), { preloadFonts: false });
        const before = {
          identities: identities(parsed),
          order: commentOrder(parsed),
          anchors: anchoredCommentIds(parsed),
        };

        const saved = await repackDocx(parsed, { updateModifiedDate: false });
        const reparsed = await parseDocx(saved, { preloadFonts: false });

        expect({
          identities: identities(reparsed),
          order: commentOrder(reparsed),
          anchors: anchoredCommentIds(reparsed),
        }).toEqual(before);
      }),
      propertyConfig({ numRuns: 40 }),
    );
  });

  test("the editor round trip moves no comment's facts onto another id", async () => {
    await fc.assert(
      fc.asyncProperty(commentForest, async (forest) => {
        const parsed = await parseDocx(await buildDocx(forest), { preloadFonts: false });
        const before = { identities: identities(parsed), order: commentOrder(parsed) };

        const edited = fromProseDoc(toProseDoc(parsed), parsed);
        const saved = await repackDocx(edited, { updateModifiedDate: false });
        const reparsed = await parseDocx(saved, { preloadFonts: false });

        expect({ identities: identities(reparsed), order: commentOrder(reparsed) }).toEqual(before);
      }),
      propertyConfig({ numRuns: 40 }),
    );
  });
});
