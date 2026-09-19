/**
 * A comment's mark keeps its place among the marks beside it.
 *
 * `w:commentReference` is the run Word paints a comment's visible mark with,
 * so the order of the references at one boundary is the order of the marks a
 * reader sees. It is authored data, not a consequence of where the range ends:
 * across the public corpus, of the 153 boundaries where two or more comments
 * close together, 95 interleave each reference with its own end, 8 group every
 * reference after a run of ends, and 50 do neither — all three written by Word
 * itself. So an editor round trip that rebuilds the position from a rule
 * reorders the marks, and it did: `end#0 end#1 ref#0 ref#1` came back as
 * `end#0 ref#0 end#1 ref#1`, and a comment spanning three paragraphs came back
 * painted three times.
 *
 * The example pins the boundary that was found; the property generates the
 * arrangements because the arrangement is the variable. Ends that close at one
 * boundary come back in ascending id order: nothing reads their order (an end
 * paints nothing), and every one of the 72 multi-end runs in the corpus is
 * already ascending, so the generator writes them that way too.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig } from "../../../../test/property-testing";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import { parseDocx } from "./parser";
import { createDocx, createEmptyDocx } from "./rezip";
import type { Document } from "../types/document";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const RELATIONSHIP_NAMESPACE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

type MarkerKind = "start" | "end" | "ref";
type Marker = `${MarkerKind}#${number}`;

const markerXml = (marker: Marker): string => {
  const [kind, rawId] = marker.split("#");
  const id = Number(rawId);
  if (kind === "start") {
    return `<w:commentRangeStart w:id="${id}"/>`;
  }
  if (kind === "end") {
    return `<w:commentRangeEnd w:id="${id}"/>`;
  }
  return `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`;
};

/** One paragraph: text segments with a marker list in each gap between them. */
type ParagraphPlan = {
  segments: readonly string[];
  /** `gaps[i]` sits before `segments[i]`; the last entry is the paragraph tail. */
  gaps: readonly (readonly Marker[])[];
};

const paragraphXml = ({ segments, gaps }: ParagraphPlan): string => {
  const parts: string[] = [];
  for (const [index, segment] of segments.entries()) {
    parts.push(...(gaps[index] ?? []).map(markerXml));
    parts.push(`<w:r><w:t xml:space="preserve">${segment}</w:t></w:r>`);
  }
  parts.push(...(gaps[segments.length] ?? []).map(markerXml));
  return `<w:p>${parts.join("")}</w:p>`;
};

type CommentFacts = {
  id: number;
  author: string;
  text: string;
  parentId: number | null;
};

const paraIdFor = (id: number): string =>
  (0x0000_1000 + id).toString(16).toUpperCase().padStart(8, "0");

const buildDocx = async (
  paragraphs: readonly ParagraphPlan[],
  comments: readonly CommentFacts[],
): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());

  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}<w:document xmlns:w="${W_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">` +
      `<w:body>${paragraphs.map(paragraphXml).join("")}</w:body></w:document>`,
  );
  zip.file(
    "word/comments.xml",
    `${XML_DECLARATION}<w:comments xmlns:w="${W_NAMESPACE}" xmlns:w14="${W14_NAMESPACE}">` +
      comments
        .map(
          ({ id, author, text }) =>
            `<w:comment w:id="${id}" w:author="${author}" w:date="2026-01-0${(id % 9) + 1}T00:00:00Z">` +
            `<w:p w14:paraId="${paraIdFor(id)}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` +
            `</w:comment>`,
        )
        .join("") +
      "</w:comments>",
  );
  zip.file(
    "word/commentsExtended.xml",
    `${XML_DECLARATION}<w15:commentsEx xmlns:w="${W_NAMESPACE}" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">` +
      comments
        .map(({ id, parentId }) => {
          const parent = parentId === null ? "" : ` w15:paraIdParent="${paraIdFor(parentId)}"`;
          return `<w15:commentEx w15:paraId="${paraIdFor(id)}"${parent} w15:done="0"/>`;
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
      `<Relationship Id="rIdPropComments" Type="${RELATIONSHIP_NAMESPACE}/comments" Target="comments.xml"/>` +
        '<Relationship Id="rIdPropCommentsEx" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>',
    ),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const MARKER_PATTERN =
  /<w:(commentRangeStart|commentRangeEnd|commentReference)\b[^>]*\bw:id="(\d+)"/gu;

const MARKER_KIND_BY_ELEMENT = new Map<string, MarkerKind>([
  ["commentRangeStart", "start"],
  ["commentRangeEnd", "end"],
  ["commentReference", "ref"],
]);

/** The marker sequence each body paragraph carries, in document order. */
const markersByParagraph = async (buffer: ArrayBuffer): Promise<Marker[][]> => {
  const zip = await JSZip.loadAsync(buffer);
  const xml = (await zip.file("word/document.xml")?.async("text")) ?? "";
  return xml
    .split(/<w:p[\s>]/u)
    .slice(1)
    .map((paragraph) =>
      [...paragraph.matchAll(MARKER_PATTERN)].map((match): Marker => {
        const kind = MARKER_KIND_BY_ELEMENT.get(match[1] ?? "");
        if (kind === undefined) {
          throw new Error(`Unmatched comment marker element: ${match[0]}`);
        }
        return `${kind}#${Number(match[2])}`;
      }),
    );
};

const editorRoundTrip = async (source: ArrayBuffer): Promise<ArrayBuffer> => {
  const parsed = await parseDocx(source, { preloadFonts: false, detectVariables: false });
  return createDocx(updateDocumentContent(parsed, toProseDoc(parsed)));
};

const commentFacts = (document: Document): Record<number, unknown> =>
  Object.fromEntries(
    (document.package.document.comments ?? []).map((comment) => [
      comment.id,
      {
        author: comment.author,
        parentId: comment.parentId ?? null,
        text: (comment.content ?? [])
          .flatMap((paragraph) => paragraph.content ?? [])
          .flatMap((item) => (item.type === "run" ? (item.content ?? []) : []))
          .map((run) => (run.type === "text" ? run.text : ""))
          .join(""),
      },
    ]),
  );

describe("comment markers keep their authored order through an editor round trip", () => {
  test("two comments closing at one boundary keep their references grouped", async () => {
    const plan: ParagraphPlan = {
      segments: ["shared anchor"],
      gaps: [
        ["start#1", "start#2"],
        ["end#1", "end#2", "ref#1", "ref#2"],
      ],
    };
    const source = await buildDocx(
      [plan],
      [
        { id: 1, author: "Reviewer One", text: "first note", parentId: null },
        { id: 2, author: "Reviewer Two", text: "second note", parentId: null },
      ],
    );

    expect(await markersByParagraph(source)).toEqual([
      ["start#1", "start#2", "end#1", "end#2", "ref#1", "ref#2"],
    ]);
    expect(await markersByParagraph(await editorRoundTrip(source))).toEqual([
      ["start#1", "start#2", "end#1", "end#2", "ref#1", "ref#2"],
    ]);
  });

  test("a comment spanning three paragraphs is marked once", async () => {
    const source = await buildDocx(
      [
        { segments: ["first"], gaps: [["start#1"], []] },
        { segments: ["middle"], gaps: [[], []] },
        { segments: ["last"], gaps: [[], ["end#1", "ref#1"]] },
      ],
      [{ id: 1, author: "Reviewer One", text: "spans the block", parentId: null }],
    );

    const roundTripped = await markersByParagraph(await editorRoundTrip(source));
    expect(roundTripped.flat().filter((marker) => marker === "ref#1")).toHaveLength(1);
    expect(roundTripped.at(-1)?.at(-1)).toBe("ref#1");
  });
});

/** A comment's range and where its reference sits, as gap indexes. */
type GeneratedComment = {
  startGap: number;
  endGap: number;
  /** Gap holding the reference; never before `endGap`. */
  referenceGap: number;
  author: string;
  text: string;
  /** Index of the comment this one replies to, or `null` for a thread root. */
  parent: number | null;
};

const SEGMENT_WORDS = ["alpha", "beta", "gamma", "delta", "epsilon"] as const;

/**
 * Lay a generated comment set out over one paragraph.
 *
 * At a gap the ends go first, in ascending id order, and the references that
 * belong there are merged into them by a generated bit stream: a reference
 * waits for its own end, so the merge produces the interleaved, the grouped
 * and the mixed arrangements the corpus shows. Range starts close the gap,
 * because a range opening where another closes opens after it.
 */
const layOutParagraph = (
  comments: readonly GeneratedComment[],
  gapCount: number,
  merge: readonly boolean[],
): ParagraphPlan => {
  const gaps: Marker[][] = Array.from({ length: gapCount }, () => []);
  let mergeCursor = 0;

  for (const [gap, markers] of gaps.entries()) {
    const ends = comments
      .map((comment, index) => ({ comment, id: index + 1 }))
      .filter(({ comment }) => comment.endGap === gap)
      .toSorted((first, second) => first.id - second.id);
    const references = comments
      .map((comment, index) => ({ comment, id: index + 1 }))
      .filter(({ comment }) => comment.referenceGap === gap);

    const closed = new Set<number>();
    let nextEnd = 0;
    let nextReference = 0;
    while (nextEnd < ends.length || nextReference < references.length) {
      const reference = references[nextReference];
      const referenceReady =
        reference !== undefined && (reference.comment.endGap !== gap || closed.has(reference.id));
      const takeEnd =
        nextEnd < ends.length && (!referenceReady || (merge[mergeCursor++ % merge.length] ?? true));
      if (takeEnd) {
        const end = ends[nextEnd++];
        if (end) {
          closed.add(end.id);
          markers.push(`end#${end.id}`);
        }
        continue;
      }
      if (reference) {
        nextReference += 1;
        markers.push(`ref#${reference.id}`);
      }
    }

    for (const { id } of comments
      .map((comment, index) => ({ comment, id: index + 1 }))
      .filter(({ comment }) => comment.startGap === gap)
      .toSorted((first, second) => first.id - second.id)) {
      markers.push(`start#${id}`);
    }
  }

  return {
    segments: Array.from(
      { length: gapCount - 1 },
      (_, index) => SEGMENT_WORDS[index % SEGMENT_WORDS.length] ?? "word",
    ),
    gaps,
  };
};

const commentSet = fc
  .record({
    segmentCount: fc.integer({ min: 1, max: 4 }),
    rows: fc.array(
      fc.record({
        startOffset: fc.nat({ max: 4 }),
        span: fc.integer({ min: 1, max: 4 }),
        referenceDelay: fc.nat({ max: 3 }),
        author: fc.constantFrom("Reviewer One", "Reviewer Two", "Reviewer Three"),
        text: fc.stringMatching(/^[A-Za-z0-9 ]{1,12}$/u),
        parentOffset: fc.option(fc.nat({ max: 5 }), { nil: null }),
      }),
      { minLength: 2, maxLength: 6 },
    ),
    merge: fc.array(fc.boolean(), { minLength: 1, maxLength: 12 }),
  })
  .map(({ segmentCount, rows, merge }) => {
    const gapCount = segmentCount + 1;
    const comments: GeneratedComment[] = rows.map((row, index) => {
      const startGap = row.startOffset % segmentCount;
      const endGap = Math.min(startGap + row.span, gapCount - 1);
      return {
        startGap,
        endGap,
        referenceGap: Math.min(endGap + row.referenceDelay, gapCount - 1),
        author: row.author,
        text: row.text,
        parent: row.parentOffset === null || index === 0 ? null : row.parentOffset % index,
      };
    });
    return {
      plan: layOutParagraph(comments, gapCount, merge),
      facts: comments.map(
        (comment, index): CommentFacts => ({
          id: index + 1,
          author: comment.author,
          text: comment.text,
          parentId: comment.parent === null ? null : comment.parent + 1,
        }),
      ),
    };
  });

describe("arbitrary comment sets in one paragraph", () => {
  test("keep their marker sequence and every comment's own facts", async () => {
    await fc.assert(
      fc.asyncProperty(commentSet, async ({ plan, facts }) => {
        const source = await buildDocx([plan], facts);
        const authored = await markersByParagraph(source);

        const saved = await editorRoundTrip(source);
        expect(await markersByParagraph(saved)).toEqual(authored);

        const reparsed = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
        expect(commentFacts(reparsed)).toEqual(
          Object.fromEntries(
            facts.map(({ id, author, text, parentId }) => [id, { author, parentId, text }]),
          ),
        );
      }),
      propertyConfig({ numRuns: 60 }),
    );
  });
});
