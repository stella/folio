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
 *
 * A range is a story fact rather than a paragraph one, so the same
 * arrangements are also laid out across several paragraphs: the marker
 * sequence has to survive that, and every paragraph strictly inside a range
 * has to carry the mark, which is the highlight between the two boundaries.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import type { Node as PMNode } from "prosemirror-model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { updateDocumentContent } from "../prosemirror/conversion/fromProseDoc";
import { parseDocx } from "./parser";
import { createDocx, createEmptyDocx } from "./rezip";
import type { Document } from "../types/document";

setDefaultTimeout(propertyTestTimeout(30_000));

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

/** Body paragraph indexes whose every text node carries `commentId`'s mark. */
const fullyMarkedParagraphs = (pmDoc: PMNode, commentId: number): Set<number> => {
  const marked = new Set<number>();
  for (let index = 0; index < pmDoc.childCount; index++) {
    const paragraph = pmDoc.child(index);
    let texts = 0;
    let carrying = 0;
    for (let child = 0; child < paragraph.childCount; child++) {
      const node = paragraph.child(child);
      if (!node.isText) {
        continue;
      }
      texts += 1;
      if (
        node.marks.some(
          (mark) => mark.type.name === "comment" && mark.attrs["commentId"] === commentId,
        )
      ) {
        carrying += 1;
      }
    }
    if (texts > 0 && texts === carrying) {
      marked.add(index);
    }
  }
  return marked;
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

  test("a comment spanning three paragraphs keeps one range over all three", async () => {
    const authored: Marker[][] = [["start#1"], [], ["end#1", "ref#1"]];
    const source = await buildDocx(
      [
        { segments: ["first"], gaps: [["start#1"], []] },
        { segments: ["middle"], gaps: [[], []] },
        { segments: ["last"], gaps: [[], ["end#1", "ref#1"]] },
      ],
      [{ id: 1, author: "Reviewer One", text: "spans the block", parentId: null }],
    );

    expect(await markersByParagraph(source)).toEqual(authored);

    // The paragraph between the two boundaries is inside the range, so its
    // text carries the mark: the highlight the reader sees is unbroken.
    const parsed = await parseDocx(source, { preloadFonts: false, detectVariables: false });
    expect(fullyMarkedParagraphs(toProseDoc(parsed), 1)).toEqual(new Set([0, 1, 2]));

    expect(await markersByParagraph(await editorRoundTrip(source))).toEqual(authored);
  });
});

/**
 * A comment's range and where its reference sits, as story slots.
 *
 * A slot is a gap in a paragraph, numbered over the whole story: paragraph
 * `Math.floor(slot / gapCount)`, gap `slot % gapCount`. A range is a story
 * fact, so its two ends need not share a paragraph.
 */
type GeneratedComment = {
  startSlot: number;
  endSlot: number;
  /** Slot holding the reference; never before `endSlot`. */
  referenceSlot: number;
  author: string;
  text: string;
  /** Index of the comment this one replies to, or `null` for a thread root. */
  parent: number | null;
};

const SEGMENT_WORDS = ["alpha", "beta", "gamma", "delta", "epsilon"] as const;

/**
 * Lay a generated comment set out over a story's paragraphs.
 *
 * At a slot the ends go first, in ascending id order, and the references that
 * belong there are merged into them by a generated bit stream: a reference
 * waits for its own end, so the merge produces the interleaved, the grouped
 * and the mixed arrangements the corpus shows. Range starts close the slot,
 * because a range opening where another closes opens after it.
 */
const layOutStory = (
  comments: readonly GeneratedComment[],
  paragraphCount: number,
  gapCount: number,
  merge: readonly boolean[],
): ParagraphPlan[] => {
  const plans: ParagraphPlan[] = [];
  let mergeCursor = 0;

  for (let paragraph = 0; paragraph < paragraphCount; paragraph++) {
    const gaps: Marker[][] = Array.from({ length: gapCount }, () => []);

    for (const [gap, markers] of gaps.entries()) {
      const slot = paragraph * gapCount + gap;
      const ends = comments
        .map((comment, index) => ({ comment, id: index + 1 }))
        .filter(({ comment }) => comment.endSlot === slot)
        .toSorted((first, second) => first.id - second.id);
      const references = comments
        .map((comment, index) => ({ comment, id: index + 1 }))
        .filter(({ comment }) => comment.referenceSlot === slot);

      const closed = new Set<number>();
      let nextEnd = 0;
      let nextReference = 0;
      while (nextEnd < ends.length || nextReference < references.length) {
        const reference = references[nextReference];
        const referenceReady =
          reference !== undefined &&
          (reference.comment.endSlot !== slot || closed.has(reference.id));
        const takeEnd =
          nextEnd < ends.length &&
          (!referenceReady || (merge[mergeCursor++ % merge.length] ?? true));
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
        .filter(({ comment }) => comment.startSlot === slot)
        .toSorted((first, second) => first.id - second.id)) {
        markers.push(`start#${id}`);
      }
    }

    plans.push({
      segments: Array.from(
        { length: gapCount - 1 },
        (_, index) => SEGMENT_WORDS[index % SEGMENT_WORDS.length] ?? "word",
      ),
      gaps,
    });
  }

  return plans;
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
      const startSlot = row.startOffset % segmentCount;
      const endSlot = Math.min(startSlot + row.span, gapCount - 1);
      return {
        startSlot,
        endSlot,
        referenceSlot: Math.min(endSlot + row.referenceDelay, gapCount - 1),
        author: row.author,
        text: row.text,
        parent: row.parentOffset === null || index === 0 ? null : row.parentOffset % index,
      };
    });
    return {
      plans: layOutStory(comments, 1, gapCount, merge),
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
      fc.asyncProperty(commentSet, async ({ plans, facts }) => {
        const source = await buildDocx(plans, facts);
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

/**
 * The same arrangements, spread over up to four paragraphs.
 *
 * A range starts before some text in its paragraph and ends after some, so the
 * generator keeps a start out of a paragraph's tail gap and an end out of gap
 * 0: a boundary with no text beside it in its own paragraph is the same
 * document as one at the neighbouring paragraph's edge, and an editor that
 * carries the range as a mark cannot tell the two apart.
 */
const spanningCommentSet = fc
  .record({
    paragraphCount: fc.integer({ min: 1, max: 4 }),
    segmentCount: fc.integer({ min: 1, max: 3 }),
    rows: fc.array(
      fc.record({
        startOffset: fc.nat({ max: 11 }),
        span: fc.integer({ min: 1, max: 8 }),
        referenceDelay: fc.nat({ max: 3 }),
        author: fc.constantFrom("Reviewer One", "Reviewer Two", "Reviewer Three"),
        text: fc.stringMatching(/^[A-Za-z0-9 ]{1,12}$/u),
        parentOffset: fc.option(fc.nat({ max: 5 }), { nil: null }),
      }),
      { minLength: 1, maxLength: 5 },
    ),
    merge: fc.array(fc.boolean(), { minLength: 1, maxLength: 12 }),
  })
  .map(({ paragraphCount, segmentCount, rows, merge }) => {
    const gapCount = segmentCount + 1;
    const lastSlot = paragraphCount * gapCount - 1;
    const comments: GeneratedComment[] = rows.map((row, index) => {
      const startParagraph = Math.floor(row.startOffset / segmentCount) % paragraphCount;
      const startSlot = startParagraph * gapCount + (row.startOffset % segmentCount);
      const candidateEnd = Math.min(startSlot + row.span, lastSlot);
      // A range ending at gap 0 closes before its paragraph's first word,
      // which is the previous paragraph's tail written differently.
      const endSlot = candidateEnd % gapCount === 0 ? candidateEnd + 1 : candidateEnd;
      return {
        startSlot,
        endSlot,
        referenceSlot: Math.min(endSlot + row.referenceDelay, lastSlot),
        author: row.author,
        text: row.text,
        parent: row.parentOffset === null || index === 0 ? null : row.parentOffset % index,
      };
    });
    return {
      gapCount,
      plans: layOutStory(comments, paragraphCount, gapCount, merge),
      comments,
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

describe("arbitrary comment sets spanning several paragraphs", () => {
  test("keep one range over the paragraphs it covers, and mark every one of them", async () => {
    await fc.assert(
      fc.asyncProperty(spanningCommentSet, async ({ comments, facts, gapCount, plans }) => {
        const source = await buildDocx(plans, facts);
        const authored = await markersByParagraph(source);

        const parsed = await parseDocx(source, { preloadFonts: false, detectVariables: false });
        const pmDoc = toProseDoc(parsed);
        // A paragraph strictly inside a range holds no boundary of its own, so
        // all of its text is covered: this is the middle of the highlight the
        // reader sees.
        for (const [index, { startSlot, endSlot }] of comments.entries()) {
          const marked = fullyMarkedParagraphs(pmDoc, index + 1);
          for (
            let paragraph = Math.floor(startSlot / gapCount) + 1;
            paragraph < Math.floor(endSlot / gapCount);
            paragraph++
          ) {
            expect(marked).toContain(paragraph);
          }
        }

        expect(await markersByParagraph(await editorRoundTrip(source))).toEqual(authored);
      }),
      propertyConfig({ numRuns: 60 }),
    );
  });
});
