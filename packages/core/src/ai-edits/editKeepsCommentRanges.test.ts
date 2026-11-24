/**
 * An edit that replaces a block's text keeps the comment ranges that block
 * anchored, and the link its text carried.
 *
 * The comment shape is the public corpus's `Comment021`: five ranges that all
 * start in the first paragraph and end in later ones. Replacing the first
 * paragraph's text used to drop every `comment` mark it carried, because the
 * mark is `inclusive: false` and the replacement reached the end of what it
 * covered. The ends and references in the later paragraphs stayed, so the save
 * wrote five `commentRangeEnd` elements with no `commentRangeStart`: invalid
 * OOXML, and five comments anchored to nothing. `hyperlink` is non-inclusive
 * for the same reason and lost the same way, leaving prose that had been a
 * link pointing nowhere.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { ZipArchive } from "@stll/docx-core/zip";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type {
  Comment,
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../types/document";
import { FolioDocxReviewer } from "./headless";

setDefaultTimeout(propertyTestTimeout(30_000));

const COMMENT_IDS = [0, 1, 2, 3, 4] as const;
const EDIT_MODES = ["direct", "tracked-changes"] as const;
const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];
/** Paragraph each comment's range ends in; every range starts in paragraph 0. */
const RANGE_END_PARAGRAPH = [1, 1, 2, 3, 3] as const;
const FIRST_PARAGRAPH_TEXT = "First paragraph text.";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const comment = (id: number): Comment => ({
  id,
  author: "Dana Lindqvist",
  initials: "DL",
  date: "2024-01-01T00:00:00Z",
  content: [{ type: "paragraph", content: [run(`Note ${id}`)] }],
});

const buildDocument = (): Document => {
  const opening: ParagraphContent[] = COMMENT_IDS.map((id) => ({
    type: "commentRangeStart",
    id,
  }));
  opening.push(run(FIRST_PARAGRAPH_TEXT));

  const content: Paragraph[] = [
    { type: "paragraph", paraId: "10000001", content: opening },
  ];
  for (let index = 1; index <= 3; index++) {
    const paragraph: ParagraphContent[] = [run(`Paragraph ${index} text.`)];
    for (const id of COMMENT_IDS) {
      if (RANGE_END_PARAGRAPH[id] === index) {
        paragraph.push(
          { type: "commentRangeEnd", id },
          { type: "commentReference", id },
        );
      }
    }
    content.push({
      type: "paragraph",
      paraId: `1000000${index + 1}`,
      content: paragraph,
    });
  }

  return {
    package: { document: { comments: COMMENT_IDS.map(comment), content } },
  };
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await ZipArchive.read(buffer);
  return (await zip.entry("word/document.xml")?.text()) ?? "";
};

const markerIds = (xml: string, marker: string): number[] =>
  [...xml.matchAll(new RegExp(`<w:${marker}\\s[^>]*w:id="(\\d+)"`, "gu"))].map(
    ([, id]) => Number(id),
  );

const replaceFirstParagraph = async (
  session: HeadlessReviewerSession,
  mode: "direct" | "tracked-changes",
): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(buildDocument()),
    { author: "Editor", session },
  );
  const [target] = reviewer.snapshot().blocks;
  expect(target?.text).toBe(FIRST_PARAGRAPH_TEXT);
  const receipt = reviewer.applyOperations(
    [
      {
        id: "r1",
        type: "replaceInBlock",
        blockId: target?.id ?? "",
        find: FIRST_PARAGRAPH_TEXT,
        replace: "Wholly different wording.",
      },
    ],
    { undefinedReferences: "refuse", mode },
  );
  expect(receipt.skipped).toEqual([]);
  return reviewer;
};

const commentEndAndReferenceOrder = (xml: string): string[] =>
  [
    ...xml.matchAll(
      /<w:(commentRangeEnd|commentReference)\b[^>]*\bw:id="(\d+)"/gu,
    ),
  ].map(([, kind, id]) => `${kind}:${id}`);

describe.each(SESSIONS)(
  "%s: replaceInBlock over a block that opens comment ranges",
  (session) => {
    test.each(["direct", "tracked-changes"] as const)(
      "keeps every range's start in %s mode",
      async (mode) => {
        const reviewer = await replaceFirstParagraph(session, mode);
        const buffer = await reviewer.toBuffer();
        const xml = await documentXml(buffer);
        const authoredOrder = COMMENT_IDS.flatMap((id) => [
          `commentRangeEnd:${id}`,
          `commentReference:${id}`,
        ]);
        expect(commentEndAndReferenceOrder(xml)).toEqual(authoredOrder);

        expect(markerIds(xml, "commentRangeStart").toSorted()).toEqual([
          ...COMMENT_IDS,
        ]);
        expect(markerIds(xml, "commentRangeEnd").toSorted()).toEqual([
          ...COMMENT_IDS,
        ]);
        expect(markerIds(xml, "commentReference").toSorted()).toEqual([
          ...COMMENT_IDS,
        ]);

        // The starts move to the beginning of the replacement, which is still the
        // first paragraph: every range covers the new text, as Word's does.
        const [firstParagraph] = xml.split("</w:p>");
        for (const id of COMMENT_IDS) {
          expect(firstParagraph).toContain(
            `<w:commentRangeStart w:id="${id}"/>`,
          );
        }
        const reopened = await FolioDocxReviewer.fromBuffer(buffer, {
          session,
        });
        expect(reopened.snapshot().blocks.at(0)?.text).toBe(
          "Wholly different wording.",
        );
        if (mode === "tracked-changes") {
          reopened.rejectAll();
          const rejectedXml = await documentXml(await reopened.toBuffer());
          expect(rejectedXml).toContain(FIRST_PARAGRAPH_TEXT);
          expect(commentEndAndReferenceOrder(rejectedXml)).toEqual(
            authoredOrder,
          );
        }
      },
    );
  },
);

/**
 * Replacing a block whose comments each cover their own words. Legacy widens
 * every range over the direct replacement and refuses the tracked one, whose
 * rejection would widen them. Canonical deletes the comments a direct
 * replacement covers whole, as the tracked one does once accepted; rejecting
 * it restores every range as authored.
 */
const DISJOINT_REPLACEMENT_OUTCOMES = {
  legacy: { direct: "widened", tracked: "refused" },
  canonical: { direct: "deleted", tracked: "applied" },
} as const satisfies Record<
  HeadlessReviewerSession,
  { direct: "widened" | "deleted"; tracked: "refused" | "applied" }
>;

describe.each(SESSIONS)("%s", (session) => {
  const outcome = DISJOINT_REPLACEMENT_OUTCOMES[session];
  test("tracked replacement never widens disjoint comment ranges on rejection", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(fc.integer({ min: 1, max: 8 }), {
          minLength: 2,
          maxLength: 5,
        }),
        async (lengths) => {
          const content: ParagraphContent[] = [];
          for (const [id, length] of lengths.entries()) {
            content.push(
              { type: "commentRangeStart", id },
              run(String(id).repeat(length)),
              { type: "commentRangeEnd", id },
              { type: "commentReference", id },
              run(" between "),
            );
          }
          const source = await createDocx({
            package: {
              document: {
                comments: lengths.map((_, id) => comment(id)),
                content: [{ type: "paragraph", paraId: "40000001", content }],
              },
            },
          });
          const authored = lengths.map((length, id) =>
            String(id).repeat(length),
          );
          const replaced = async (mode: (typeof EDIT_MODES)[number]) => {
            const reviewer = await FolioDocxReviewer.fromBuffer(source, {
              author: "Editor",
              session,
            });
            const target = reviewer.snapshot().blocks.at(0);
            expect(target).toBeDefined();
            const before = reviewer.toDocument();
            const result = reviewer.applyOperations(
              [
                {
                  id: "replace",
                  type: "replaceBlock",
                  blockId: target?.id ?? "",
                  text: "New clause text.",
                },
              ],
              { undefinedReferences: "refuse", mode },
            );
            if (mode === "tracked-changes" && outcome.tracked === "refused") {
              expect(result.applied).toEqual([]);
              expect(result.skipped).toEqual([
                {
                  id: "replace",
                  reason: "invalidResult",
                  message:
                    "Reject All would change the original content or formatting.",
                },
              ]);
              expect(reviewer.toDocument()).toEqual(before);
            } else {
              expect(result.applied).toHaveLength(1);
            }
            return reviewer.toBuffer();
          };
          const resolvedAnchors = async (
            buffer: ArrayBuffer,
            resolution: "acceptAll" | "rejectAll",
          ) => {
            const saved = await FolioDocxReviewer.fromBuffer(buffer, {
              author: "Editor",
              session,
            });
            saved[resolution]();
            const reopened = await FolioDocxReviewer.fromBuffer(
              await saved.toBuffer(),
              { author: "Editor", session },
            );
            return reopened
              .getComments()
              .map(({ anchoredText }) => anchoredText);
          };

          const direct = await resolvedAnchors(
            await replaced("direct"),
            "acceptAll",
          );
          expect(direct).toEqual(
            outcome.direct === "widened"
              ? lengths.map(() => "New clause text.")
              : [],
          );
          const tracked = await replaced("tracked-changes");
          expect(await resolvedAnchors(tracked, "rejectAll")).toEqual(authored);
          expect(await resolvedAnchors(tracked, "acceptAll")).toEqual(
            outcome.tracked === "refused" ? authored : direct,
          );
        },
      ),
      propertyConfig({ numRuns: 12 }),
    );
  });
});

test.each(["canonical"] as const)(
  "accepting an authored replacement keeps both comment anchors in %s",
  async (session) => {
    const replacement = "New clause text.";
    const source = await createDocx({
      package: {
        document: {
          comments: [comment(0), comment(1)],
          content: [
            {
              type: "paragraph",
              paraId: "40000001",
              content: [
                { type: "commentRangeStart", id: 0 },
                {
                  type: "deletion",
                  info: { id: 1, author: "Editor" },
                  content: [run("0 between ")],
                },
                { type: "commentRangeStart", id: 1 },
                {
                  type: "deletion",
                  info: { id: 2, author: "Editor" },
                  content: [run("1 between ")],
                },
                {
                  type: "insertion",
                  info: { id: 3, author: "Editor" },
                  content: [run(replacement)],
                },
                { type: "commentRangeEnd", id: 1 },
                { type: "commentReference", id: 1 },
                { type: "commentRangeEnd", id: 0 },
                { type: "commentReference", id: 0 },
              ],
            },
          ],
        },
      },
    });
    const reviewer = await FolioDocxReviewer.fromBuffer(source, { session });
    expect(reviewer.acceptAll().resolved).toBeGreaterThan(0);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    expect(reopened.getChanges()).toEqual([]);
    const quotes = reopened
      .getComments()
      .map(({ anchoredText }) => anchoredText);
    expect(quotes).toEqual([replacement, replacement]);
  },
);

const LINK_TEXT = "the standard terms";
const LINK_TARGET = "https://example.invalid/standard-terms";
const REPLACED_LINK_TEXT = "the revised terms of engagement";

const buildLinkedDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "30000001",
          content: [
            {
              type: "hyperlink",
              href: LINK_TARGET,
              tooltip: "Standard terms",
              children: [run(LINK_TEXT)],
            },
          ],
        },
      ],
    },
  },
});

/** The text of each hyperlink in a saved body, and the text outside every link. */
const linkReading = (xml: string) => {
  const texts = (part: string) =>
    [...part.matchAll(/<w:(?:t|delText)\b[^>]*>([^<]*)<\/w:(?:t|delText)>/gu)]
      .map(([, text]) => text ?? "")
      .join("");
  const links = [
    ...xml.matchAll(/<w:hyperlink\b[^>]*>([\s\S]*?)<\/w:hyperlink>/gu),
  ].map(([, inside]) => texts(inside ?? ""));
  return {
    links,
    outside: texts(
      xml.replaceAll(/<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>/gu, ""),
    ),
  };
};

/** Replacements of all of the link's words, and of some of them. */
const LINK_EDITS = [
  {
    name: "its text",
    find: LINK_TEXT,
    replace: REPLACED_LINK_TEXT,
    replaced: REPLACED_LINK_TEXT,
  },
  {
    name: "words inside it",
    find: "standard",
    replace: "revised",
    replaced: "the revised terms",
  },
  {
    name: "its last word",
    find: "terms",
    replace: "rules",
    replaced: "the standard rules",
  },
] as const;

describe.each(["canonical"] as const)(
  "replaceInBlock over a block whose text is a hyperlink (%s)",
  (session) => {
    test.each(
      (["direct", "tracked-changes"] as const).flatMap((mode) =>
        LINK_EDITS.map(({ name, find, replace, replaced }) => ({
          mode,
          name,
          find,
          replace,
          replaced,
        })),
      ),
    )(
      "keeps one link on the new text replacing $name in $mode mode",
      async ({ mode, find, replace, replaced }) => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await createDocx(buildLinkedDocument()),
          { author: "Editor", session },
        );
        const [target] = reviewer.snapshot().blocks;
        expect(target?.text).toBe(LINK_TEXT);
        const receipt = reviewer.applyOperations(
          [
            {
              id: "r1",
              type: "replaceInBlock",
              blockId: target?.id ?? "",
              find,
              replace,
            },
          ],
          { undefinedReferences: "refuse", mode },
        );
        expect(receipt.skipped).toEqual([]);

        const saved = await reviewer.toBuffer();
        const xml = await documentXml(saved);
        const relationships = await (async () => {
          const zip = await ZipArchive.read(saved);
          return (
            (await zip.entry("word/_rels/document.xml.rels")?.text()) ?? ""
          );
        })();

        // The link still wraps the replacement, and still points where it
        // did; nothing outside a link carries the new text.
        const reading = linkReading(xml);
        expect(reading.links.join("")).toContain(replace);
        expect(reading.outside).toBe("");
        expect(relationships).toContain(LINK_TARGET);
        if (mode === "direct") {
          expect(reading.links).toEqual([replaced]);
          return;
        }

        // Accepting leaves the new text in the one link, rejecting the old.
        for (const [resolve, kept] of [
          ["accept", replaced],
          ["reject", LINK_TEXT],
        ] as const) {
          const reopened = await FolioDocxReviewer.fromBuffer(saved, {
            session,
          });
          if (resolve === "accept") reopened.acceptAll();
          else reopened.rejectAll();
          const resolved = await FolioDocxReviewer.fromBuffer(
            await reopened.toBuffer(),
            { session },
          );
          expect(resolved.getChanges()).toEqual([]);
          expect(
            linkReading(await documentXml(await resolved.toBuffer())),
          ).toEqual({ links: [kept], outside: "" });
        }
      },
    );
  },
);
