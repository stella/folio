/**
 * Where a missing `w:commentReference` is filled in, and where it is not.
 *
 * Both engines read the repair at open: the parser adds the reference after
 * the range end in the model and reports it. A save writes it only into a
 * paragraph the save rewrites anyway, as with minted paragraph ids, so an
 * unedited save stays byte-identical. Markdown has no comment syntax to arrive
 * with a half-written comment through, and the serializer writes the model as
 * authored.
 */

import { describe, expect, test } from "bun:test";

import { PARSE_WARNING_CODES } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { toMarkdown } from "../markdown";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { Comment, Document, ParagraphContent } from "../types/document";
import { openCanonicalDocx } from "./canonicalSessionInput";
import { parseDocx } from "./parser";
import { createDocx } from "./rezip";

const COMMENT_ID = 0;
const COMMENTED_TEXT = "Commented body text.";
const OTHER_TEXT = "Second body paragraph.";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const comment: Comment = {
  id: COMMENT_ID,
  author: "Dana Lindqvist",
  initials: "DL",
  date: "2024-01-01T00:00:00Z",
  content: [{ type: "paragraph", content: [run("Check this wording.")] }],
};

/** A range whose comment has no `commentReference` anywhere in the story. */
const referencelessDocument = (): Document => ({
  package: {
    document: {
      comments: [comment],
      content: [
        {
          type: "paragraph",
          paraId: "61000000",
          content: [
            { type: "commentRangeStart", id: COMMENT_ID },
            run(COMMENTED_TEXT),
            { type: "commentRangeEnd", id: COMMENT_ID },
          ],
        },
        {
          type: "paragraph",
          paraId: "61000001",
          content: [run(OTHER_TEXT)],
        },
      ],
    },
  },
});

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const documentXml = async (buffer: ArrayBuffer): Promise<string> =>
  (await (await ZipArchive.read(buffer)).entry("word/document.xml")?.text()) ??
  "";

const referenceCount = (xml: string): number =>
  xml.split("<w:commentReference").length - 1;

/** The mark is painted where the reference sits, so it must follow the end. */
const expectReferenceAfterEnd = (xml: string): void => {
  expect(referenceCount(xml)).toBe(1);
  expect(xml.indexOf("<w:commentReference")).toBeGreaterThan(
    xml.indexOf("<w:commentRangeEnd"),
  );
};

const replaceParagraph = async (
  session: HeadlessReviewerSession,
  text: string,
): Promise<string> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await createDocx(referencelessDocument()),
    { author: "Editor", session },
  );
  const target = reviewer
    .snapshot()
    .blocks.find((block) => block.text === text);
  expect(target).toBeDefined();
  const result = reviewer.applyOperations(
    [
      {
        id: "edit",
        type: "replaceInBlock",
        blockId: target?.id ?? "",
        find: text.slice(0, -1),
        replace: "Superseded wording",
      },
    ],
    { undefinedReferences: "refuse", mode: "direct" },
  );
  expect(result.applied).toHaveLength(1);
  return documentXml(await reviewer.toBuffer());
};

describe("completing a missing comment reference", () => {
  test("both engines add the reference at open and report it", async () => {
    const source = await createDocx(referencelessDocument());
    expect(referenceCount(await documentXml(source))).toBe(0);

    const legacy = await parseDocx(source, {
      preloadFonts: false,
      detectVariables: false,
    });
    const canonical = await openCanonicalDocx(source, {
      preloadFonts: false,
      detectVariables: false,
    });
    if (canonical.isErr()) throw canonical.error;
    for (const parsed of [legacy, canonical.value.document]) {
      expect(
        parsed.parseWarnings?.filter(
          ({ code }) => code === PARSE_WARNING_CODES.missingCommentReference,
        ),
      ).toEqual([expect.objectContaining({ count: 1 })]);
      expect(parsed.package.document.content.at(0)).toMatchObject({
        content: [
          { type: "commentRangeStart", id: COMMENT_ID },
          { type: "run" },
          { type: "commentRangeEnd", id: COMMENT_ID },
          { type: "commentReference", id: COMMENT_ID },
        ],
      });
    }
  });

  for (const session of SESSIONS) {
    test(`${session}: an unedited save keeps the document part byte-identical`, async () => {
      const source = await createDocx(referencelessDocument());
      const reviewer = await FolioDocxReviewer.fromBuffer(source, { session });
      expect(
        reviewer.getComments().map(({ anchoredText }) => anchoredText),
      ).toEqual([COMMENTED_TEXT]);
      expect(await documentXml(await reviewer.toBuffer())).toBe(
        await documentXml(source),
      );
    });

    test(`${session}: editing the commented paragraph writes its reference`, async () => {
      expectReferenceAfterEnd(await replaceParagraph(session, COMMENTED_TEXT));
    });

    // The commented paragraph is not rewritten, so it keeps its bytes.
    test(`${session}: editing another paragraph leaves the commented one as authored`, async () => {
      const saved = await replaceParagraph(session, OTHER_TEXT);
      expect(referenceCount(saved)).toBe(0);
      expect(saved).toContain('<w:commentRangeEnd w:id="0"/>');
    });
  }

  // The markdown reader has no comment production at all, so an import cannot
  // arrive with a range at all, let alone one missing its reference.
  test("a markdown import carries no comment marker to complete", () => {
    const exported = toMarkdown(
      {
        package: {
          document: {
            comments: [comment],
            content: [
              {
                type: "paragraph",
                content: [
                  { type: "commentRangeStart", id: COMMENT_ID },
                  run("Commented body text."),
                  { type: "commentRangeEnd", id: COMMENT_ID },
                  { type: "commentReference", id: COMMENT_ID },
                ],
              },
            ],
          },
        },
      },
      { annotations: "html" },
    );

    const imported = fromMarkdown(exported);
    expect(imported.package.document.comments).toBeUndefined();
    expect(JSON.stringify(imported.package)).not.toContain("commentRange");
  });

  // Completion belongs to the open, which reports it; the serializer writes
  // the model it is given.
  test("the save path leaves an authored model as it found it", async () => {
    const authored = referencelessDocument();
    const saved = await documentXml(await createDocx(authored));

    expect(referenceCount(saved)).toBe(0);
    expect(saved).toContain('<w:commentRangeStart w:id="0"/>');
    expect(saved).toContain('<w:commentRangeEnd w:id="0"/>');
  });
});
