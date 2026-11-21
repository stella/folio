/**
 * Comment packages both engines open by repairing them: range markers that
 * name no comment or miss a boundary, and comment paragraphs repeating an id.
 * Each opens on the legacy and the canonical session, reports its repair,
 * keeps every comment's text and is stable across a save and a reopen.
 */
import { describe, expect, test } from "bun:test";

import {
  PARSE_WARNING_CODES,
  type ParseWarningCode,
} from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import { commentDocument } from "../../../../test/consumer-scenarios/support/documents";
import { FolioDocxReviewer } from "../ai-edits/headless";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { openCanonicalDocx } from "./canonicalSessionInput";
import { parseDocx } from "./parser";

const DOCUMENT_PART = "word/document.xml";
const COMMENTS_PART = "word/comments.xml";
const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

/** The fixture's comment `w:id`s and paragraph ids, as `commentDocument` writes them. */
const RESOLVED_COMMENT_ID = "1";
const MISSING_COMMENT_ID = "47";
const RESOLVED_COMMENT_PARA_ID = "63BD4A7E";

type RepairShape = {
  name: string;
  part: string;
  rewrite: (xml: string) => string;
  repairs: readonly ParseWarningCode[];
};

const replaceOnce = (xml: string, find: string, replace: string): string => {
  const at = xml.indexOf(find);
  if (at < 0) throw new Error(`fixture has no ${find}`);
  return `${xml.slice(0, at)}${replace}${xml.slice(at + find.length)}`;
};

const marker = (type: string, id: string): string =>
  `<w:${type} w:id="${id}"/>`;

const SHAPES: readonly RepairShape[] = [
  {
    name: "a range start naming no comment",
    part: DOCUMENT_PART,
    rewrite: (xml) =>
      replaceOnce(
        xml,
        marker("commentRangeStart", RESOLVED_COMMENT_ID),
        marker("commentRangeStart", MISSING_COMMENT_ID),
      ),
    repairs: [
      PARSE_WARNING_CODES.danglingCommentReference,
      PARSE_WARNING_CODES.unpairedCommentRangeBoundary,
    ],
  },
  {
    name: "a reference naming no comment",
    part: DOCUMENT_PART,
    rewrite: (xml) =>
      replaceOnce(
        xml,
        marker("commentReference", RESOLVED_COMMENT_ID),
        marker("commentReference", MISSING_COMMENT_ID),
      ),
    repairs: [PARSE_WARNING_CODES.danglingCommentReference],
  },
  {
    name: "a range start without its end",
    part: DOCUMENT_PART,
    rewrite: (xml) =>
      replaceOnce(xml, marker("commentRangeEnd", RESOLVED_COMMENT_ID), ""),
    repairs: [PARSE_WARNING_CODES.unpairedCommentRangeBoundary],
  },
  {
    name: "a range end without its start",
    part: DOCUMENT_PART,
    rewrite: (xml) =>
      replaceOnce(xml, marker("commentRangeStart", RESOLVED_COMMENT_ID), ""),
    repairs: [PARSE_WARNING_CODES.unpairedCommentRangeBoundary],
  },
  {
    name: "a comment repeating its reference",
    part: DOCUMENT_PART,
    rewrite: (xml) =>
      replaceOnce(
        xml,
        marker("commentReference", RESOLVED_COMMENT_ID),
        `${marker("commentReference", RESOLVED_COMMENT_ID)}</w:r><w:r>${marker("commentReference", RESOLVED_COMMENT_ID)}`,
      ),
    repairs: [PARSE_WARNING_CODES.duplicateCommentReference],
  },
  {
    name: "a range boundary between a table's rows and one between a row's cells, each without its partner",
    part: DOCUMENT_PART,
    rewrite: (xml) => {
      const cell =
        '<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Cell</w:t></w:r></w:p></w:tc>';
      const table = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr>${cell}</w:tr>${marker("commentRangeEnd", RESOLVED_COMMENT_ID)}<w:tr>${cell}${marker("commentRangeStart", "2")}</w:tr></w:tbl><w:p/>`;
      return replaceOnce(xml, "<w:sectPr", `${table}<w:sectPr`);
    },
    repairs: [PARSE_WARNING_CODES.unpairedCommentRangeBoundary],
  },
  {
    name: "a comment repeating its own paragraph id",
    part: COMMENTS_PART,
    rewrite: (xml) => {
      const paragraph = new RegExp(
        `<w:p w14:paraId="${RESOLVED_COMMENT_PARA_ID}">[\\s\\S]*?</w:p>`,
        "u",
      ).exec(xml)?.[0];
      if (paragraph === undefined) throw new Error("fixture has no paragraph");
      return replaceOnce(xml, paragraph, `${paragraph}${paragraph}`);
    },
    repairs: [PARSE_WARNING_CODES.repeatedCommentParagraphId],
  },
  {
    name: "a comment repeating another comment's paragraph id",
    part: COMMENTS_PART,
    // A first paragraph, so the comment's own thread id stays its last one.
    rewrite: (xml) => {
      const open = /<w:comment w:id="2"[^>]*>/u.exec(xml)?.[0];
      if (open === undefined) throw new Error("fixture has no comment 2");
      return replaceOnce(
        xml,
        open,
        `${open}<w:p w14:paraId="${RESOLVED_COMMENT_PARA_ID}"><w:r><w:t>See clause 4.</w:t></w:r></w:p>`,
      );
    },
    repairs: [PARSE_WARNING_CODES.repeatedCommentParagraphId],
  },
];

const fixture = commentDocument();

const rewritten = async ({
  part,
  rewrite,
}: RepairShape): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await fixture);
  const entry = zip.entry(part);
  if (!entry) throw new Error(`fixture has no ${part}`);
  const xml = new TextDecoder().decode(await entry.bytes());
  const changed = rewrite(xml);
  if (changed === xml) throw new Error("rewrite made no change");
  zip.set(part, new TextEncoder().encode(changed));
  return (await zip.generate({ level: 6 })).slice().buffer;
};

const commentRead = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getComments()
    .map(({ author, text, anchoredText, done, replies }) => ({
      author,
      text,
      anchoredText,
      done,
      replies: replies.map((reply) => reply.text),
    }));

const commentTexts = (reviewer: FolioDocxReviewer): string[] => {
  const texts: string[] = [];
  for (const { text, replies } of reviewer.getComments()) {
    texts.push(text);
    for (const reply of replies) texts.push(reply.text);
  }
  return texts;
};

const warningCodes = (warnings: readonly { code: ParseWarningCode }[] = []) =>
  new Set(warnings.map(({ code }) => code));

describe("comment packages repaired on open", () => {
  test("the fixture opens on both sessions as it is", async () => {
    const bytes = (await fixture).slice().buffer;
    for (const session of SESSIONS) {
      // oxlint-disable-next-line no-await-in-loop -- sessions are compared in order
      const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { session });
      expect(commentTexts(reviewer).sort()).toEqual([
        "Define good order.",
        "Good order means undamaged.",
        "Thirty days is long.",
      ]);
    }
  });

  for (const shape of SHAPES) {
    test(`${shape.name}: both engines read the repair`, async () => {
      const bytes = await rewritten(shape);
      const legacy = await parseDocx(bytes, {
        preloadFonts: false,
        detectVariables: false,
      });
      const canonical = await openCanonicalDocx(bytes, {
        preloadFonts: false,
        detectVariables: false,
      });
      if (canonical.isErr()) throw canonical.error;
      for (const code of shape.repairs) {
        expect(warningCodes(legacy.parseWarnings)).toContain(code);
        expect(warningCodes(canonical.value.document.parseWarnings)).toContain(
          code,
        );
      }
      const paraIds = (comments = legacy.package.document.comments ?? []) =>
        comments.flatMap(({ content }) => content.map(({ paraId }) => paraId));
      expect(
        paraIds(canonical.value.document.package.document.comments),
      ).toEqual(paraIds());
      expect(new Set(paraIds()).size).toBe(paraIds().length);
    });

    for (const session of SESSIONS) {
      test(`${shape.name}: opens on ${session}, keeps comment text and reopens stable`, async () => {
        const bytes = await rewritten(shape);
        const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { session });
        const texts = commentTexts(reviewer);
        for (const text of [
          "Define good order.",
          "Good order means undamaged.",
          "Thirty days is long.",
        ])
          expect(texts.some((held) => held.includes(text))).toBe(true);
        const before = commentRead(reviewer);

        const saved = await reviewer.toBuffer();
        const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
        expect(commentRead(reopened)).toEqual(before);
        const again = await FolioDocxReviewer.fromBuffer(
          await reopened.toBuffer(),
          { session },
        );
        expect(commentRead(again)).toEqual(before);
      });
    }
  }
});
