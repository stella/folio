/**
 * A comment added to a document is reachable in what it saves: the comment
 * parts exist under their content types, and the document part relates to
 * them, however the opened package held (or lacked) the parts and their
 * relationships. Each package is read back by its relationships alone.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { COMMENT_PART_RELATIONSHIPS } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import {
  commentDocument,
  plainDocument,
} from "../../../../test/consumer-scenarios/support/documents";
import { FolioDocxReviewer } from "../ai-edits/headless";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { resolvePackageRelationshipTarget } from "./packageParts";
import { parseRelationships } from "./relsParser";
import { repackDocx } from "./rezip";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const DOCUMENT_RELS = "word/_rels/document.xml.rels";
const CONTENT_TYPES = "[Content_Types].xml";
const COMMENT_PARTS = [
  "word/comments.xml",
  "word/commentsExtended.xml",
] as const;
const TARGET_WORDS = "The Buyer pays each invoice within thirty days.";
const COMMENT_TEXT = "Which thirty days?";
const REPLY_TEXT = "From receipt.";

const withoutCommentRelationships = (rels: string): string =>
  rels.replaceAll(
    /<Relationship\b[^>]*Target="comments(?:Extended)?\.xml"[^>]*\/>/gu,
    "",
  );

const withoutExtendedRelationship = (rels: string): string =>
  rels.replaceAll(
    /<Relationship\b[^>]*Target="commentsExtended\.xml"[^>]*\/>/gu,
    "",
  );

/** The same relationships or content types under a prefixed root. */
const prefixed = (xml: string): string =>
  xml
    .replaceAll(
      /<(?<close>\/?)(?<name>Relationships?|Types|Override|Default)\b/gu,
      "<$<close>p:$<name>",
    )
    .replaceAll(" xmlns=", " xmlns:p=");

const HYPERLINK =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const LINKS = [
  "https://example.com/comments.xml",
  "https://example.com/commentsExtended.xml",
];
const withLinks = (rels: string): string =>
  rels.replace(
    "</Relationships>",
    `${LINKS.map(
      (target, index) =>
        `<Relationship Id="rId${901 + index}" Type="${HYPERLINK}" Target="${target}" TargetMode="External"/>`,
    ).join("")}</Relationships>`,
  );

type PackageShape = {
  name: string;
  base: () => Promise<Uint8Array>;
  rewrite?: Partial<Record<string, (xml: string) => string | null>>;
};

const SHAPES: readonly PackageShape[] = [
  {
    name: "comment parts without their relationships",
    base: commentDocument,
    rewrite: { [DOCUMENT_RELS]: withoutCommentRelationships },
  },
  {
    name: "a commentsExtended part without its relationship",
    base: commentDocument,
    rewrite: { [DOCUMENT_RELS]: withoutExtendedRelationship },
  },
  { name: "no comment part", base: plainDocument },
  {
    name: "links to files named like the comment parts",
    base: plainDocument,
    rewrite: { [DOCUMENT_RELS]: withLinks },
  },
  {
    name: "a prefixed relationships part",
    base: plainDocument,
    rewrite: { [DOCUMENT_RELS]: prefixed },
  },
  {
    name: "a prefixed content-type registry",
    base: plainDocument,
    rewrite: { [CONTENT_TYPES]: prefixed },
  },
];

const packageOf = async ({
  base,
  rewrite = {},
}: PackageShape): Promise<ArrayBuffer> => {
  const zip = await ZipArchive.read(await base());
  for (const [part, change] of Object.entries(rewrite)) {
    const xml = await zip.entry(part)?.text();
    if (xml === undefined || change === undefined) panic(`no ${part}`);
    const changed = change(xml);
    if (changed === null || changed === xml) panic(`${part} did not change`);
    zip.set(part, new TextEncoder().encode(changed));
  }
  return (await zip.generate({ level: 6 })).slice().buffer;
};

/** The comment parts the saved document part relates to, by path. */
const relatedCommentParts = async (
  saved: ArrayBuffer,
): Promise<Map<string, string>> => {
  const zip = await ZipArchive.read(saved);
  const rels = (await zip.entry(DOCUMENT_RELS)?.text()) ?? "";
  const types = (await zip.entry(CONTENT_TYPES)?.text()) ?? "";
  const parts = new Map<string, string>();
  const ids = new Set<string>();
  for (const [id, { type, target }] of parseRelationships(rels)) {
    expect(ids.has(id)).toBe(false);
    ids.add(id);
    const related = Object.values(COMMENT_PART_RELATIONSHIPS).some(
      (part) => part.type === type,
    );
    if (!related) continue;
    // Package part names compare without case.
    const resolved = resolvePackageRelationshipTarget(target, DOCUMENT_RELS);
    const path = zip.paths().find((name) => name.toLowerCase() === resolved);
    const xml = path === undefined ? undefined : await zip.entry(path)?.text();
    if (path === undefined || xml === undefined)
      panic(`no comment part ${target}`);
    expect(parts.has(path)).toBe(false);
    expect(types).toMatch(
      new RegExp(`<(?:p:)?Override[^>]*PartName="/${path}"`, "u"),
    );
    parts.set(path, xml);
  }
  return parts;
};

const expectReachable = async (
  saved: ArrayBuffer,
  session: HeadlessReviewerSession,
): Promise<void> => {
  const parts = await relatedCommentParts(saved);
  expect([...parts.keys()].toSorted()).toEqual([...COMMENT_PARTS]);
  expect(parts.get("word/comments.xml")).toContain(COMMENT_TEXT);
  expect(parts.get("word/comments.xml")).toContain(REPLY_TEXT);
  expect(parts.get("word/commentsExtended.xml")).toContain("paraIdParent");
  const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
  const thread = reopened
    .getComments()
    .find(({ text }) => text === COMMENT_TEXT);
  expect(thread?.anchoredText).toBe(TARGET_WORDS);
  expect(thread?.replies.map(({ text }) => text)).toEqual([REPLY_TEXT]);
};

const commented = async (
  bytes: ArrayBuffer,
  session: HeadlessReviewerSession,
): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
    session,
    author: "Reviewer",
  });
  const block = reviewer
    .snapshot()
    .blocks.find(({ text }) => text === TARGET_WORDS);
  if (block === undefined) panic("no target paragraph");
  const { skipped } = reviewer.applyOperations(
    [
      {
        id: "comment",
        type: "commentOnBlock",
        blockId: block.id,
        comment: { text: COMMENT_TEXT },
      },
    ],
    { undefinedReferences: "refuse" },
  );
  expect(skipped).toEqual([]);
  const thread = reviewer
    .getComments()
    .find(({ text }) => text === COMMENT_TEXT);
  if (thread === undefined) panic("no comment");
  expect(reviewer.replyTo(thread.id, { text: REPLY_TEXT })).not.toBeNull();
  return reviewer;
};

for (const session of SESSIONS)
  test(`${session}: deleting the comments keeps a link to a file named like their part`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await packageOf({
        name: "threads beside a link",
        base: commentDocument,
        rewrite: { [DOCUMENT_RELS]: withLinks },
      }),
      { session },
    );
    reviewer.deleteComments({
      ids: reviewer.getComments().map(({ id }) => id),
    });
    const zip = await ZipArchive.read(await reviewer.toBuffer());
    const rels = parseRelationships(
      (await zip.entry(DOCUMENT_RELS)?.text()) ?? "",
    );
    expect(
      [...rels.values()].flatMap(({ type, target }) =>
        type === HYPERLINK ? [target] : [],
      ),
    ).toEqual(LINKS);
    expect(
      [...rels.values()].some(
        ({ type }) => type === COMMENT_PART_RELATIONSHIPS.commentsExtended.type,
      ),
    ).toBe(false);
    expect(zip.paths()).not.toContain("word/commentsExtended.xml");
  });

describe("comments added to a package are reachable in what it saves", () => {
  for (const shape of SHAPES)
    for (const session of SESSIONS) {
      test(`${session}: ${shape.name}`, async () => {
        const reviewer = await commented(await packageOf(shape), session);
        const saved = await reviewer.save({ repack: "refuse" });
        if (saved.type !== "selective") panic(`not saved: ${saved.type}`);
        await expectReachable(saved.buffer, session);
      });

      test(`${session}: ${shape.name}, written by a full repack`, async () => {
        const bytes = await packageOf(shape);
        const reviewer = await commented(bytes, session);
        const repacked = await repackDocx({
          ...reviewer.toDocument(),
          originalBuffer: bytes,
        });
        await expectReachable(repacked, session);
      });
    }
});
