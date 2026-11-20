/**
 * A comment's id belongs to its document: the same edits to the same
 * document save the same parts on both engines, whatever else the process
 * allocated before, including comment ids up to the OOXML maximum.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { MAX_REVISION_ID } from "@stll/docx-core/model";
import { ZipArchive } from "@stll/docx-core/zip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { Document } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = ["legacy", "canonical"] as const;
const DATE = "2026-01-02T03:04:05Z";

/** A document holding one comment, with the id given, on its first paragraph. */
const commentedSource = async (commentId: number): Promise<ArrayBuffer> => {
  const document: Document = fromMarkdown("First clause.\n\nSecond clause.");
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") return panic("Expected a paragraph");
  paragraph.content = [
    { type: "commentRangeStart", id: commentId },
    ...paragraph.content,
    { type: "commentRangeEnd", id: commentId },
  ];
  document.package.document.comments = [
    {
      id: commentId,
      author: "Reviewer",
      date: DATE,
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Why?" }] }],
        },
      ],
    },
  ];
  return await createDocx(document);
};

/** Every part's text but `docProps/core.xml`, which records the save time. */
const partsOf = async (bytes: ArrayBuffer): Promise<Map<string, string>> => {
  const zip = await ZipArchive.read(bytes);
  const parts = new Map<string, string>();
  for (const path of zip.paths())
    if (path !== "docProps/core.xml")
      parts.set(path, (await zip.entry(path)?.text()) ?? "");
  return parts;
};

type Edited = { parts: Map<string, string>; commentIds: number[] };

type EditOptions = {
  source: ArrayBuffer;
  session: (typeof SESSIONS)[number];
  /** Whether every comment must be made; at the maximum one may refuse. */
  refusals: "none" | "allowed";
};

/** Comment on every block and reply once, stamped so only ids can vary. */
const edit = async ({
  source,
  session,
  refusals,
}: EditOptions): Promise<Edited> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source.slice(0), {
    session,
    author: "AI",
  });
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: reviewer.snapshot().blocks.map(({ id }, index) => ({
        id: `comment-${String(index)}`,
        type: "commentOnBlock" as const,
        blockId: id,
        comment: { text: `Note ${String(index)}` },
      })),
    },
    { revisionStamp: { date: DATE, idSeed: 100 } },
  );
  if (refusals === "none") {
    expect(result.skipped).toEqual([]);
    const root = reviewer.getComments().at(0) ?? panic("Expected a comment");
    reviewer.replyTo(root.id, { text: "Agreed.", date: DATE });
  }
  const commentIds: number[] = [];
  for (const { id, replies } of reviewer.getComments()) {
    commentIds.push(id);
    for (const reply of replies ?? []) commentIds.push(reply.id);
  }
  return { parts: await partsOf(await reviewer.toBuffer()), commentIds };
};

describe("comment ids per document", () => {
  for (const session of SESSIONS)
    test(`${session}: the same edits save the same parts whatever the process allocated before`, async () => {
      const source = await commentedSource(1);
      const first = await edit({ source, session, refusals: "none" });
      expect(new Set(first.commentIds).size).toBe(first.commentIds.length);
      expect(first.commentIds).toContain(1);

      // Another document allocates up to the maximum, and past it where the
      // engine wraps.
      await edit({
        source: await commentedSource(MAX_REVISION_ID - 1),
        session,
        refusals: "allowed",
      });

      const second = await edit({ source, session, refusals: "none" });
      expect(second.commentIds).toEqual(first.commentIds);
      expect(second.parts).toEqual(first.parts);
    });
});
