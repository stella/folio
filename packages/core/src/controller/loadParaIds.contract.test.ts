/**
 * The live editor and the headless reviewer open the same `.docx` with the
 * same block ids, including for paragraphs the package gives no `w14:paraId`,
 * whichever engine each side runs. A host reads ids from the editor and
 * applies edits headlessly, so an id the editor shows has to resolve on the
 * reviewer on the first try.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { prepareCanonicalDocxInput } from "../docx/canonicalSessionInput";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { createHiddenEditorState } from "./hiddenEditorManager";

const DOCUMENTS = {
  clauses: [
    "# Services agreement",
    "1. The supplier provides the services described in the schedule.",
    "2. Either party may terminate on 3 months' notice.",
    "3. This agreement is governed by the laws of the seat.",
  ].join("\n\n"),
  repeatedText: ["Reserved.", "Reserved.", "", "Reserved."].join("\n\n"),
  table: [
    "Payment terms follow.",
    "| Term | Value |\n| --- | --- |\n| Period | 12 months |\n| Fee | 500 |",
    "Signed by both parties.",
  ].join("\n\n"),
} as const;

const docxWithoutParaIds = async (markdown: string): Promise<ArrayBuffer> => {
  const bytes = await createDocx(fromMarkdown(markdown));
  const zip = await ZipArchive.read(bytes);
  const body = await zip.entry("word/document.xml")?.text();
  // The case under test: no paragraph carries an id of its own.
  expect(body).toBeDefined();
  expect(body).not.toContain("w14:paraId");
  return bytes;
};

/** A save stamps the time it was made, whatever it changed. */
const SAVE_STAMPED_PART = "docProps/core.xml";

/** Every part of a package but the save's stamp, by name, as text. */
const partTexts = async (bytes: ArrayBuffer): Promise<Map<string, string>> => {
  const texts = new Map<string, string>();
  for (const entry of (await ZipArchive.read(bytes)).entries())
    if (entry.path !== SAVE_STAMPED_PART)
      texts.set(entry.path, await entry.text());
  return texts;
};

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];
const PAIRS = SESSIONS.flatMap((editor) =>
  SESSIONS.map((reviewer) => ({ editor, reviewer })),
);

/** The editor's model: a canonical editor opens the package it prepares. */
const editorInput = async (
  bytes: ArrayBuffer,
  session: HeadlessReviewerSession,
): Promise<ArrayBuffer | Uint8Array> => {
  if (session === "legacy") return bytes;
  const prepared = await prepareCanonicalDocxInput(bytes);
  if (prepared.isErr()) throw prepared.error;
  return prepared.value;
};

const liveSnapshot = async (
  bytes: ArrayBuffer,
  session: HeadlessReviewerSession = "legacy",
) => {
  const document = await parseDocx(await editorInput(bytes, session), {
    preloadFonts: false,
  });
  const state = createHiddenEditorState({
    document,
    styles: document.package.styles,
    reason: "mount",
  });
  return createFolioAIEditSnapshot(state.doc);
};

const headlessSnapshot = async (
  bytes: ArrayBuffer,
  session: HeadlessReviewerSession,
) =>
  (
    await FolioDocxReviewer.fromBuffer(bytes, { author: "Test", session })
  ).snapshot();

const identities = (blocks: readonly { id: string; idStability?: string }[]) =>
  blocks.map(({ id, idStability }) => ({ id, idStability }));

const ids = (blocks: readonly { id: string }[]) => blocks.map(({ id }) => id);

describe("load-time paragraph ids", () => {
  test.each(
    Object.keys(DOCUMENTS).flatMap((name) =>
      PAIRS.map(({ editor, reviewer }) => ({ name, editor, reviewer })),
    ),
  )(
    "the editor and the headless reviewer derive the same block ids ($name, $editor editor, $reviewer reviewer)",
    async ({ name, editor, reviewer }) => {
      const bytes = await docxWithoutParaIds(
        DOCUMENTS[name as keyof typeof DOCUMENTS],
      );
      const live = await liveSnapshot(bytes, editor);
      const headless = await headlessSnapshot(bytes, reviewer);

      expect(live.blocks.length).toBeGreaterThan(0);
      expect(ids(live.blocks)).toEqual(ids(headless.blocks));
      for (const block of live.blocks) {
        expect(block.id).toMatch(/^[0-9A-F]{8}$/u);
      }
      for (const block of headless.blocks) {
        expect(block.idStability).toBe("positional");
      }
      // A legacy editor reads the package as it is, so it knows which ids it
      // minted; a canonical editor opens the package its session prepared.
      if (editor === "legacy") {
        expect(identities(live.blocks)).toEqual(identities(headless.blocks));
      }
    },
  );

  test("re-opening the same bytes gives the same ids", async () => {
    const bytes = await docxWithoutParaIds(DOCUMENTS.clauses);
    expect(identities((await liveSnapshot(bytes)).blocks)).toEqual(
      identities((await liveSnapshot(bytes)).blocks),
    );
  });

  test.each(
    (["tracked-changes", "direct"] as const).flatMap((mode) =>
      PAIRS.map(({ editor, reviewer }) => ({ mode, editor, reviewer })),
    ),
  )(
    "an edit addressed by an editor block id applies headlessly on the first try ($mode, $editor editor, $reviewer reviewer)",
    async ({ mode, editor, reviewer }) => {
      const bytes = await docxWithoutParaIds(DOCUMENTS.clauses);
      const live = await liveSnapshot(bytes, editor);
      const target = live.blocks.find(({ text }) =>
        text.includes("3 months' notice"),
      );
      expect(target).toBeDefined();
      const blockId = target?.id ?? "";

      const review = await FolioDocxReviewer.fromBuffer(bytes, {
        author: "Test",
        session: reviewer,
      });
      const result = review.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [
            {
              id: "notice",
              type: "replaceInBlock",
              blockId,
              find: "3 months'",
              replace: "60 days'",
            },
            {
              id: "whole",
              type: "replaceBlock",
              blockId: live.blocks.at(-1)?.id ?? "",
              text: "This agreement is governed by the laws of the seat of the supplier.",
            },
          ],
        },
        { undefinedReferences: "refuse" },
      );

      expect(result.skipped).toEqual([]);
      expect(result.applied.map(({ id }) => id).toSorted()).toEqual([
        "notice",
        "whole",
      ]);
      const saved = await FolioDocxReviewer.fromBuffer(
        await review.toBuffer(),
        {
          author: "Test",
          session: reviewer,
        },
      );
      const texts = saved.getContent().map(({ text }) => text);
      if (mode === "direct") {
        expect(texts).toContain(
          "Either party may terminate on 60 days' notice.",
        );
      }
      expect(saved.getChanges().length > 0).toBe(mode === "tracked-changes");
    },
  );

  // The ids are minted the same way on every open, so a save need not write
  // them for a reopen to find them: it writes them only into a part it
  // rewrites anyway, and leaves every other part as the package held it.
  test.each(SESSIONS)(
    "a save writes minted ids only into a part it rewrites (%s)",
    async (session) => {
      const bytes = await docxWithoutParaIds(DOCUMENTS.table);
      const source = await partTexts(bytes);

      const unedited = await FolioDocxReviewer.fromBuffer(bytes, {
        author: "Test",
        session,
      });
      expect(await partTexts(await unedited.toBuffer())).toEqual(source);

      const edited = await FolioDocxReviewer.fromBuffer(bytes, {
        author: "Test",
        session,
      });
      const target = edited.snapshot().blocks.at(0);
      expect(target).toBeDefined();
      const result = edited.applyOperations(
        [
          {
            id: "edit",
            type: "replaceInBlock",
            blockId: target?.id ?? "",
            find: "follow",
            replace: "apply",
          },
        ],
        { undefinedReferences: "refuse", mode: "direct" },
      );
      expect(result.applied).toHaveLength(1);
      const saved = await partTexts(await edited.toBuffer());
      expect(saved.get("word/document.xml")).toContain("apply");
      expect(saved.get("word/document.xml")).not.toContain("follow");
      saved.delete("word/document.xml");
      source.delete("word/document.xml");
      expect(saved).toEqual(source);
    },
  );
});
