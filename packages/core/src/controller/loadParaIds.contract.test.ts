/**
 * The live editor and the headless reviewer open the same `.docx` with the
 * same block ids, including for paragraphs the package gives no `w14:paraId`.
 * A host reads ids from
 * the editor and applies edits headlessly, so an id the editor shows has to
 * resolve on the reviewer on the first try.
 */
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { parseDocx } from "../docx/parser";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
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

const liveSnapshot = async (bytes: ArrayBuffer) => {
  const document = await parseDocx(bytes, { preloadFonts: false });
  const state = createHiddenEditorState({
    document,
    styles: document.package.styles,
    reason: "mount",
  });
  return createFolioAIEditSnapshot(state.doc);
};

const headlessSnapshot = async (bytes: ArrayBuffer) =>
  (await FolioDocxReviewer.fromBuffer(bytes, { author: "Test" })).snapshot();

const identities = (blocks: readonly { id: string; idStability?: string }[]) =>
  blocks.map(({ id, idStability }) => ({ id, idStability }));

describe("load-time paragraph ids", () => {
  test.each(Object.entries(DOCUMENTS))(
    "the editor and the headless reviewer derive the same block ids (%s)",
    async (_name, markdown) => {
      const bytes = await docxWithoutParaIds(markdown);
      const live = await liveSnapshot(bytes);
      const headless = await headlessSnapshot(bytes);

      expect(live.blocks.length).toBeGreaterThan(0);
      expect(identities(live.blocks)).toEqual(identities(headless.blocks));
      for (const block of live.blocks) {
        expect(block.id).toMatch(/^[0-9A-F]{8}$/u);
        expect(block.idStability).toBe("positional");
      }
    },
  );

  test("re-opening the same bytes gives the same ids", async () => {
    const bytes = await docxWithoutParaIds(DOCUMENTS.clauses);
    expect(identities((await liveSnapshot(bytes)).blocks)).toEqual(
      identities((await liveSnapshot(bytes)).blocks),
    );
  });

  test.each(["tracked-changes", "direct"] as const)(
    "an edit addressed by an editor block id applies headlessly on the first try (%s)",
    async (mode) => {
      const bytes = await docxWithoutParaIds(DOCUMENTS.clauses);
      const live = await liveSnapshot(bytes);
      const target = live.blocks.find(({ text }) =>
        text.includes("3 months' notice"),
      );
      expect(target).toBeDefined();
      const blockId = target?.id ?? "";

      const reviewer = await FolioDocxReviewer.fromBuffer(bytes, {
        author: "Test",
      });
      const result = reviewer.applyDocumentOperations(
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
        await reviewer.toBuffer(),
        { author: "Test" },
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
});
