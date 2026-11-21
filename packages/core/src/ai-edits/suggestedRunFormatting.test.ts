/**
 * A suggestion formatting a run whose properties already carry a pending
 * change that is not the suggestion's own. A run holds one property change
 * and a suggestion is resolved by its id, so folding into the other change
 * would let rejecting the suggestion discard it. Both engines refuse, and
 * the pending change stays as it was through save and rejection.
 */
import { panic } from "better-result";
import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIInlineFormatting } from "./types";

const SESSIONS = ["legacy", "canonical"] as const;

/** How each engine reports the refusal: by the same reason. */
const REFUSAL = {
  legacy: { reason: "pendingRunPropertyChange" },
  canonical: {
    reason: "pendingRunPropertyChange",
    message:
      "A suggestion cannot take over another pending run property change.",
  },
} as const satisfies Record<(typeof SESSIONS)[number], object>;

/** "Reviewer" made the first word bold, as tracked change 7. */
const PENDING_BOLD_RUN =
  '<w:r><w:rPr><w:b/><w:rPrChange w:id="7" w:author="Reviewer" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>Formatting</w:t></w:r>';

const source = async (): Promise<ArrayBuffer> => {
  const document: Document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "A1000001",
      content: [
        {
          type: "run",
          formatting: { bold: true },
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: { id: 7, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
              previousFormatting: {},
            },
          ],
          content: [{ type: "text", text: "Formatting" }],
        },
        { type: "run", content: [{ type: "text", text: " target" }] },
      ],
    },
  ];
  return createDocx(document);
};

const documentXml = async (bytes: ArrayBuffer): Promise<string> =>
  (await (await ZipArchive.read(bytes)).entry("word/document.xml")?.text()) ??
  panic("Expected word/document.xml");

type SuggestOptions = {
  reviewer: FolioDocxReviewer;
  id: string;
  formatting: FolioAIInlineFormatting;
};

const suggestFirstWord = ({ reviewer, id, formatting }: SuggestOptions) => {
  const block = reviewer.snapshot().blocks.at(0) ?? panic("Expected a block");
  const range =
    createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 0,
      endOffset: "Formatting".length,
    }) ?? panic("Expected a range");
  return reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [{ id, type: "formatRange", range, formatting }],
    },
    { undefinedReferences: "refuse" },
  );
};

describe.each(SESSIONS)(
  "%s: a suggestion over a run's pending change",
  (session) => {
    test("is refused, and another author's tracked change stays as it was", async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
        author: "Editor",
        session,
      });
      const result = suggestFirstWord({
        reviewer,
        id: "italic",
        formatting: { italic: true },
      });
      expect(result.applied).toEqual([]);
      expect(result.skipped).toEqual([
        expect.objectContaining({ id: "italic", ...REFUSAL[session] }),
      ]);
      // A deliberate refusal names no missing capability.
      expect(result.skipped.at(0)).not.toHaveProperty("canonicalRefusal");
      expect(reviewer.getChanges()).toEqual([
        expect.objectContaining({
          id: 7,
          type: "formatting",
          author: "Reviewer",
        }),
      ]);
      expect(await documentXml(await reviewer.toBuffer())).toContain(
        PENDING_BOLD_RUN,
      );
      expect(reviewer.rejectSuggestion("italic")).toBe(false);
      expect(reviewer.getChanges().map(({ id }) => id)).toEqual([7]);
    });

    test("is refused when the pending change is another suggestion's", async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await source(), {
        author: "Editor",
        session,
      });
      expect(reviewer.rejectChange(7)).toBe(true);
      const first = suggestFirstWord({
        reviewer,
        id: "bold",
        formatting: { bold: true },
      });
      expect(first.skipped).toEqual([]);
      const second = suggestFirstWord({
        reviewer,
        id: "italic",
        formatting: { italic: true },
      });
      expect(second.applied).toEqual([]);
      expect(second.skipped).toEqual([
        expect.objectContaining({ id: "italic", ...REFUSAL[session] }),
      ]);
      // A deliberate refusal names no missing capability.
      expect(second.skipped.at(0)).not.toHaveProperty("canonicalRefusal");
      // The first suggestion still owns its change, so rejecting it restores
      // the run.
      expect(reviewer.rejectSuggestion("bold")).toBe(true);
      expect(reviewer.getChanges()).toEqual([]);
      expect(await documentXml(await reviewer.toBuffer())).not.toContain(
        "<w:b/>",
      );
    });
  },
);
