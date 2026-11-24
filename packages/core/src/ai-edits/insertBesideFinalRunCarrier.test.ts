/**
 * A batch that deletes the paragraphs a container ends with leaves the
 * paragraph before them holding the container's final mark, under the final
 * paragraph's identity. A paragraph the same batch inserts beside that one
 * still lands beside its words.
 */

import { describe, expect, test } from "bun:test";

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { FolioDocxReviewer } from "./headless";

const TEXTS = ["p0", "p1", "p2", "p3", "p4"] as const;
const SESSIONS = ["canonical"] as const;

const CASES = [
  {
    side: "insertBeforeBlock",
    mode: "direct",
    expected: ["p0", "p1", "x", "p2"],
  },
  {
    side: "insertBeforeBlock",
    mode: "tracked-changes",
    expected: ["p0", "p1", "x", "p2"],
  },
  {
    side: "insertAfterBlock",
    mode: "direct",
    expected: ["p0", "p1", "p2", "x"],
  },
] as const;

const source = (async () => {
  const bytes = (
    await ensureParaIds(await createDocx(fromMarkdown(TEXTS.join("\n\n"))))
  ).docx;
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
})();

describe.each(SESSIONS)("%s session", (session) => {
  test.each(CASES)(
    "$side p2 while the batch deletes p3 and p4 ($mode)",
    async ({ side, mode, expected }) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await source, {
        author: "Test",
        session,
      });
      const [, , anchor, third, last] = reviewer
        .getContent()
        .map(({ id }) => id);
      if (anchor === undefined || third === undefined || last === undefined)
        throw new Error("The fixture has fewer than five paragraphs.");
      const operations: FolioDocumentOperation[] = [
        { id: "d4", type: "deleteBlock", blockId: last },
        { id: "d3", type: "deleteBlock", blockId: third },
        { id: "x", type: side, blockId: anchor, text: "x" },
      ];
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations,
        },
        {},
      );
      expect(result.skipped).toEqual([]);
      reviewer.acceptAll();
      expect(
        reviewer
          .getContent()
          .map((block) => ("text" in block ? block.text : block.kind)),
      ).toEqual([...expected]);
    },
  );
});
