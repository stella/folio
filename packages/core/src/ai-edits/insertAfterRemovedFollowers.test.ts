/**
 * Several insertions after one paragraph, in a batch that also deletes the
 * paragraphs after it, land where that paragraph stood, in the order they are
 * stated, whatever the batch removes around them.
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

const TEXTS = ["p0", "p1", "p2", "p3", "p4", "p5", "p6"] as const;
const SESSIONS = ["canonical"] as const;
const MODES = ["direct", "tracked-changes"] as const;

const CASES = [
  { name: "the anchor and its follower", deleted: [2, 3] },
  { name: "the anchor and two followers", deleted: [2, 3, 4] },
  { name: "the anchor and every follower", deleted: [2, 3, 4, 5, 6] },
  { name: "followers of a kept anchor", deleted: [3, 4] },
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
  test.each(
    CASES.flatMap(({ name, deleted }) =>
      MODES.map((mode) => ({ name, deleted, mode })),
    ),
  )(
    "two insertions after p2 when the batch deletes $name ($mode)",
    async ({ deleted, mode }) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await source, {
        author: "Test",
        session,
      });
      const ids = reviewer.getContent().map(({ id }) => id);
      const anchor = ids.at(2);
      if (anchor === undefined) throw new Error("The fixture has no p2.");
      const operations: FolioDocumentOperation[] = [
        { id: "a", type: "insertAfterBlock", blockId: anchor, text: "a" },
        { id: "b", type: "insertAfterBlock", blockId: anchor, text: "b" },
      ];
      for (const index of deleted) {
        const blockId = ids.at(index);
        if (blockId === undefined) throw new Error(`No p${index}.`);
        operations.push({ id: `d${index}`, type: "deleteBlock", blockId });
      }
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
      const removed = new Set<string>(deleted.map((index) => `p${index}`));
      const expected = [...TEXTS.slice(0, 3), "a", "b", ...TEXTS.slice(3)];
      expect(
        reviewer
          .getContent()
          .map((block) => ("text" in block ? block.text : block.kind)),
      ).toEqual(expected.filter((text) => !removed.has(text)));
    },
  );
});
