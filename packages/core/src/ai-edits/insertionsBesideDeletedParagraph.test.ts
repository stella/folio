/**
 * A direct batch that deletes a paragraph and puts blocks before and after
 * it: the blocks put before it stand before those put after it, each side in
 * request order, whatever follows the deleted paragraph.
 */

import { expect, test } from "bun:test";

import type { FolioDocumentOperation } from "../document-operations";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import { FolioDocxReviewer } from "./headless";

const SESSIONS = [
  "legacy",
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const FOLLOWERS = ["a table", "a paragraph", "nothing"] as const;

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer
    .snapshot()
    .blocks.map(({ text, table }) =>
      table === undefined ? text : `[${text}]`,
    );

for (const session of SESSIONS)
  for (const follower of FOLLOWERS)
    test(`blocks beside a deleted paragraph followed by ${follower} keep their sides (${session})`, async () => {
      const reviewer = await FolioDocxReviewer.fromBuffer(
        (
          await ensureParaIds(
            await createDocx(fromMarkdown("Opening.\n\nAnchor.")),
          )
        ).docx,
        { author: "Test", session },
      );
      const anchor = reviewer
        .snapshot()
        .blocks.find(({ text }) => text === "Anchor.")?.id;
      if (anchor === undefined) throw new Error("Missing anchor");
      const followers = {
        "a table": [
          { id: "t", type: "insertTable", blockId: anchor, rows: [["Cell"]] },
        ],
        "a paragraph": [
          { id: "t", type: "insertAfterBlock", blockId: anchor, text: "Tail." },
        ],
        nothing: [],
      } as const satisfies Record<
        (typeof FOLLOWERS)[number],
        readonly FolioDocumentOperation[]
      >;
      expect(
        reviewer.applyOperations([...followers[follower]], { mode: "direct" })
          .skipped,
      ).toEqual([]);
      const after = (id: string): FolioDocumentOperation => ({
        id,
        type: "insertAfterBlock",
        blockId: anchor,
        text: id,
      });
      const before = (id: string): FolioDocumentOperation => ({
        id,
        type: "insertBeforeBlock",
        blockId: anchor,
        text: id,
      });
      expect(
        reviewer.applyOperations(
          [
            after("After one."),
            before("Before one."),
            after("After two."),
            before("Before two."),
            { id: "delete", type: "deleteBlock", blockId: anchor },
          ],
          { mode: "direct" },
        ).skipped,
      ).toEqual([]);
      const rest = {
        "a table":
          // The canonical session closes a body that ends in a table with an
          // empty paragraph.
          session === "canonical" ? ["[Cell]", ""] : ["[Cell]"],
        "a paragraph": ["Tail."],
        nothing: [],
      } as const satisfies Record<
        (typeof FOLLOWERS)[number],
        readonly string[]
      >;
      expect(texts(reviewer)).toEqual([
        "Opening.",
        "Before one.",
        "Before two.",
        "After one.",
        "After two.",
        ...rest[follower],
      ]);
    });
