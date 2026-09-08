import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState, type Transaction } from "prosemirror-state";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { ParagraphFormatting } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { applyFolioAIEditOperations } from "./apply";
import { getTrackedChangesFromDoc } from "./read";
import { createFolioAIEditSnapshot } from "./snapshot";

const INSERTION_RESERVATION_CASES = [
  {
    label: "two aligned paragraphs",
    lines: ["Aligned heading", "Aligned body"],
    formatting: { alignment: "center" },
  },
  {
    label: "three styled paragraphs",
    lines: ["Styled heading", "Styled body", "Styled conclusion"],
    formatting: { styleId: "Heading2" },
  },
  {
    label: "four styled and aligned paragraphs",
    lines: ["First clause", "Second clause", "Third clause", "Fourth clause"],
    formatting: { styleId: "Heading3", alignment: "right" },
  },
] as const satisfies readonly {
  label: string;
  lines: readonly string[];
  formatting: ParagraphFormatting;
}[];

const insertionView = (formatting: ParagraphFormatting) => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      textId: "12345678",
      formatting,
      content: [{ type: "run", content: [{ type: "text", text: "Anchor paragraph." }] }],
    },
  ];
  const view = {
    state: EditorState.create({ doc: toProseDoc(document) }),
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  return view;
};

describe("unstamped revision id allocation", () => {
  test.each(INSERTION_RESERVATION_CASES)(
    "reserves the synthetic final-mark revision for $label",
    ({ lines, formatting }) => {
      const view = insertionView(formatting);
      const firstSnapshot = createFolioAIEditSnapshot(view.state.doc);
      const firstAnchor = firstSnapshot.blocks.at(0);
      if (!firstAnchor) {
        panic("expected the first insertion anchor");
      }
      const first = applyFolioAIEditOperations({
        view,
        snapshot: firstSnapshot,
        operations: [
          {
            id: "multiline-insert",
            type: "insertAfterBlock",
            blockId: firstAnchor.id,
            text: lines.join("\n"),
            inheritFormatting: true,
          },
        ],
        mode: "tracked-changes",
      });
      const firstReceipt = first.applied.at(0);
      if (!firstReceipt?.revisionIds) {
        panic("expected multiline insertion revision ids");
      }
      expect(firstReceipt.revisionIds).toHaveLength(lines.length * 2);

      const firstChanges = getTrackedChangesFromDoc(view.state.doc);
      const receiptIds = new Set(firstReceipt.revisionIds);
      const syntheticChanges = firstChanges.filter(({ id }) => !receiptIds.has(id));
      expect(syntheticChanges).toEqual([
        expect.objectContaining({
          id: first.nextRevisionId - 1,
          type: "paragraphPropertiesChanged",
        }),
      ]);
      const firstIds = new Set(firstChanges.map(({ id }) => id));
      expect(firstIds.size).toBe(lines.length * 2 + 1);

      const secondSnapshot = createFolioAIEditSnapshot(view.state.doc);
      const secondAnchor = secondSnapshot.blocks.at(0);
      if (!secondAnchor) {
        panic("expected the second insertion anchor");
      }
      const second = applyFolioAIEditOperations({
        view,
        snapshot: secondSnapshot,
        operations: [
          {
            id: "following-insert",
            type: "insertBeforeBlock",
            blockId: secondAnchor.id,
            text: "A later batch must use fresh ids.",
          },
        ],
        mode: "tracked-changes",
      });
      const secondRevisionIds = second.applied.at(0)?.revisionIds;
      if (!secondRevisionIds) {
        panic("expected following insertion revision ids");
      }
      expect(secondRevisionIds.every((id) => !firstIds.has(id))).toBe(true);

      const allIds = getTrackedChangesFromDoc(view.state.doc).map(({ id }) => id);
      expect(new Set(allIds).size).toBe(allIds.length);
    },
  );
});
