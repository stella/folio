import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import { EditorState, type Transaction } from "prosemirror-state";

import { acceptAIEditRevision, rejectAIEditRevision } from "../prosemirror/commands/comments";
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
      if (!firstReceipt?.revisionIds || firstReceipt.revisionId === undefined) {
        panic("expected multiline insertion revision ids");
      }
      const firstRevisionId = firstReceipt.revisionId;
      const expectedFirstIds = Array.from(
        { length: lines.length * 2 + 1 },
        (_, index) => firstRevisionId + index,
      );
      expect(firstReceipt.revisionIds).toEqual(expectedFirstIds);

      const firstChanges = getTrackedChangesFromDoc(view.state.doc);
      const receiptIds = new Set(firstReceipt.revisionIds);
      const syntheticChanges = firstChanges.filter(({ id }) => !receiptIds.has(id));
      expect(syntheticChanges).toEqual([]);
      expect(firstChanges).toContainEqual(
        expect.objectContaining({
          id: first.nextRevisionId - 1,
          type: "paragraphPropertiesChanged",
        }),
      );
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

  test.each(["accept", "reject"] as const)(
    "%s resolves every revision owned by an inserted final paragraph",
    (resolution) => {
      const view = insertionView({ alignment: "center" });
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const anchor = snapshot.blocks.at(0);
      if (!anchor) {
        panic("expected the final insertion anchor");
      }
      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "tail-insert",
            type: "insertAfterBlock",
            blockId: anchor.id,
            text: "A new unformatted final paragraph.",
            inheritFormatting: false,
          },
        ],
        mode: "tracked-changes",
        revisionStamp: { date: "2026-09-08T00:00:00.000Z", idSeed: 10 },
      });
      expect(outcome.applied).toEqual([
        { id: "tail-insert", revisionId: 10, revisionIds: [10, 11, 12] },
      ]);
      expect(outcome.nextRevisionId).toBe(13);
      expect(
        getTrackedChangesFromDoc(view.state.doc)
          .map(({ id }) => id)
          .toSorted((left, right) => left - right),
      ).toEqual([10, 11, 12]);

      const receiptIds = outcome.applied.at(0)?.revisionIds;
      if (!receiptIds) {
        panic("expected the inserted paragraph receipt ids");
      }
      const command = resolution === "accept" ? acceptAIEditRevision : rejectAIEditRevision;
      expect(command(receiptIds)(view.state, view.dispatch)).toBe(true);
      expect(getTrackedChangesFromDoc(view.state.doc)).toEqual([]);
      expect(view.state.doc.childCount).toBe(resolution === "accept" ? 2 : 1);
    },
  );

  test("assigns a rotated final-mark revision to only its responsible insertion", () => {
    const document = createEmptyDocument();
    document.package.document.content = [
      {
        type: "paragraph",
        paraId: "12345678",
        textId: "12345678",
        formatting: { alignment: "center" },
        content: [{ type: "run", content: [{ type: "text", text: "First anchor." }] }],
      },
      {
        type: "paragraph",
        paraId: "23456789",
        textId: "23456789",
        formatting: { alignment: "right" },
        content: [{ type: "run", content: [{ type: "text", text: "Final anchor." }] }],
      },
    ];
    const view = {
      state: EditorState.create({ doc: toProseDoc(document) }),
      dispatch(transaction: Transaction) {
        view.state = view.state.apply(transaction);
      },
    };
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const firstAnchor = snapshot.blocks.at(0);
    const finalAnchor = snapshot.blocks.at(1);
    if (!firstAnchor || !finalAnchor) {
      panic("expected both insertion anchors");
    }
    const outcome = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: [
        {
          id: "middle-insert",
          type: "insertAfterBlock",
          blockId: firstAnchor.id,
          text: "Accepted middle paragraph.",
          inheritFormatting: false,
        },
        {
          id: "tail-insert",
          type: "insertAfterBlock",
          blockId: finalAnchor.id,
          text: "Rejected tail paragraph.",
          inheritFormatting: false,
        },
      ],
      mode: "tracked-changes",
      revisionStamp: { date: "2026-09-08T00:00:00.000Z", idSeed: 20 },
    });
    const middleReceipt = outcome.applied.find(({ id }) => id === "middle-insert");
    const tailReceipt = outcome.applied.find(({ id }) => id === "tail-insert");
    expect(middleReceipt).toEqual({
      id: "middle-insert",
      revisionId: 22,
      revisionIds: [22, 23],
    });
    expect(tailReceipt).toEqual({
      id: "tail-insert",
      revisionId: 20,
      revisionIds: [20, 21, 24],
    });
    expect(outcome.nextRevisionId).toBe(25);
    expect(
      getTrackedChangesFromDoc(view.state.doc)
        .map(({ id }) => id)
        .toSorted((left, right) => left - right),
    ).toEqual([20, 21, 22, 23, 24]);

    if (!middleReceipt?.revisionIds || !tailReceipt?.revisionIds) {
      panic("expected both insertion receipts to own revision ids");
    }
    expect(acceptAIEditRevision(middleReceipt.revisionIds)(view.state, view.dispatch)).toBe(true);
    expect(
      getTrackedChangesFromDoc(view.state.doc)
        .map(({ id }) => id)
        .toSorted((left, right) => left - right),
    ).toEqual([20, 21, 24]);
    expect(rejectAIEditRevision(tailReceipt.revisionIds)(view.state, view.dispatch)).toBe(true);
    expect(getTrackedChangesFromDoc(view.state.doc)).toEqual([]);
    expect(view.state.doc.textContent).toBe("First anchor.Accepted middle paragraph.Final anchor.");
  });
});
