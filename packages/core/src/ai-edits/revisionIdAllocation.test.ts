import { describe, expect, test } from "bun:test";
import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";

import { createDocx } from "../docx/rezip";
import {
  acceptAIEditRevision,
  acceptAllChanges,
  acceptSuggestion,
  getSuggestions,
  rejectAIEditRevision,
  rejectAllChanges,
} from "../prosemirror/commands/comments";
import { expectParagraphAttrs } from "../prosemirror/attrs";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { ParagraphFormatting } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { applyFolioAIEditOperations } from "./apply";
import { FolioDocxReviewer } from "./headless";
import { getTrackedChangesFromDoc } from "./read";
import { createFolioAIEditSnapshot, createFolioAITextRangeHandle } from "./snapshot";

const INSERTION_RESERVATION_CASES = [
  {
    label: "two aligned paragraphs",
    lines: ["Aligned heading", "Aligned body"],
    formatting: { alignment: "center" },
    propertyChangeCount: 1,
  },
  {
    label: "three styled paragraphs",
    lines: ["Styled heading", "Styled body", "Styled conclusion"],
    formatting: { styleId: "Heading2" },
    propertyChangeCount: 2,
  },
  {
    label: "four styled and aligned paragraphs",
    lines: ["First clause", "Second clause", "Third clause", "Fourth clause"],
    formatting: { styleId: "Heading3", alignment: "right" },
    propertyChangeCount: 3,
  },
] as const satisfies readonly {
  label: string;
  lines: readonly string[];
  formatting: ParagraphFormatting;
  propertyChangeCount: number;
}[];

const insertionView = (formatting: ParagraphFormatting, text = "Anchor paragraph.") => {
  const document = createEmptyDocument();
  if (formatting.numPr?.numId !== undefined) {
    document.package.numbering = {
      abstractNums: [
        {
          abstractNumId: 0,
          levels: [
            { ilvl: 0, start: 1, numFmt: "decimal", lvlText: "%1." },
            { ilvl: 1, start: 1, numFmt: "lowerLetter", lvlText: "%2." },
          ],
        },
      ],
      nums: [{ numId: formatting.numPr.numId, abstractNumId: 0 }],
    };
  }
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      textId: "12345678",
      formatting,
      content: text.length > 0 ? [{ type: "run", content: [{ type: "text", text }] }] : [],
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

const formattingReservationView = () => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "12345678",
      content: [
        { bold: true },
        { italic: true },
        { underline: { style: "single" as const } },
        { strike: true },
        { fontSize: 24 },
        { color: { rgb: "C00000" } },
      ].map((formatting, index) => ({
        type: "run" as const,
        formatting: { ...formatting, highlight: "yellow" as const },
        content: [{ type: "text" as const, text: String.fromCharCode(65 + index) }],
      })),
    },
    {
      type: "paragraph",
      paraId: "23456789",
      content: [{ type: "run", content: [{ type: "text", text: "following" }] }],
    },
  ];
  return viewFromDoc(toProseDoc(document));
};

const viewFromDoc = (doc: PMNode) => {
  const view = {
    state: EditorState.create({ doc }),
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  return view;
};

const reopenedView = async (doc: PMNode) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(fromProseDoc(doc)));
  return viewFromDoc(toProseDoc(reviewer.toDocument()));
};

const SAME_ANCHOR_INSERTIONS = [
  { id: "first", text: "First inserted.", alignment: "center" },
  { id: "second", text: "Second inserted.", alignment: "right" },
  { id: "third", text: "Third inserted.", alignment: "both" },
] as const;

type SameAnchorOperationId = (typeof SAME_ANCHOR_INSERTIONS)[number]["id"];

const SAME_ANCHOR_RESOLUTION_ORDERS = [
  ["first", "second", "third"],
  ["first", "third", "second"],
  ["second", "first", "third"],
  ["second", "third", "first"],
  ["third", "first", "second"],
  ["third", "second", "first"],
] as const satisfies readonly (readonly SameAnchorOperationId[])[];

const SAME_ANCHOR_RESOLUTION_DECISIONS = [
  ["reject", "reject", "reject"],
  ["accept", "reject", "reject"],
  ["reject", "accept", "reject"],
  ["reject", "reject", "accept"],
  ["accept", "accept", "reject"],
  ["accept", "reject", "accept"],
  ["reject", "accept", "accept"],
  ["accept", "accept", "accept"],
] as const satisfies readonly (readonly ("accept" | "reject")[])[];

const EMPTY_CARRIER_FORMATTING = {
  styleId: "Heading2",
  numPr: { numId: 1, ilvl: 1 },
  alignment: "both",
} as const satisfies ParagraphFormatting;

type SameAnchorTrackedInsertionsOptions = {
  anchorFormatting?: ParagraphFormatting;
  anchorText?: string;
};

const sameAnchorTrackedInsertions = ({
  anchorFormatting = { alignment: "left" },
  anchorText,
}: SameAnchorTrackedInsertionsOptions = {}) => {
  const view = insertionView(anchorFormatting, anchorText);
  const snapshot = createFolioAIEditSnapshot(view.state.doc);
  const anchor = snapshot.blocks.at(0);
  if (!anchor) {
    return panic("expected the same-anchor insertion target");
  }
  const outcome = applyFolioAIEditOperations({
    view,
    snapshot,
    operations: SAME_ANCHOR_INSERTIONS.map(({ id, text, alignment }) => ({
      id,
      type: "insertAfterBlock" as const,
      blockId: anchor.id,
      text,
      inheritFormatting: false,
      alignment,
    })),
    mode: "tracked-changes",
    revisionStamp: { date: "2026-09-08T00:00:00.000Z", idSeed: 30 },
  });
  return { outcome, view };
};

const operationDecision = (
  operationId: SameAnchorOperationId,
  decisions: (typeof SAME_ANCHOR_RESOLUTION_DECISIONS)[number],
) => {
  switch (operationId) {
    case "first":
      return decisions[0];
    case "second":
      return decisions[1];
    case "third":
      return decisions[2];
  }
};

const operationRevisionIds = (
  outcome: ReturnType<typeof applyFolioAIEditOperations>,
  operationId: string,
): number[] => {
  const ids = outcome.applied.find(({ id }) => id === operationId)?.revisionIds;
  return ids ? [...ids] : panic(`expected revision ids for ${operationId}`);
};

const paragraphState = (doc: PMNode) =>
  Array.from({ length: doc.childCount }, (_, index) => {
    const node = doc.child(index);
    const attrs = expectParagraphAttrs(node);
    const mark: unknown = node.attrs["pPrMark"];
    const info = typeof mark === "object" && mark !== null && "info" in mark ? mark.info : null;
    const markId = typeof info === "object" && info !== null && "id" in info ? info.id : null;
    const propertyChanges = attrs._propertyChanges;
    return {
      text: node.textContent,
      alignment: attrs.alignment,
      markId,
      changes: Array.isArray(propertyChanges)
        ? propertyChanges.map(({ info: changeInfo, previousFormatting }) => ({
            revisionId: changeInfo.id,
            previousFormatting,
          }))
        : null,
    };
  });

describe("unstamped revision id allocation", () => {
  test.each([
    { label: "without a background revision", highlight: undefined, expectedIds: [10, 11, 12] },
    { label: "with a background revision", highlight: "yellow", expectedIds: [10, 11, 12, 13] },
  ] as const)(
    "a block replacement owns exactly its serialized revisions $label",
    ({ highlight, expectedIds }) => {
      const document = createEmptyDocument();
      document.package.document.content = [
        {
          type: "paragraph",
          paraId: "12345678",
          formatting: { styleId: "BodyText" },
          content: [
            {
              type: "run",
              ...(highlight !== undefined && { formatting: { highlight } }),
              content: [{ type: "text", text: "Original text." }],
            },
          ],
        },
      ];
      const view = viewFromDoc(toProseDoc(document));
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const block = snapshot.blocks.at(0);
      if (!block) {
        panic("expected a replacement block");
      }

      const outcome = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [
          {
            id: "replacement",
            type: "replaceBlock",
            blockId: block.id,
            text: "Replacement text.",
            styleId: "Heading2",
          },
        ],
        mode: "tracked-changes",
        revisionStamp: { date: "2026-09-08T00:00:00.000Z", idSeed: 10 },
      });
      const receiptIds = outcome.applied.at(0)?.revisionIds;
      const serializedChanges = getTrackedChangesFromDoc(view.state.doc);
      const serializedIds = serializedChanges
        .map(({ id }) => id)
        .toSorted((left, right) => left - right);

      expect(receiptIds).toEqual(expectedIds);
      expect(serializedIds).toEqual(expectedIds);
      expect(outcome.nextRevisionId).toBe((expectedIds.at(-1) ?? 9) + 1);
      expect(serializedChanges.filter(({ type }) => type === "formatting")).toHaveLength(
        highlight === undefined ? 0 : 1,
      );
    },
  );

  test.each(["formatRange", "replaceRange"] as const)(
    "%s advances the shared cursor past every physical formatting carrier",
    (operationType) => {
      const view = formattingReservationView();
      const snapshot = createFolioAIEditSnapshot(view.state.doc);
      const firstBlock = snapshot.blocks.at(0);
      if (!firstBlock) {
        panic("expected the segmented formatting block");
      }
      const firstRange = createFolioAITextRangeHandle({
        blockId: firstBlock.id,
        text: firstBlock.text,
        startOffset: 0,
        endOffset: firstBlock.text.length,
      });
      if (!firstRange) {
        panic("expected the segmented formatting range");
      }
      const operation =
        operationType === "formatRange"
          ? {
              id: "many-carriers",
              type: operationType,
              range: firstRange,
              formatting: { color: "00AA00" },
            }
          : {
              id: "many-carriers",
              type: operationType,
              range: firstRange,
              replace: "replacement",
            };
      const first = applyFolioAIEditOperations({
        view,
        snapshot,
        operations: [operation],
        mode: "tracked-changes",
      });
      const firstIds = operationRevisionIds(first, operation.id);
      expect(firstIds.length).toBeGreaterThan(4);

      const followingSnapshot = createFolioAIEditSnapshot(view.state.doc);
      const followingBlock = followingSnapshot.blocks.at(1);
      if (!followingBlock) {
        panic("expected the following formatting block");
      }
      const followingRange = createFolioAITextRangeHandle({
        blockId: followingBlock.id,
        text: followingBlock.text,
        startOffset: 0,
        endOffset: followingBlock.text.length,
      });
      if (!followingRange) {
        panic("expected the following formatting range");
      }
      const following = applyFolioAIEditOperations({
        view,
        snapshot: followingSnapshot,
        operations: [
          {
            id: "following-batch",
            type: "formatRange",
            range: followingRange,
            formatting: { bold: true },
          },
        ],
        mode: "tracked-changes",
      });
      const followingIds = operationRevisionIds(following, "following-batch");
      expect(Math.min(...followingIds)).toBeGreaterThanOrEqual(first.nextRevisionId);
      expect(followingIds.some((id) => firstIds.includes(id))).toBe(false);

      const allIds = getTrackedChangesFromDoc(view.state.doc).map(({ id }) => id);
      expect(new Set(allIds).size).toBe(allIds.length);
    },
  );

  test("claims every physical formatting carrier before dispatch can re-enter", () => {
    const primaryView = formattingReservationView();
    const primarySnapshot = createFolioAIEditSnapshot(primaryView.state.doc);
    const primaryBlock = primarySnapshot.blocks.at(0);
    if (!primaryBlock) {
      panic("expected the primary segmented formatting block");
    }
    const primaryRange = createFolioAITextRangeHandle({
      blockId: primaryBlock.id,
      text: primaryBlock.text,
      startOffset: 0,
      endOffset: primaryBlock.text.length,
    });
    if (!primaryRange) {
      panic("expected the primary segmented formatting range");
    }

    const reentrantView = formattingReservationView();
    const reentrantSnapshot = createFolioAIEditSnapshot(reentrantView.state.doc);
    const reentrantBlock = reentrantSnapshot.blocks.at(1);
    if (!reentrantBlock) {
      panic("expected the re-entrant formatting block");
    }
    const reentrantRange = createFolioAITextRangeHandle({
      blockId: reentrantBlock.id,
      text: reentrantBlock.text,
      startOffset: 0,
      endOffset: reentrantBlock.text.length,
    });
    if (!reentrantRange) {
      panic("expected the re-entrant formatting range");
    }

    let reentrantOutcome: ReturnType<typeof applyFolioAIEditOperations> | undefined;
    const commitPrimary = primaryView.dispatch;
    primaryView.dispatch = (transaction) => {
      commitPrimary(transaction);
      reentrantOutcome = applyFolioAIEditOperations({
        view: reentrantView,
        snapshot: reentrantSnapshot,
        operations: [
          {
            id: "re-entrant-batch",
            type: "formatRange",
            range: reentrantRange,
            formatting: { bold: true },
          },
        ],
        mode: "tracked-changes",
      });
    };

    const primaryOutcome = applyFolioAIEditOperations({
      view: primaryView,
      snapshot: primarySnapshot,
      operations: [
        {
          id: "primary-many-carriers",
          type: "formatRange",
          range: primaryRange,
          formatting: { color: "00AA00" },
        },
      ],
      mode: "tracked-changes",
    });
    if (!reentrantOutcome) {
      panic("expected dispatch to apply the re-entrant batch");
    }

    const primaryIds = operationRevisionIds(primaryOutcome, "primary-many-carriers");
    const reentrantIds = operationRevisionIds(reentrantOutcome, "re-entrant-batch");
    expect(primaryIds.length).toBeGreaterThan(4);
    expect(Math.min(...reentrantIds)).toBeGreaterThanOrEqual(primaryOutcome.nextRevisionId);
    expect(reentrantIds.some((id) => primaryIds.includes(id))).toBe(false);
  });

  test.each(INSERTION_RESERVATION_CASES)(
    "reserves the synthetic final-mark revision for $label",
    ({ lines, formatting, propertyChangeCount }) => {
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
        { length: lines.length * 2 + propertyChangeCount },
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
      expect(firstChanges.filter(({ type }) => type === "paragraphPropertiesChanged")).toHaveLength(
        propertyChangeCount,
      );
      const firstIds = new Set(firstChanges.map(({ id }) => id));
      expect(firstIds.size).toBe(lines.length * 2 + propertyChangeCount);

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

  test("rotates a same-anchor run one boundary left without changing operation ownership", async () => {
    const { outcome, view } = sameAnchorTrackedInsertions();
    const firstIds = operationRevisionIds(outcome, "first");
    const secondIds = operationRevisionIds(outcome, "second");
    const thirdIds = operationRevisionIds(outcome, "third");

    expect([firstIds.length, secondIds.length, thirdIds.length]).toEqual([3, 3, 3]);
    expect(firstIds.at(-1)).not.toBe(firstIds.at(1));
    expect(secondIds.at(-1)).not.toBe(secondIds.at(1));
    expect(thirdIds.at(-1)).not.toBe(thirdIds.at(1));

    expect(paragraphState(view.state.doc)).toEqual([
      {
        text: "Anchor paragraph.",
        alignment: "left",
        markId: firstIds.at(1),
        changes: null,
      },
      {
        text: "First inserted.",
        alignment: "center",
        markId: secondIds.at(1),
        changes: [
          expect.objectContaining({
            revisionId: firstIds.at(-1),
            previousFormatting: expect.objectContaining({ alignment: "left" }),
          }),
        ],
      },
      {
        text: "Second inserted.",
        alignment: "right",
        markId: thirdIds.at(1),
        changes: [
          expect.objectContaining({
            revisionId: secondIds.at(-1),
            previousFormatting: expect.objectContaining({ alignment: "left" }),
          }),
        ],
      },
      {
        text: "Third inserted.",
        alignment: "both",
        markId: null,
        changes: [
          expect.objectContaining({
            revisionId: thirdIds.at(-1),
            previousFormatting: expect.objectContaining({ alignment: "left" }),
          }),
        ],
      },
    ]);
    expect(outcome.nextRevisionId).toBe(39);
    expect(new Set(outcome.applied.flatMap(({ revisionIds }) => revisionIds ?? [])).size).toBe(9);

    const reopened = await reopenedView(view.state.doc);
    expect(paragraphState(reopened.state.doc)).toEqual(paragraphState(view.state.doc));
    expect(
      getTrackedChangesFromDoc(reopened.state.doc)
        .map(({ id }) => id)
        .toSorted((left, right) => left - right),
    ).toEqual(
      outcome.applied
        .flatMap(({ revisionIds }) => revisionIds ?? [])
        .toSorted((left, right) => left - right),
    );
  });

  test.each(SAME_ANCHOR_RESOLUTION_ORDERS.map((order) => ({ label: order.join(" then "), order })))(
    "targeted same-anchor resolutions are independent in $label order",
    async ({ order }) => {
      const { outcome, view } = sameAnchorTrackedInsertions();
      for (const operationId of order) {
        const command = operationId === "second" ? rejectAIEditRevision : acceptAIEditRevision;
        expect(command(operationRevisionIds(outcome, operationId))(view.state, view.dispatch)).toBe(
          true,
        );
      }

      expect(getTrackedChangesFromDoc(view.state.doc)).toEqual([]);
      expect(
        paragraphState(view.state.doc).map(({ text, alignment }) => ({ text, alignment })),
      ).toEqual([
        { text: "Anchor paragraph.", alignment: "left" },
        { text: "First inserted.", alignment: "center" },
        { text: "Third inserted.", alignment: "both" },
      ]);
      const reopened = await reopenedView(view.state.doc);
      expect(getTrackedChangesFromDoc(reopened.state.doc)).toEqual([]);
      expect(
        paragraphState(reopened.state.doc).map(({ text, alignment, markId, changes }) => ({
          text,
          alignment,
          markId,
          changes,
        })),
      ).toEqual(
        paragraphState(view.state.doc).map(({ text, alignment, markId, changes }) => ({
          text,
          alignment,
          markId,
          changes,
        })),
      );
    },
  );

  test("an empty formatted carrier preserves every boundary owner and resolves in every order", () => {
    for (const decisions of SAME_ANCHOR_RESOLUTION_DECISIONS) {
      let canonical: unknown;
      for (const order of SAME_ANCHOR_RESOLUTION_ORDERS) {
        const { outcome, view } = sameAnchorTrackedInsertions({
          anchorFormatting: EMPTY_CARRIER_FORMATTING,
          anchorText: "",
        });
        const initial = paragraphState(view.state.doc);
        for (const [index, insertion] of SAME_ANCHOR_INSERTIONS.entries()) {
          const revisionIds = operationRevisionIds(outcome, insertion.id);
          expect(initial[index]?.markId).toBe(revisionIds.at(1));
          expect(initial[index + 1]?.changes).toEqual([
            expect.objectContaining({
              revisionId: revisionIds.at(-1),
              previousFormatting: expect.objectContaining(EMPTY_CARRIER_FORMATTING),
            }),
          ]);
        }

        for (const operationId of order) {
          const command =
            operationDecision(operationId, decisions) === "accept"
              ? acceptAIEditRevision
              : rejectAIEditRevision;
          expect(
            command(operationRevisionIds(outcome, operationId))(view.state, view.dispatch),
          ).toBe(true);
        }

        expect(getTrackedChangesFromDoc(view.state.doc)).toEqual([]);
        const expectedInsertions = SAME_ANCHOR_INSERTIONS.filter(
          (_insertion, index) => decisions[index] === "accept",
        );
        expect(
          paragraphState(view.state.doc).map(({ text, alignment }) => ({ text, alignment })),
        ).toEqual([
          { text: "", alignment: EMPTY_CARRIER_FORMATTING.alignment },
          ...expectedInsertions.map(({ text, alignment }) => ({ text, alignment })),
        ]);
        expect(expectParagraphAttrs(view.state.doc.child(0))).toMatchObject({
          ...EMPTY_CARRIER_FORMATTING,
          _originalFormatting: EMPTY_CARRIER_FORMATTING,
        });

        if (canonical === undefined) {
          canonical = view.state.doc.toJSON();
        } else {
          expect(view.state.doc.toJSON()).toEqual(canonical);
        }
      }
    }
  });

  test.each(["accept", "reject"] as const)(
    "%s all resolves and reopens a same-anchor terminal chain",
    async (resolution) => {
      const { view } = sameAnchorTrackedInsertions();
      const command = resolution === "accept" ? acceptAllChanges : rejectAllChanges;
      expect(command()(view.state, view.dispatch)).toBe(true);
      expect(getTrackedChangesFromDoc(view.state.doc)).toEqual([]);
      expect(view.state.doc.textContent).toBe(
        resolution === "accept"
          ? "Anchor paragraph.First inserted.Second inserted.Third inserted."
          : "Anchor paragraph.",
      );

      const reopened = await reopenedView(view.state.doc);
      expect(getTrackedChangesFromDoc(reopened.state.doc)).toEqual([]);
      expect(reopened.state.doc.textContent).toBe(view.state.doc.textContent);
    },
  );

  test.each(["accept", "reject"] as const)(
    "%s all preserves an empty formatted terminal carrier",
    (resolution) => {
      const { view } = sameAnchorTrackedInsertions({
        anchorFormatting: EMPTY_CARRIER_FORMATTING,
        anchorText: "",
      });
      const command = resolution === "accept" ? acceptAllChanges : rejectAllChanges;
      expect(command()(view.state, view.dispatch)).toBe(true);
      expect(getTrackedChangesFromDoc(view.state.doc)).toEqual([]);
      expect(view.state.doc.textContent).toBe(
        resolution === "accept" ? "First inserted.Second inserted.Third inserted." : "",
      );
      expect(expectParagraphAttrs(view.state.doc.child(0))).toMatchObject({
        ...EMPTY_CARRIER_FORMATTING,
        _originalFormatting: EMPTY_CARRIER_FORMATTING,
      });
    },
  );

  test("resolves same-anchor suggested insertions out of order after save and reopen", async () => {
    const view = insertionView({ alignment: "left" });
    const snapshot = createFolioAIEditSnapshot(view.state.doc);
    const anchor = snapshot.blocks.at(0);
    if (!anchor) {
      return panic("expected the same-anchor suggestion target");
    }
    const outcome = applyFolioAIEditOperations({
      view,
      snapshot,
      operations: SAME_ANCHOR_INSERTIONS.map(({ id, text, alignment }) => ({
        id,
        type: "insertAfterBlock" as const,
        blockId: anchor.id,
        text,
        inheritFormatting: false,
        alignment,
      })),
      mode: "suggested",
      author: "Assistant",
      revisionStamp: { date: "2026-09-08T00:00:00.000Z", idSeed: 50 },
    });
    expect(outcome.skipped).toEqual([]);
    expect(
      outcome.applied.map(({ id, suggestionId, revisionIds }) => ({
        id,
        suggestionId,
        revisionIds,
      })),
    ).toEqual([
      { id: "third", suggestionId: "third", revisionIds: [50] },
      { id: "second", suggestionId: "second", revisionIds: [51] },
      { id: "first", suggestionId: "first", revisionIds: [52] },
    ]);

    for (const suggestionId of ["first", "third", "second"]) {
      expect(
        acceptSuggestion(suggestionId, {
          author: "Reviewer",
          date: "2026-09-08T00:00:00.000Z",
        })(view.state, view.dispatch),
      ).toBe(true);
    }
    expect(getSuggestions(view.state)).toEqual([]);

    const reopened = await reopenedView(view.state.doc);
    const accepting = viewFromDoc(reopened.state.doc);
    expect(acceptAllChanges()(accepting.state, accepting.dispatch)).toBe(true);
    expect(getTrackedChangesFromDoc(accepting.state.doc)).toEqual([]);
    expect(accepting.state.doc.textContent).toBe(
      "Anchor paragraph.First inserted.Second inserted.Third inserted.",
    );

    const rejecting = viewFromDoc(reopened.state.doc);
    expect(rejectAllChanges()(rejecting.state, rejecting.dispatch)).toBe(true);
    expect(getTrackedChangesFromDoc(rejecting.state.doc)).toEqual([]);
    expect(rejecting.state.doc.textContent).toBe("Anchor paragraph.");
  });
});
