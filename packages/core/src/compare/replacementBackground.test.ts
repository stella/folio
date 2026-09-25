/**
 * A comparison never writes a run-property change that changes nothing.
 *
 * Replacing text clears a highlight or `w:shd` under it, because text typed
 * over a highlighted placeholder is new text and the marker that said "fill
 * this in" should not survive into the finished document. A comparison is not
 * authoring: it holds the revised document's own run properties and writes
 * them itself, so clearing them first only records a `w:rPrChange` that the
 * provenance pass takes straight back, leaving the reader a revision whose
 * before and after are identical.
 */

import { expect, test } from "bun:test";
import JSZip from "jszip";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Transaction } from "prosemirror-state";

import { applyFolioAIEditOperations } from "../ai-edits/apply";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { createDocx } from "../docx/rezip";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, TextFormatting } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { compareDocx } from "./compare";

const OPTIONS = {
  author: "compare",
  timestamp: "2024-03-01T00:00:00.000Z",
  onUnverified: "emit",
} as const;

/** The drafting-note look these cells carry: italic on a grey background. */
const DRAFTING_NOTE: TextFormatting = {
  italic: true,
  highlight: "lightGray",
  color: { rgb: "000000" },
  fontSize: 16,
};

const paragraph = (paraId: string, text: string, formatting: TextFormatting) =>
  ({
    type: "paragraph",
    paraId,
    textId: paraId,
    content: [{ type: "run", formatting, content: [{ type: "text", text }] }],
  }) as const satisfies BlockContent;

const PAYMENT_TERMS = (days: string) =>
  `Pay by invoice: Customer will pay each invoice within ${days} days of the invoice date.`;

const orderForm = (days: string): BlockContent => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [
        {
          type: "tableCell",
          content: [paragraph("33333333", "Payment Process", { fontSize: 16 })],
        },
        {
          type: "tableCell",
          content: [paragraph("22222222", PAYMENT_TERMS(days), DRAFTING_NOTE)],
        },
      ],
    },
  ],
});

const buildDocx = (days: string) => {
  const result = createEmptyDocument();
  result.package.document.content = [
    paragraph("11111111", "This agreement is made between the parties.", {}),
    orderForm(days),
  ];
  return createDocx(result);
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const entry = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!entry) {
    throw new Error("The compared package has no main document part.");
  }
  return await entry.async("string");
};

test("a replacement in a highlighted cell paragraph records no run-property change", async () => {
  const result = await compareDocx(await buildDocx("30"), await buildDocx("45"), OPTIONS);
  if (result.isErr()) {
    throw result.error;
  }

  const xml = await documentXml(result.value.buffer);

  expect(result.value.verification).toEqual({ status: "verified" });
  expect(xml).not.toContain("<w:rPrChange ");
  // The highlight the revised document states, on the text that survives and
  // on the text the replacement adds.
  expect(xml).toContain('<w:highlight w:val="lightGray"/>');
  expect(xml).toContain("<w:delText");
});

/**
 * The other half of the option, so neither side can drift: an authoring
 * replacement still clears the background it types over.
 */
test("an authoring replacement still clears the background it types over", () => {
  const document = createEmptyDocument();
  document.package.document.content = [paragraph("22222222", PAYMENT_TERMS("30"), DRAFTING_NOTE)];
  const doc: PMNode = toProseDoc(document);
  const view = {
    state: EditorState.create({ doc }),
    dispatch(transaction: Transaction) {
      view.state = view.state.apply(transaction);
    },
  };
  const snapshot = createFolioAIEditSnapshot(view.state.doc);
  const block = snapshot.blocks.at(0);
  if (!block) {
    throw new Error("expected a replacement block");
  }

  applyFolioAIEditOperations({
    view,
    snapshot,
    operations: [
      { id: "replacement", type: "replaceBlock", blockId: block.id, text: PAYMENT_TERMS("45") },
    ],
    mode: "tracked-changes",
    revisionStamp: { date: "2026-09-08T00:00:00.000Z", idSeed: 10 },
  });

  // Every character the paragraph keeps or gains loses the background; the
  // removed characters keep theirs, as deleted text.
  const highlighted: { text: string; deleted: boolean }[] = [];
  view.state.doc.descendants((node) => {
    if (node.isText && node.marks.some(({ type }) => type.name === "highlight")) {
      highlighted.push({
        text: node.text ?? "",
        deleted: node.marks.some(({ type }) => type.name === "deletion"),
      });
    }
    return true;
  });
  expect(highlighted.every(({ deleted }) => deleted)).toBe(true);
});
