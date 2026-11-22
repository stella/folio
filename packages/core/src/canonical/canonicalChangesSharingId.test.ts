/**
 * A tracked deletion an inline content control interrupts is listed on both
 * sides of the control under one revision id, and the words after a bookmark
 * end carry an id of their own. Resolving the changes resolves every record.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { Paragraph, Run } from "../types/document";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const content = (
  control: Paragraph["content"],
  between: Paragraph["content"],
): Paragraph["content"] => [
  ...control,
  ...between,
  { type: "bookmarkEnd", id: 1 },
  run("last"),
];

const CONTROL: Paragraph["content"] = [
  {
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 1, tag: "clause" },
    content: [run("control")],
  },
];

const LINKED_TAB: Paragraph["content"] = [
  {
    type: "hyperlink",
    rId: "rId9",
    href: "https://example.org",
    children: [{ type: "run", content: [{ type: "tab" }] }],
  },
];

const SHAPES: Record<string, Paragraph["content"]> = {
  "a control before a linked tab": content(CONTROL, LINKED_TAB),
  "a control before its words": content(CONTROL, [run("between")]),
  "a linked tab alone": content([], LINKED_TAB),
};

const opened = async (words: Paragraph["content"]) => {
  const bytes = await createDocx({
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [] },
          { type: "paragraph", paraId: "00000002", content: words },
        ],
      },
    },
  });
  return FolioDocxReviewer.fromBuffer(bytes, {
    author: "Editor",
    session: "canonical",
  });
};

const texts = (reviewer: FolioDocxReviewer): string[] =>
  reviewer.getContent().map(({ text }) => text);

const deleted = async (
  words: Paragraph["content"],
): Promise<FolioDocxReviewer> => {
  const reviewer = await opened(words);
  const { skipped } = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "op", type: "deleteBlock", blockId: "00000002" }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(skipped).toEqual([]);
  return reviewer;
};

for (const [shape, words] of Object.entries(SHAPES)) {
  test(`rejecting every change gives the words back after ${shape}`, async () => {
    const before = texts(await opened(words));
    const reviewer = await deleted(words);
    reviewer.rejectAll();
    expect(texts(reviewer)).toEqual(before);
    expect(reviewer.getChanges()).toEqual([]);
  });

  test(`accepting every change leaves none after ${shape}`, async () => {
    const reviewer = await deleted(words);
    reviewer.acceptAll();
    expect(texts(reviewer).join("")).toBe("");
    expect(reviewer.getChanges()).toEqual([]);
  });

  test(`rejecting each listed change gives the words back after ${shape}`, async () => {
    const before = texts(await opened(words));
    const reviewer = await deleted(words);
    for (
      let change = reviewer.getChanges().at(0);
      change !== undefined;
      change = reviewer.getChanges().at(0)
    )
      expect(reviewer.rejectChange(change)).toBe(true);
    expect(texts(reviewer)).toEqual(before);
  });
}
