import { panic } from "better-result";
import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import * as blocks from "../blocks";
import { idKey } from "../ids";
import { paragraphLength } from "../offsets";
import { planTrackedDeletion, selectedParagraphRuns } from "../plan";
import { rangeEndAfterDeletion } from "../rangeReplacement";
import { DOCUMENT_OP_TYPES, OP_STORIES, type TextPosition } from "../types";

const paragraph = (text: string, paraId?: string): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const documentOf = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});
const at = (blockId: string, offset: number): TextPosition => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

// The prior implementation built a last-wins map for every range. These
// fixtures contain only text, so positions have no zero-width leaves.
const oldEndpoint = ({
  before,
  after,
  from,
  to,
}: Parameters<typeof rangeEndAfterDeletion>[0]): TextPosition => {
  const run = selectedParagraphRuns(before, from, to).unwrap().at(0) ?? [];
  const source = run.at(-1)?.paragraph;
  const byId = new Map(
    blocks
      .storyParagraphs(after.package.document)
      .map(({ paragraph: item }) => [idKey(item.paraId ?? ""), item]),
  );
  const holder = run.findLast(({ paragraph: item }) =>
    byId.has(idKey(item.paraId ?? "")),
  )?.paragraph;
  const target = byId.get(idKey(holder?.paraId ?? ""));
  if (source === undefined || holder === undefined || target === undefined)
    panic("Expected a surviving source");
  const trailing =
    idKey(source.paraId ?? "") === idKey(to.blockId)
      ? paragraphLength(source) - to.offset
      : 0;
  return {
    story: from.story,
    blockId: holder.paraId ?? to.blockId,
    offset: paragraphLength(target) - trailing,
    zeroWidthBefore: 0,
  };
};

test.each([64, 256])(
  "retained tracked endpoints avoid enumerating the %i-paragraph edited body",
  (width) => {
    const before = documentOf([
      paragraph("abcd", "000000AA"),
      paragraph("efghi", "000000BB"),
      paragraph("jklmn", "000000CC"),
      ...Array.from({ length: width - 3 }, (_, index) =>
        paragraph("untouched", String(index + 1).padStart(8, "0")),
      ),
    ]);
    const from = at("000000aa", 1);
    const to = at("000000cc", 2);
    const ops = planTrackedDeletion(before, {
      from,
      to,
      revision: { id: 1000, author: "Author", date: "2026-01-11T00:00:00Z" },
      newIds: { revision: [1001, 1002, 1003, 1004, 1005] },
    }).unwrap();
    const edited = applyDocumentOps(before, ops).unwrap();
    const args = { before, after: edited.document, from, to };
    const expected = oldEndpoint(args);
    expect(expected.blockId).toBe("000000CC");
    const walks = spyOn(blocks, "storyParagraphs");
    try {
      expect(rangeEndAfterDeletion(args)).toEqual(expected);
      expect(
        walks.mock.calls.filter(
          ([body]) => body === args.after.package.document,
        ),
      ).toHaveLength(0);
    } finally {
      walks.mockRestore();
    }
    expect(
      applyDocumentOps(edited.document, edited.inverse).unwrap().document,
    ).toEqual(before);
  },
);

test("deletion follows the earlier retained paragraph when the last ID is removed", () => {
  const first = paragraph("abcd", "000000AA");
  const last = paragraph("efghi", "000000BB");
  const before = documentOf([first, last, paragraph("jklmn", "000000CC")]);
  const from = at("000000aa", 1);
  const to = at("000000bb", 2);
  const edited = applyDocumentOp(before, {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: OP_STORIES.MAIN,
    expected: [first, last],
    blocks: [paragraph("aghi", "000000AA")],
  }).unwrap();
  const args = { before, after: edited.document, from, to };
  expect(oldEndpoint(args)).toEqual({
    ...at("000000AA", 1),
    zeroWidthBefore: 0,
  });
  expect(rangeEndAfterDeletion(args)).toEqual(oldEndpoint(args));
  expect(
    applyDocumentOps(edited.document, edited.inverse).unwrap().document,
  ).toEqual(before);
});

test.each([
  {
    name: "duplicate mixed-case IDs",
    before: documentOf([
      paragraph("a", "000000AA"),
      paragraph("last", "000000Bb"),
    ]),
    after: documentOf([
      paragraph("short", "000000bb"),
      paragraph("canonical last", "000000BB"),
    ]),
    to: at("000000bb", 2),
    expected: at("000000Bb", 12),
  },
  {
    name: "unnamed surviving paragraphs",
    before: documentOf([
      paragraph("a", "000000AA"),
      paragraph("interior"),
      paragraph("last", "000000BB"),
    ]),
    after: documentOf([paragraph("short", ""), paragraph("canonical last")]),
    to: at("000000bb", 2),
    expected: at("000000bb", 12),
  },
  {
    name: "removed unnamed paragraphs",
    before: documentOf([
      paragraph("retained", "000000AA"),
      paragraph("unnamed"),
      paragraph("empty ID", ""),
      paragraph("last", "000000BB"),
    ]),
    after: documentOf([paragraph("remaining", "000000AA")]),
    to: at("000000bb", 2),
    expected: at("000000AA", 7),
  },
])(
  "preserves the last physical occurrence for $name",
  ({ before, after, to, expected }) => {
    const args = { before, after, from: at("000000aa", 0), to };
    const previous = oldEndpoint(args);
    expect(previous).toEqual({ ...expected, zeroWidthBefore: 0 });
    const walks = spyOn(blocks, "storyParagraphs");
    try {
      expect(rangeEndAfterDeletion(args)).toEqual(previous);
      expect(
        walks.mock.calls.filter(([body]) => body === after.package.document),
      ).toHaveLength(1);
    } finally {
      walks.mockRestore();
    }
  },
);
