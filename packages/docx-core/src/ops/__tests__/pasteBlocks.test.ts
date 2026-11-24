/**
 * Clipboard blocks with a table: pasted directly, the table stands between
 * the text before and after the caret; suggested, accepting reads as the
 * direct paste and rejecting restores the source.
 */
import { expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  Run,
  Table,
  TableCell,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (text: string, paraId?: string): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: text === "" ? [] : [run(text)],
});
const cell = (text: string): TableCell => ({
  type: "tableCell",
  content: [paragraph(text, "0000AA01")],
});
const table = (): Table => ({
  type: "table",
  rows: [
    { type: "tableRow", cells: [cell("A1"), cell("B1")] },
    { type: "tableRow", cells: [cell("A2"), cell("B2")] },
  ],
});

const source = (text: string): Document => ({
  package: {
    document: {
      content: [paragraph(text, "00000001"), paragraph("next", "00000002")],
    },
  },
});

const SUGGESTING = {
  type: "suggesting",
  revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
  newIds: { revision: Array.from({ length: 80 }, (_, index) => 91 + index) },
} as const satisfies EditorIntentMode;

type PasteCase = {
  text: string;
  offset: number;
  openStart: 0 | 1;
  openEnd: 0 | 1;
};

const paste = (
  { text, offset, openStart, openEnd }: PasteCase,
  mode: EditorIntentMode,
): Document => {
  const document = source(text);
  const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset };
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "pasteBlocks",
      from: at,
      to: at,
      blocks: [paragraph("Lead"), table(), paragraph("Trail")],
      openStart,
      openEnd,
    },
    mode,
  });
  if (compiled.isErr()) throw compiled.error;
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const revisionIds = (document: Document): number[] => {
  const ids = new Set<number>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    const info: unknown = Reflect.get(value, "info");
    if (typeof info === "object" && info !== null) {
      const id: unknown = Reflect.get(info, "id");
      if (typeof id === "number") ids.add(id);
    }
    for (const child of Object.values(value)) visit(child);
  };
  visit(document.package.document.content);
  return [...ids];
};

const resolved = (
  document: Document,
  decision: (typeof REVISION_DECISIONS)[keyof typeof REVISION_DECISIONS],
): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: revisionIds(document),
    decision,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

/** Each block's text: a paragraph's, or a table's cells in rows. */
const shape = (document: Document): unknown[] => {
  const textOf = (block: Paragraph): string =>
    block.content
      .flatMap((item) =>
        item.type === "run"
          ? item.content.flatMap((piece) =>
              piece.type === "text" ? [piece.text] : [],
            )
          : [],
      )
      .join("");
  return document.package.document.content.map((block) => {
    if (block.type === "paragraph") return textOf(block);
    if (block.type !== "table") return block.type;
    return block.rows.map((row) =>
      row.cells.map((each) =>
        each.content
          .flatMap((inner) =>
            inner.type === "paragraph" ? [textOf(inner)] : [],
          )
          .join("|"),
      ),
    );
  });
};

const CELLS = [
  ["A1", "B1"],
  ["A2", "B2"],
];

for (const [name, pasteCase, expected] of [
  [
    "mid-paragraph, open edges",
    { text: "abcdef", offset: 3, openStart: 1, openEnd: 1 },
    ["abcLead", CELLS, "Traildef", "next"],
  ],
  [
    "mid-paragraph, closed edges",
    { text: "abcdef", offset: 3, openStart: 0, openEnd: 0 },
    ["abc", "Lead", CELLS, "Trail", "def", "next"],
  ],
  [
    "at a paragraph's start",
    { text: "abc", offset: 0, openStart: 0, openEnd: 0 },
    ["Lead", CELLS, "Trail", "abc", "next"],
  ],
  [
    "in an empty paragraph",
    { text: "", offset: 0, openStart: 1, openEnd: 1 },
    ["Lead", CELLS, "Trail", "next"],
  ],
] as const) {
  test(`${name}: direct, accepted and rejected`, () => {
    const direct = paste(pasteCase, { type: "editing" });
    expect(shape(direct)).toEqual([...expected]);
    const suggested = paste(pasteCase, SUGGESTING);
    expect(revisionIds(suggested).length).toBeGreaterThan(0);
    expect(shape(resolved(suggested, REVISION_DECISIONS.ACCEPT))).toEqual(
      shape(direct),
    );
    expect(shape(resolved(suggested, REVISION_DECISIONS.REJECT))).toEqual(
      shape(source(pasteCase.text)),
    );
  });
}
