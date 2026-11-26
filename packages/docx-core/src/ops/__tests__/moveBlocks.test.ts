/**
 * Blocks with a table dragged to another place: directly, the table leaves
 * its source and stands at the target; suggested, the source is recorded as
 * deleted and the copy as inserted, accepting leaving the moved table at the
 * target alone and rejecting restoring the source.
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
import {
  compileEditorIntent,
  type EditorIntent,
  type EditorIntentMode,
} from "../editorIntent";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type TextPosition,
} from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (text: string, paraId?: string): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: text === "" ? [] : [run(text)],
});
const cell = (text: string, paraId?: string): TableCell => ({
  type: "tableCell",
  content: [paragraph(text, paraId)],
});
/** The source table carries its own ids; a dragged copy carries none. */
const table = (ids: boolean): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [
        cell("A1", ids ? "0000AA01" : undefined),
        cell("B1", ids ? "0000AA02" : undefined),
      ],
    },
  ],
});
const CELLS = [["A1", "B1"]];

const source = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("zero", "00000001"),
        paragraph("one", "00000002"),
        table(true),
        paragraph("two", "00000003"),
        paragraph("three", "00000004"),
      ],
    },
  },
});

const SUGGESTING = {
  type: "suggesting",
  revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
  newIds: { revision: Array.from({ length: 80 }, (_, index) => 91 + index) },
} as const satisfies EditorIntentMode;

const at = (blockId: string, offset: number): TextPosition => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset,
});

/** The table alone, dragged from the end of "one" to the start of "two". */
const moveTable = (target: TextPosition): EditorIntent => ({
  type: "moveBlocks",
  from: at("00000002", 3),
  to: at("00000003", 0),
  target,
  blocks: [paragraph(""), table(false), paragraph("")],
  openStart: 1,
  openEnd: 1,
});

const compiled = (
  document: Document,
  intent: EditorIntent,
  mode: EditorIntentMode,
) => {
  const result = compileEditorIntent(document, { intent, mode });
  if (result.isErr()) throw result.error;
  return result.value;
};

const moved = (intent: EditorIntent, mode: EditorIntentMode) => {
  const document = source();
  const { ops } = compiled(document, intent, mode);
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) throw applied.error;
  return applied.value;
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

for (const [name, target, direct, accepted] of [
  [
    "down",
    at("00000004", 2),
    ["zero", "onetwo", "th", CELLS, "ree"],
    ["zero", "one", "two", "th", CELLS, "ree"],
  ],
  [
    "up",
    at("00000001", 2),
    ["ze", CELLS, "ro", "onetwo", "three"],
    ["ze", CELLS, "ro", "one", "two", "three"],
  ],
] as const) {
  test(`a table dragged ${name} moves directly, and suggested is accepted or rejected`, () => {
    const editing = moved(moveTable(target), { type: "editing" });
    expect(shape(editing.document)).toEqual([...direct]);
    // The inverse puts the source back exactly.
    const undone = applyDocumentOps(editing.document, editing.inverse);
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toEqual(source());

    const suggestion = moved(moveTable(target), SUGGESTING).document;
    expect(revisionIds(suggestion).length).toBeGreaterThan(0);
    // Both tables stand while the suggestion is pending.
    expect(
      suggestion.package.document.content.filter(
        (block) => block.type === "table",
      ),
    ).toHaveLength(2);
    // A suggested deletion keeps the paragraphs beside the table apart.
    expect(shape(resolved(suggestion, REVISION_DECISIONS.ACCEPT))).toEqual([
      ...accepted,
    ]);
    expect(shape(resolved(suggestion, REVISION_DECISIONS.REJECT))).toEqual(
      shape(source()),
    );
  });
}

test("a table moved with the text around it takes that text along", () => {
  const document = source();
  const intent: EditorIntent = {
    type: "moveBlocks",
    from: at("00000002", 1),
    to: at("00000003", 1),
    target: at("00000004", 5),
    blocks: [paragraph("ne"), table(false), paragraph("t")],
    openStart: 1,
    openEnd: 1,
  };
  const applied = applyDocumentOps(
    document,
    compiled(document, intent, { type: "editing" }).ops,
  );
  if (applied.isErr()) throw applied.error;
  expect(shape(applied.value.document)).toEqual([
    "zero",
    "owo",
    "threene",
    CELLS,
    "t",
  ]);
});

test("a table dropped into itself or onto the range it moves with moves nothing", () => {
  for (const target of [at("0000AA02", 1), at("00000002", 3)])
    for (const mode of [{ type: "editing" } as const, SUGGESTING])
      expect(compiled(source(), moveTable(target), mode).ops).toEqual([]);
});

test("moved blocks without a table move as a fragment", () => {
  const document = source();
  const intent: EditorIntent = {
    type: "moveBlocks",
    from: at("00000001", 0),
    to: at("00000001", 2),
    target: at("00000004", 5),
    blocks: [paragraph("ze")],
    openStart: 1,
    openEnd: 1,
  };
  const { blocks: _blocks, ...range } = intent;
  expect(compiled(document, intent, { type: "editing" })).toEqual(
    compiled(
      document,
      { ...range, type: "moveFragment", paragraphs: [paragraph("ze")] },
      { type: "editing" },
    ),
  );
});
