/**
 * `containerStart` names the place before a block list's first block when
 * that block holds no paragraph to name it by, such as an empty content
 * control: paragraphs and tables go there, direct or tracked, in every kind
 * of container, and each inverse restores the source exactly.
 */
import { expect, test } from "bun:test";

import type {
  BlockContent,
  BlockSdt,
  Document,
  Paragraph,
  Table,
} from "../../model/document";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
} from "../apply";
import { storyBody } from "../blocks";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
  type EditorIntent,
} from "../editorIntent";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type InsertBlocksOp,
  type InsertTableOp,
  OP_STORIES,
  type OpStory,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const control = (content: BlockContent[]): BlockSdt => ({
  type: "blockSdt",
  properties: { sdtType: "richText" },
  content,
});
const emptyControl = (): BlockSdt => control([]);

/** A block content control first in its list, with or without a paragraph. */
const CONTROLS = [
  { name: "an empty control", control: emptyControl },
  {
    name: "a control holding a paragraph",
    control: () => control([paragraph("00000005", "Inside.")]),
  },
] as const;

const inCell = (content: BlockContent[]): Table => ({
  type: "table",
  rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
});

const textBoxHost = (paraId: string, boxed: BlockContent[]): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [
    {
      type: "run",
      content: [
        {
          type: "shape",
          shape: {
            type: "shape",
            shapeType: "textBox",
            size: { width: 2_000_000, height: 800_000 },
            textBody: { content: boxed },
          },
        },
      ],
    },
  ],
});

/** The list under test: a control first, then its final paragraph. */
const list = (first: BlockSdt): BlockContent[] => [
  first,
  paragraph("00000001", "Kept."),
];

type Container = {
  name: string;
  story: OpStory;
  document: (first: BlockSdt) => Document;
  /** The list under test, in `document`. */
  blocks: (document: Document) => readonly BlockContent[];
};

const header = { kind: "header", rId: "rIdHeader" } as const;
const footnote = { kind: "footnote", id: 1 } as const;

const CONTAINERS: readonly Container[] = [
  {
    name: "body",
    story: OP_STORIES.MAIN,
    document: (first) => ({ package: { document: { content: list(first) } } }),
    blocks: (document) => document.package.document.content,
  },
  {
    name: "cell",
    story: OP_STORIES.MAIN,
    document: (first) => ({
      package: {
        document: { content: [inCell(list(first)), paragraph("00000009", "")] },
      },
    }),
    blocks: (document) => {
      const table = document.package.document.content.at(0);
      return table?.type === "table"
        ? (table.rows.at(0)?.cells.at(0)?.content ?? [])
        : [];
    },
  },
  {
    name: "header",
    story: header,
    document: (first) => ({
      package: {
        document: { content: [paragraph("00000009", "")] },
        headers: new Map([
          [
            header.rId,
            { type: "header", hdrFtrType: "default", content: list(first) },
          ],
        ]),
      },
    }),
    blocks: (document) => storyBody(document, header).content,
  },
  {
    name: "footnote",
    story: footnote,
    document: (first) => ({
      package: {
        document: { content: [paragraph("00000009", "")] },
        footnotes: [{ type: "footnote", id: 1, content: list(first) }],
      },
    }),
    blocks: (document) => storyBody(document, footnote).content,
  },
  {
    name: "text box",
    story: OP_STORIES.MAIN,
    document: (first) => ({
      package: {
        document: { content: [textBoxHost("00000009", list(first))] },
      },
    }),
    blocks: (document) => {
      const host = document.package.document.content.at(0);
      const run = host?.type === "paragraph" ? host.content.at(0) : undefined;
      const shape = run?.type === "run" ? run.content.at(0) : undefined;
      return shape?.type === "shape"
        ? (shape.shape.textBody?.content ?? [])
        : [];
    },
  },
];

const stamp = { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" };
const newIds = {
  revision: Array.from({ length: 32 }, (_, index) => 101 + index),
};

const apply = (document: Document, op: DocumentOp): AppliedDocumentOp =>
  applyDocumentOp(document, op).unwrap();

const undo = (applied: AppliedDocumentOp): Document =>
  applyDocumentOps(applied.document, applied.inverse).unwrap().document;

const resolve = (
  applied: AppliedDocumentOp,
  story: OpStory,
  decision: RevisionDecision,
): Document =>
  apply(applied.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story,
    revisionIds: applied.revisions,
    decision,
  }).document;

const insertion = (story: OpStory): InsertBlocksOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story,
  at: { type: "containerStart", blockId: "00000001" },
  blocks: [paragraph("00000002", "Heading.")],
});

const tableInsertion = (story: OpStory): InsertTableOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_TABLE,
  story,
  at: { type: "containerStart", blockId: "00000001" },
  table: inCell([paragraph("00000003", "Cell.")]),
});

const bodySource = (): Document => ({
  package: { document: { content: list(emptyControl()) } },
});

const OPS = [
  { name: "paragraphs", op: insertion, first: "paragraph" },
  { name: "a table", op: tableInsertion, first: "table" },
] as const;

for (const container of CONTAINERS)
  for (const { name, op, first } of OPS)
    for (const starting of CONTROLS)
      test(`${name} go before ${starting.name} that starts a ${container.name}`, () => {
        const source = container.document(starting.control());
        const direct = apply(source, op(container.story));
        expect(
          container.blocks(direct.document).map(({ type }) => type),
        ).toEqual([first, "blockSdt", "paragraph"]);
        expect(undo(direct)).toEqual(source);
        const tracked = apply(source, {
          ...op(container.story),
          revision: stamp,
          newIds,
        });
        expect(tracked.revisions.length).toBeGreaterThan(0);
        expect(
          resolve(tracked, container.story, REVISION_DECISIONS.ACCEPT),
        ).toEqual(direct.document);
        expect(
          resolve(tracked, container.story, REVISION_DECISIONS.REJECT),
        ).toEqual(source);
        expect(undo(tracked)).toEqual(source);
      });

test("removing paragraphs before an empty control that starts its list puts them back there", () => {
  const source = bodySource();
  const inserted = apply(source, insertion(OP_STORIES.MAIN));
  const removed = applyDocumentOps(
    inserted.document,
    inserted.inverse,
  ).unwrap();
  expect(removed.document).toEqual(source);
  const restored = applyDocumentOps(removed.document, removed.inverse).unwrap();
  expect(restored.document).toEqual(inserted.document);
});

test("the start of a list that begins with a paragraph is before that paragraph", () => {
  const source = {
    package: {
      document: {
        content: [
          paragraph("00000004", "First."),
          emptyControl(),
          paragraph("00000001", "Kept."),
        ],
      },
    },
  } satisfies Document;
  for (const revision of [undefined, stamp]) {
    const tracked = revision === undefined ? {} : { revision, newIds };
    const atStart = apply(source, {
      ...insertion(OP_STORIES.MAIN),
      ...tracked,
    });
    const before = apply(source, {
      ...insertion(OP_STORIES.MAIN),
      at: { type: "before", blockId: "00000004" },
      ...tracked,
    });
    expect(atStart.document).toEqual(before.document);
  }
});

test("a list start names a paragraph the story holds", () => {
  const refused = applyDocumentOp(bodySource(), {
    ...insertion(OP_STORIES.MAIN),
    at: { type: "containerStart", blockId: "0000000F" },
  });
  expect(refused.isErr() && refused.error.reason).toBe("blockNotFound");
});

for (const starting of CONTROLS)
  test(`a tracked paragraph after a table ${starting.name} follows resolves exactly`, () => {
    const source = {
      package: {
        document: {
          content: [
            inCell([paragraph("00000006", "Cell.")]),
            starting.control(),
            paragraph("00000001", "Kept."),
          ],
        },
      },
    } satisfies Document;
    const op = {
      ...insertion(OP_STORIES.MAIN),
      at: { type: "afterTable", blockId: "00000006" },
    };
    const direct = apply(source, op);
    const tracked = apply(source, { ...op, revision: stamp, newIds });
    expect(
      resolve(tracked, OP_STORIES.MAIN, REVISION_DECISIONS.ACCEPT),
    ).toEqual(direct.document);
    expect(
      resolve(tracked, OP_STORIES.MAIN, REVISION_DECISIONS.REJECT),
    ).toEqual(source);
    expect(undo(tracked)).toEqual(source);
  });

const INTENTS = [
  {
    type: "insertParagraphs",
    story: OP_STORIES.MAIN,
    at: { type: "containerStart", blockId: "00000001" },
    paragraphs: [{ type: "paragraph", content: [] }],
  },
  {
    type: "tableBlock",
    operation: {
      type: DOCUMENT_OP_TYPES.INSERT_TABLE,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      anchor: "containerStart",
      table: inCell([{ type: "paragraph", content: [] }]),
    },
  },
] as const satisfies readonly EditorIntent[];

for (const intent of INTENTS)
  test(`the ${intent.type} intent compiles to one insertion at the list's start`, () => {
    const source = bodySource();
    const ids = createEditorIntentIdAllocator()(source, intent);
    const { ops } = compileEditorIntent(source, {
      intent,
      mode: { type: "editing", newIds: ids.newIds },
    }).unwrap();
    expect(ops.map((op) => ("at" in op ? op.at.type : op.type))).toEqual([
      "containerStart",
    ]);
    const applied = applyDocumentOps(source, ops).unwrap();
    expect(
      applied.document.package.document.content.map(({ type }) => type),
    ).toEqual([
      intent.type === "tableBlock" ? "table" : "paragraph",
      "blockSdt",
      "paragraph",
    ]);
  });
