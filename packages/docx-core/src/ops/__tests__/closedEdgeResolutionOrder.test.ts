/**
 * Changes beside content a rejection removes, resolved in every order. New
 * content joined to nothing has nothing beyond its edge: what was made inside
 * it takes the boundaries kept from its record and joins nothing past it. A
 * seam's markers wait on the changes that keep its pieces apart, and go where
 * the seam goes. Each shape stands alone, beside a range boundary, in a
 * control, in a table cell, at a section's end, and in a paragraph whose own
 * break the same resolutions take.
 */
import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const ACCEPT = REVISION_DECISIONS.ACCEPT;
const REJECT = REVISION_DECISIONS.REJECT;

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const tab: ParagraphContent = { type: "run", content: [{ type: "tab" }] };

const link = (content: ParagraphContent[]): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children: content,
});

const embedded = (content: ParagraphContent[]): ParagraphContent => ({
  type: "inlineWrapper",
  kind: "bidi",
  control: "embedding",
  content,
});

const start = (id: number): ParagraphContent => ({
  type: "bookmarkStart",
  id,
  name: `_Ref${id}`,
});
const end = (id: number): ParagraphContent => ({ type: "bookmarkEnd", id });

type Joins = {
  before: number;
  after: number;
  remove: number;
};

const NOTHING: Joins = { before: 0, after: 0, remove: 0 };

const dateOf = (index: number): string => `2026-03-04T05:00:0${index}Z`;

const insertion = (
  id: number,
  date: number,
  content: ParagraphContent[],
  joins: Joins = NOTHING,
): ParagraphContent => ({
  type: "insertion",
  info: { id, author: "A", date: dateOf(date) },
  content,
  resolutionJoins: joins,
});

/** An insertion another producer wrote, which records no joins. */
const authored = (
  id: number,
  content: ParagraphContent[],
): ParagraphContent => ({
  type: "insertion",
  info: { id, author: "B" },
  content,
});

const deletion = (
  id: number,
  date: number,
  content: ParagraphContent[],
  joins: Joins = NOTHING,
): ParagraphContent => ({
  type: "deletion",
  info: { id, author: "A", date: dateOf(date) },
  content,
  resolutionJoins: joins,
});

const added = (
  id: number,
  date: number,
  resolutionJoin: number,
): NonNullable<Paragraph["pPrMark"]> => ({
  kind: "ins",
  info: { id, author: "A", date: dateOf(date) },
  resolutionJoin,
});

const paragraph = (
  paraId: string,
  content: ParagraphContent[],
  fields: Partial<Paragraph> = {},
): Paragraph => ({ type: "paragraph", paraId, content, ...fields });

type Group = { ids: number[]; decision: RevisionDecision };

const resolved = (document: Document, { ids, decision }: Group): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const contentsOf = (document: Document) => {
  const contents: ParagraphContent[][] = [];
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") contents.push(block.content);
      else if (block.type === "blockSdt") visit(block.content);
      else if (block.type === "table")
        for (const row of block.rows)
          for (const cell of row.cells) visit(cell.content);
    }
  };
  visit(document.package.document.content);
  return contents;
};

const ordersOf = (groups: readonly Group[]): Group[][] =>
  groups.length === 0
    ? [[]]
    : groups.flatMap((group) =>
        ordersOf(groups.filter((other) => other !== group)).map((rest) =>
          [group].concat(rest),
        ),
      );

const resolvedIn = (document: Document, order: readonly Group[]): Document => {
  let current = document;
  for (const group of order) current = resolved(current, group);
  return current;
};

/** The batches of one decision each, accepted first and rejected first. */
const batchesOf = (groups: readonly Group[]): Group[][] => {
  const batch = (decision: RevisionDecision): Group[] => {
    const ids = groups.flatMap((group) =>
      group.decision === decision ? group.ids : [],
    );
    return ids.length === 0 ? [] : [{ ids, decision }];
  };
  return [
    [...batch(ACCEPT), ...batch(REJECT)],
    [...batch(REJECT), ...batch(ACCEPT)],
  ];
};

/** What every order of resolving `groups` leaves, and both batches. */
const leftBy = (document: Document, groups: readonly Group[]) => {
  const [first, ...others] = [...ordersOf(groups), ...batchesOf(groups)].map(
    (order) => ({ order, contents: contentsOf(resolvedIn(document, order)) }),
  );
  if (first === undefined) return [];
  for (const { order, contents } of others)
    expect({ order, contents }).toEqual({ order, contents: first.contents });
  return first.contents;
};

type Shape = {
  content: ParagraphContent[];
  groups: Group[];
  /** The paragraph's content once everything is resolved. */
  left: ParagraphContent[];
};

const reject = (...ids: number[]): Group => ({ ids, decision: REJECT });
const accept = (...ids: number[]): Group => ({ ids, decision: ACCEPT });

const shapes: Record<string, Shape> = {
  "content made inside rejected content joins nothing past it": {
    content: [
      link([embedded([])]),
      insertion(1, 1, [start(1)]),
      insertion(2, 2, [link([text("x")])], { before: 1, after: 3, remove: 1 }),
      insertion(3, 3, [link([text("y")])], { before: 4, after: 0, remove: 0 }),
      insertion(4, 1, [link([tab])]),
    ],
    groups: [reject(1, 4), accept(2), accept(3)],
    left: [link([embedded([])]), start(1), link([text("xy")])],
  },
  "boundaries kept from its record go into the content made inside it": {
    content: [
      insertion(1, 2, [link([text("x")])], { before: 0, after: 4, remove: 0 }),
      insertion(2, 1, [link([text("a"), start(1)])]),
      deletion(3, 1, [link([text("b")])]),
    ],
    groups: [reject(2), accept(1), reject(3)],
    left: [link([text("x"), start(1)]), link([text("b")])],
  },
  "an authored link beyond new content's edge is no piece of its record": {
    content: [
      link([]),
      insertion(1, 1, [link([text("a")])]),
      insertion(2, 2, [link([text("x")])], { before: 4, after: 3, remove: 3 }),
      insertion(3, 1, [link([text("b")]), end(1)]),
    ],
    groups: [reject(1, 3), accept(2)],
    left: [link([]), link([text("x")]), end(1)],
  },
  "a removed change's cut is read past later insertions inside it": {
    content: [
      embedded([text("w")]),
      insertion(1, 1, [text("a")]),
      insertion(2, 2, [text("b")], { before: 3, after: 2, remove: 2 }),
      insertion(3, 3, [embedded([text("x")])], {
        before: 0,
        after: 0,
        remove: 3,
      }),
      insertion(4, 2, [text("b")], { before: 3, after: 2, remove: 2 }),
      insertion(5, 1, [text("a"), tab]),
    ],
    groups: [reject(1, 5), reject(2, 4), accept(3)],
    left: [embedded([text("w")]), embedded([text("x")])],
  },
  "markers between two accepted pieces of one removal stand in its cut": {
    content: [
      insertion(4, 3, [text("x")]),
      deletion(1, 1, [link([text("a")])], { before: 0, after: 1, remove: 0 }),
      insertion(3, 2, [link([start(1)])], { before: 0, after: 0, remove: 4 }),
      deletion(2, 1, [link([text("a")])], { before: 0, after: 1, remove: 0 }),
      link([embedded([text("b")])]),
    ],
    groups: [accept(1, 2), reject(3), accept(4)],
    left: [text("x"), link([start(1), embedded([text("b")])])],
  },
  "markers between two accepted pieces of a removal that cut nothing stand": {
    content: [
      insertion(4, 3, [text("x")]),
      deletion(1, 1, [tab]),
      insertion(3, 2, [start(1)], { before: 0, after: 0, remove: 1 }),
      deletion(2, 1, [text("a")]),
      authored(5, [link([text("b")])]),
    ],
    groups: [accept(1, 2), reject(3), accept(4)],
    left: [
      text("x"),
      {
        type: "bookmarkStart",
        id: 1,
        name: "_Ref1",
        keptCut: { keptFrom: dateOf(2) },
      },
      authored(5, [link([text("b")])]),
    ],
  },
  "markers beyond the record an insertion was made in go into the cut's piece":
    {
      content: [
        embedded([tab]),
        insertion(1, 1, [text("aa")], { before: 0, after: 0, remove: 2 }),
        insertion(2, 2, [text("x")], { before: 3, after: 2, remove: 2 }),
        insertion(3, 1, [tab, link([start(1)])], {
          before: 0,
          after: 0,
          remove: 2,
        }),
        embedded([text("a")]),
      ],
      groups: [reject(1, 3), accept(2)],
      left: [embedded([tab]), text("x"), embedded([start(1), text("a")])],
    },
  "a cancelled join hands nothing on when its insertion goes": {
    content: [
      insertion(1, 1, [link([embedded([end(1)])])]),
      insertion(4, 4, [link([embedded([text("z")])])], {
        before: 0,
        after: 5,
        remove: 0,
      }),
      insertion(2, 2, [link([embedded([text("x")])])], {
        before: 3,
        after: 3,
        remove: 3,
      }),
      insertion(3, 3, [link([embedded([text("y")])])], {
        before: 5,
        after: 0,
        remove: 0,
      }),
      insertion(5, 1, [link([embedded([text("a")])])]),
    ],
    groups: [reject(1, 5), reject(2), accept(3), reject(4)],
    left: [end(1), link([embedded([text("y")])])],
  },
  "markers kept from inside rejected new content stand where it was": {
    content: [
      insertion(1, 1, [text("a")]),
      insertion(2, 2, [start(1)], { before: 0, after: 0, remove: 3 }),
      insertion(3, 3, [text("x")], { before: 1, after: 1, remove: 1 }),
      insertion(4, 2, [tab], { before: 0, after: 0, remove: 3 }),
      insertion(5, 1, [text("a")]),
      link([text("b")]),
    ],
    groups: [reject(1, 5), reject(2, 4), reject(3)],
    left: [start(1), link([text("b")])],
  },
  "boundaries with no record of their own go into the piece they join": {
    content: [
      insertion(1, 3, [link([text("x")])], { before: 0, after: 5, remove: 0 }),
      insertion(2, 4, [end(1)]),
      insertion(3, 2, [link([text("y")])], { before: 0, after: 5, remove: 0 }),
      insertion(4, 1, [link([text("a")])]),
    ],
    groups: [accept(1), reject(2), reject(3), reject(4)],
    left: [link([text("x"), end(1)])],
  },
};

const tailParagraph = paragraph("0000000F", []);

const contexts: Record<
  string,
  {
    blocks: (content: ParagraphContent[]) => BlockContent[];
    groups?: Group[];
    left: (content: ParagraphContent[]) => ParagraphContent[][];
  }
> = {
  alone: {
    blocks: (content) => [paragraph("00000001", content)],
    left: (content) => [content],
  },
  "between range boundaries": {
    blocks: (content) => [
      paragraph("00000001", [start(8), ...content, end(8)]),
    ],
    left: (content) => [[start(8), ...content, end(8)]],
  },
  "in a control": {
    blocks: (content) => [
      paragraph("00000001", [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 80, tag: "clause" },
          content,
        },
      ]),
    ],
    left: (content) => [
      [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 80, tag: "clause" },
          content,
        },
      ],
    ],
  },
  "in a table cell": {
    blocks: (content) => [
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              { type: "tableCell", content: [paragraph("00000001", content)] },
            ],
          },
        ],
      },
      tailParagraph,
    ],
    left: (content) => [content, []],
  },
  "at a section's end": {
    blocks: (content) => [
      paragraph("00000001", content, {
        sectionProperties: { pageWidth: 12_240 },
      }),
      tailParagraph,
    ],
    left: (content) => [content, []],
  },
  "before a paragraph joining it in the same resolutions": {
    blocks: (content) => [
      paragraph("00000001", content, { pPrMark: added(90, 5, 0) }),
      paragraph("00000002", [text("next")]),
    ],
    groups: [reject(90)],
    left: (content) => [[...content, text("next")]],
  },
  "after a paragraph joining it in the same resolutions": {
    blocks: (content) => [
      paragraph("00000002", [text("previous")], { pPrMark: added(90, 5, 0) }),
      paragraph("00000001", content),
    ],
    groups: [reject(90)],
    left: (content) => [[text("previous"), ...content]],
  },
};

const control = (id: number, value: string): ParagraphContent => ({
  type: "inlineSdt",
  properties: { sdtType: "richText", id, tag: "clause" },
  content: [text(value)],
});

/** Two paragraphs a tracked split (revision 1) left, resolved in every order. */
type SplitShape = {
  first: ParagraphContent[];
  depth: number;
  second: ParagraphContent[];
  groups: Group[];
  left: ParagraphContent[];
};

const splitShapes: Record<string, SplitShape> = {
  "an insertion at the edge of a split's piece is part of the piece": {
    first: [link([text("a")]), insertion(3, 2, [link([start(1)])])],
    depth: 2,
    second: [
      insertion(4, 3, [link([text("x")])], { before: 0, after: 2, remove: 0 }),
      link([tab]),
    ],
    groups: [reject(1), reject(3), accept(4)],
    left: [
      link([
        text("a"),
        start(1),
        {
          type: "run",
          content: [{ type: "text", text: "x" }, { type: "tab" }],
        },
      ]),
    ],
  },
  "markers kept at a split's seam wait with the insertion beside them": {
    first: [control(1, "a"), insertion(4, 3, [end(2)])],
    depth: 3,
    second: [insertion(2, 2, [end(1)]), control(2, "b")],
    groups: [reject(1), reject(2), reject(4)],
    left: [
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 1, tag: "clause" },
        content: [text("a"), end(2), end(1), text("b")],
      },
    ],
  },
};

const splitContexts: Record<
  string,
  {
    blocks: (first: Paragraph, second: Paragraph) => BlockContent[];
    groups?: Group[];
    left: (content: ParagraphContent[]) => ParagraphContent[][];
  }
> = {
  alone: {
    blocks: (first, second) => [first, second],
    left: (content) => [content],
  },
  "between range boundaries": {
    blocks: (first, second) => [
      { ...first, content: [start(8), ...first.content] },
      { ...second, content: [...second.content, end(8)] },
    ],
    left: (content) => [[start(8), ...content, end(8)]],
  },
  "in a block control": {
    blocks: (first, second) => [
      {
        type: "blockSdt",
        properties: { sdtType: "richText", id: 90 },
        content: [first, second],
      },
      tailParagraph,
    ],
    left: (content) => [content, []],
  },
  "in a table cell": {
    blocks: (first, second) => [
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [{ type: "tableCell", content: [first, second] }],
          },
        ],
      },
      tailParagraph,
    ],
    left: (content) => [content, []],
  },
  "at a section's end": {
    blocks: (first, second) => [
      first,
      { ...second, sectionProperties: { pageWidth: 12_240 } },
      tailParagraph,
    ],
    left: (content) => [content, []],
  },
  "before a paragraph joining it in the same resolutions": {
    blocks: (first, second) => [
      first,
      { ...second, pPrMark: added(90, 5, 0) },
      paragraph("00000009", [text("next")]),
    ],
    groups: [reject(90)],
    left: (content) => [[...content, text("next")]],
  },
};

for (const [name, shape] of Object.entries(splitShapes))
  describe(name, () => {
    for (const [where, context] of Object.entries(splitContexts))
      test(where, () => {
        const document: Document = {
          package: {
            document: {
              content: context.blocks(
                paragraph("00000001", shape.first, {
                  pPrMark: added(1, 1, shape.depth),
                }),
                paragraph("00000002", shape.second),
              ),
            },
          },
        };
        expect(
          leftBy(document, [...shape.groups, ...(context.groups ?? [])]),
        ).toEqual(context.left(shape.left));
      });
  });

for (const [name, shape] of Object.entries(shapes))
  describe(name, () => {
    for (const [where, context] of Object.entries(contexts))
      test(where, () => {
        const document: Document = {
          package: { document: { content: context.blocks(shape.content) } },
        };
        expect(
          leftBy(document, [...shape.groups, ...(context.groups ?? [])]),
        ).toEqual(context.left(shape.left));
      });
  });
