import { panic } from "better-result";
import { expect, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { storyParagraphs } from "../blocks";
import { normalizeForOps } from "../contract";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import { IDENTITY_SPACES, identityKeysIn, slotKey } from "../ids";
import { planRevisionResolution } from "../resolve";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type ResolveRevisionOp,
} from "../types";

const paragraphAt = (document: Document, id: string) =>
  storyParagraphs(document.package.document).find(
    ({ paragraph }) => paragraph.paraId === id,
  )?.paragraph ?? panic(`Missing fixture paragraph ${id}.`);

const accept = (revisionIds: readonly number[]) =>
  ({
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision: "accept",
  }) as const satisfies ResolveRevisionOp;

const resolveWithReceivers = (document: Document, op: ResolveRevisionOp) => {
  const before = structuredClone(document);
  const planned = planRevisionResolution({
    document,
    op,
    applyOps: applyDocumentOps,
  }).unwrap();
  const selected = new Set(
    op.revisionIds.map((id) =>
      slotKey({ space: IDENTITY_SPACES.REVISION, id }),
    ),
  );
  const allowed = new Set(
    planned.receiverEffects.map(({ receiver }) => receiver.paraId),
  );
  for (const { paragraph } of storyParagraphs(document.package.document))
    if (identityKeysIn(paragraph).some((key) => selected.has(key)))
      allowed.add(paragraph.paraId);
  const applied = applyDocumentOp(document, op).unwrap();
  assertExactModel(
    {
      document: applied.document,
      inverse: applied.inverse,
      touched: applied.touched,
    },
    planned.edit,
  );
  assertExactModel(document, before);
  const touched = new Set([
    ...applied.touched.modified,
    ...applied.touched.removed,
    ...applied.touched.inserted,
  ]);
  for (const id of touched) expect(allowed.has(id)).toBe(true);
  for (const { paragraph } of storyParagraphs(document.package.document))
    if (!touched.has(paragraph.paraId ?? ""))
      expect(paragraphAt(applied.document, paragraph.paraId ?? "")).toBe(
        paragraph,
      );
  const replayed: DocumentOp[] = JSON.parse(JSON.stringify(applied.inverse));
  const undone = applyDocumentOps(applied.document, replayed).unwrap();
  assertExactModel(undone.document, document);
  assertExactModel(
    applyDocumentOps(undone.document, undone.inverse).unwrap().document,
    applied.document,
  );
  return { applied, effects: planned.receiverEffects };
};

const cellReceiver = {
  type: "paragraph",
  paraId: "00000100",
  content: [],
} satisfies Paragraph;
const sectionSibling = {
  type: "paragraph",
  paraId: "00000101",
  content: [],
  sectionProperties: { pageWidth: 11906 },
} satisfies Paragraph;
const receivingTable = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [cellReceiver, sectionSibling] }],
    },
  ],
} satisfies Table;
const tail = {
  type: "paragraph",
  paraId: "00000200",
  content: [],
} satisfies Paragraph;

test("resolution declares the transitive table receiver after retiring an empty predecessor", () => {
  const document = normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "run",
                content: [{ type: "text", text: "a" }],
                propertyChanges: [
                  {
                    type: "runPropertyChange",
                    info: { id: 1, author: "A" },
                    previousFormatting: { italic: true },
                  },
                ],
              },
            ],
            propertyChanges: [
              {
                type: "paragraphPropertyChange",
                info: { id: 2, author: "A" },
                previousFormatting: { alignment: "end" },
              },
            ],
            pPrMark: { kind: "del", info: { id: 3, author: "A" } },
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content: [],
            pPrMark: { kind: "del", info: { id: 4, author: "A" } },
          },
          receivingTable,
          tail,
        ],
      },
    },
  });
  const { applied, effects } = resolveWithReceivers(
    document,
    accept([1, 2, 3, 4]),
  );
  expect(
    new Set(effects.map(({ type, receiver }) => `${type}:${receiver.paraId}`)),
  ).toEqual(new Set([`table-run-in:${cellReceiver.paraId}`]));
  expect(applied.touched.modified).toEqual([cellReceiver.paraId]);
  expect(new Set(applied.touched.removed)).toEqual(
    new Set(["00000001", "00000002"]),
  );
  expect(applied.touched.inserted).toEqual([]);
  expect(paragraphAt(applied.document, cellReceiver.paraId).content).toEqual([
    { type: "run", content: [{ type: "text", text: "a" }] },
  ]);
  expect(paragraphAt(applied.document, sectionSibling.paraId)).toBe(
    paragraphAt(document, sectionSibling.paraId),
  );
  expect(paragraphAt(applied.document, tail.paraId)).toBe(
    paragraphAt(document, tail.paraId),
  );
});

const CONTAINERS = {
  body: (blocks: BlockContent[]): BlockContent[] => blocks,
  cell: (blocks: BlockContent[]): BlockContent[] => [
    {
      type: "table",
      rows: [
        { type: "tableRow", cells: [{ type: "tableCell", content: blocks }] },
      ],
    },
    { type: "paragraph", paraId: "00000300", content: [] },
  ],
  sdt: (blocks: BlockContent[]): BlockContent[] => [
    {
      type: "blockSdt",
      properties: { sdtType: "group", id: 7 },
      content: blocks,
    },
    { type: "paragraph", paraId: "00000300", content: [] },
  ],
};

test("compiled chains declare their actual receiver across containers and selected stops", () => {
  const exercised = new Set<string>();
  for (const [container, wrap] of Object.entries(CONTAINERS)) {
    for (const length of [2, 4]) {
      for (const selectLast of [false, true]) {
        const paragraphs = Array.from({ length }, (_, index): Paragraph => ({
          type: "paragraph",
          paraId: (index + 1).toString(16).padStart(8, "0"),
          content:
            index === 0
              ? [{ type: "run", content: [{ type: "text", text: "a" }] }]
              : [],
          ...(index === length - 1
            ? { pPrMark: { kind: "del", info: { id: 1, author: "Earlier" } } }
            : {}),
        }));
        let document = normalizeForOps({
          package: {
            document: { content: wrap([...paragraphs, receivingTable, tail]) },
          },
        });
        const ids: number[] = selectLast ? [1] : [];
        for (let index = 0; index < length - 1; index += 1) {
          const intent = {
            type: "joinParagraphs",
            story: OP_STORIES.MAIN,
            blockId:
              paragraphs.at(index)?.paraId ?? panic("A predecessor exists."),
            nextBlockId:
              paragraphs.at(index + 1)?.paraId ?? panic("A successor exists."),
          } as const;
          const allocation = allocateEditorIntentIds(document, intent);
          const compiled = compileEditorIntent(document, {
            intent,
            mode: {
              type: "suggesting",
              newIds: allocation.newIds,
              revision: {
                id: allocation.revisionId,
                author: "Reviewer",
                date: "2026-03-04T05:00:00Z",
              },
            },
          }).unwrap();
          const applied = applyDocumentOps(document, compiled.ops).unwrap();
          ids.push(...applied.revisions);
          document = applied.document;
        }
        const { applied, effects } = resolveWithReceivers(
          document,
          accept(ids),
        );
        const tableEffects = effects.filter(
          ({ type }) => type === "table-run-in",
        );
        expect(
          new Set(tableEffects.map(({ receiver }) => receiver.paraId)),
        ).toEqual(new Set(selectLast ? [cellReceiver.paraId] : []));
        expect(paragraphAt(applied.document, sectionSibling.paraId)).toBe(
          paragraphAt(document, sectionSibling.paraId),
        );
        exercised.add(container);
      }
    }
  }
  expect(exercised).toEqual(new Set(Object.keys(CONTAINERS)));
});

test("one validated nested target owns table eligibility and reconstruction", () => {
  const document = normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
            pPrMark: { kind: "del", info: { id: 1, author: "A" } },
          },
          {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [
                      receivingTable,
                      { type: "paragraph", paraId: "00000102", content: [] },
                    ],
                  },
                ],
              },
            ],
          },
          tail,
        ],
      },
    },
  });
  const { effects } = resolveWithReceivers(document, accept([1]));
  expect(effects.map(({ receiver }) => receiver.paraId)).toEqual([
    cellReceiver.paraId,
  ]);
});

test("terminal removal declares both retained boundaries and its exact section recipient", () => {
  const document = normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              { type: "bookmarkStart", id: 7, name: "_Ref7" },
              { type: "run", content: [{ type: "text", text: "a" }] },
            ],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content: [{ type: "bookmarkEnd", id: 7 }],
            sectionProperties: { pageWidth: 11906 },
            pPrMark: { kind: "ins", info: { id: 1, author: "A" } },
          },
        ],
      },
    },
  });
  const { effects } = resolveWithReceivers(document, {
    ...accept([1]),
    decision: "reject",
  });
  expect(effects.map(({ type, receiver }) => [type, receiver.paraId])).toEqual([
    ["retained-range-boundaries", "00000001"],
    ["section-transfer", "00000001"],
  ]);
});

const BLOCKERS = {
  source: { sourceSection: { pageWidth: 11906 } },
  receiver: { receiverSection: { pageWidth: 11906 } },
  row: {
    rowChange: {
      type: "tableRowInsertion",
      info: { id: 10, author: "Pending" },
    },
  },
  cell: {
    cellChange: {
      type: "tableCellInsertion",
      info: { id: 10, author: "Pending" },
    },
  },
} as const;

test("blocked receiver classes refuse without inventing effects or changing the input", () => {
  const exercised = new Set<string>();
  for (const [name, blocker] of Object.entries(BLOCKERS)) {
    const document = normalizeForOps({
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [
                { type: "run", content: [{ type: "text", text: "a" }] },
              ],
              pPrMark: { kind: "del", info: { id: 1, author: "A" } },
              ...("sourceSection" in blocker
                ? { sectionProperties: blocker.sourceSection }
                : {}),
            },
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  ...("rowChange" in blocker
                    ? { structuralChange: blocker.rowChange }
                    : {}),
                  cells: [
                    {
                      type: "tableCell",
                      ...("cellChange" in blocker
                        ? { structuralChange: blocker.cellChange }
                        : {}),
                      content: [
                        {
                          type: "paragraph",
                          paraId: "00000100",
                          content: [],
                          ...("receiverSection" in blocker
                            ? { sectionProperties: blocker.receiverSection }
                            : {}),
                        },
                      ],
                    },
                  ],
                },
              ],
            },
            tail,
          ],
        },
      },
    });
    const before = structuredClone(document);
    const planned = planRevisionResolution({
      document,
      op: accept([1]),
      applyOps: applyDocumentOps,
    });
    const applied = applyDocumentOp(document, accept([1]));
    expect(planned.isErr() && planned.error.reason).toBe("untrackable");
    expect(applied.isErr() && applied.error.reason).toBe("untrackable");
    assertExactModel(document, before);
    exercised.add(name);
  }
  expect(exercised).toEqual(new Set(Object.keys(BLOCKERS)));
});
