import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import {
  type BlockContent,
  type Document,
  type Paragraph,
  type Table,
} from "../../model/document";
import {
  applyDocumentOp,
  applyDocumentOpEnvelope,
  applyDocumentOps,
} from "../apply";
import { storyParagraphs } from "../blocks";
import { contractViolation, normalizeForOps } from "../contract";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
} from "../editorIntent";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type EditorIntent,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  REVISION_DECISIONS,
  toOpEnvelope,
} from "../types";

type ParagraphOptions = { paraId: string; text: string };
const paragraph = ({ paraId, text }: ParagraphOptions): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});
const table = (paraId: string) =>
  ({
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          { type: "tableCell", content: [paragraph({ paraId, text: "Cell" })] },
        ],
      },
    ],
  }) satisfies Table;
const CONTAINERS = ["body", "cell", "sdt", "customXml"] as const;
type Container = (typeof CONTAINERS)[number];
const wrap = (
  content: BlockContent[],
  container: Container,
): BlockContent[] => {
  switch (container) {
    case "body":
      return content;
    case "cell":
      return [
        {
          type: "table",
          rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
        },
      ];
    case "sdt":
      return [
        {
          type: "blockSdt",
          properties: { sdtType: "richText", id: 77 },
          content,
        },
      ];
    case "customXml":
      return [
        {
          type: "blockCustomXml",
          openingXml: '<w:customXml w:element="clause">',
          closingXml: "</w:customXml>",
          content,
        },
      ];
    default: {
      const unreachable: never = container;
      return unreachable;
    }
  }
};
const source = (container: Container): Document =>
  normalizeForOps({
    package: {
      document: {
        content: wrap(
          [
            paragraph({ paraId: "10000001", text: "Left" }),
            table("10000011"),
            table("10000022"),
            paragraph({ paraId: "10000002", text: "Right" }),
          ],
          container,
        ),
      },
    },
  });
const text = fc.string({
  unit: fc.constantFrom("a", "b", " ", "ž", "😀", "א"),
  maxLength: 12,
});

test(
  "table-side paragraph insertion preserves source owners and exact JSON undo redo",
  async () => {
    await assertProperty(
      fc.property(
        fc.constantFrom(...CONTAINERS),
        fc.constantFrom("beforeTable", "afterTable"),
        fc.array(text, { minLength: 1, maxLength: 3 }),
        (container, side, words) => {
          const before = source(container),
            original = structuredClone(before);
          const intent = {
            type: "insertParagraphs",
            story: OP_STORIES.MAIN,
            at: {
              type: side,
              blockId: side === "beforeTable" ? "10000022" : "10000011",
            },
            paragraphs: words.map(
              (word) =>
                ({
                  type: "paragraph",
                  content:
                    word === ""
                      ? []
                      : [
                          {
                            type: "run",
                            content: [{ type: "text", text: word }],
                          },
                        ],
                }) satisfies Paragraph,
            ),
          } as const satisfies EditorIntent;
          const allocate = createEditorIntentIdAllocator();
          const ids = allocate(before, intent);
          const compiled = compileEditorIntent(before, {
            intent,
            mode: { type: "editing", newIds: ids.newIds },
          }).unwrap();
          const edited = applyDocumentOps(before, compiled.ops).unwrap();
          expect(contractViolation(edited.document)).toBeUndefined();
          const afterById = new Map(
            storyParagraphs(edited.document.package.document).map(
              ({ paragraph: item }) => [item.paraId, item],
            ),
          );
          for (const { paragraph: item } of storyParagraphs(
            before.package.document,
          ))
            expect(afterById.get(item.paraId)).toBe(item);
          const envelopes = JSON.parse(
            JSON.stringify(edited.inverse.map(toOpEnvelope)),
          );
          let restored = edited.document;
          const redo = [];
          for (const envelope of envelopes) {
            const undone = applyDocumentOpEnvelope(restored, envelope).unwrap();
            restored = undone.document;
            redo.unshift(...undone.inverse);
          }
          expect(restored).toStrictEqual(before);
          const redone = applyDocumentOps(restored, redo).unwrap();
          expect(redone.document).toStrictEqual(edited.document);
          const independent = {
            type: DOCUMENT_OP_TYPES.INSERT_TEXT,
            at: { story: OP_STORIES.MAIN, blockId: "10000002", offset: 0 },
            text: "!",
            runProps: INHERIT_RUN_PROPS,
          } as const;
          const elsewhere = applyDocumentOp(
            edited.document,
            independent,
          ).unwrap();
          expect(
            applyDocumentOps(elsewhere.document, edited.inverse).unwrap()
              .document,
          ).toStrictEqual(
            applyDocumentOp(before, independent).unwrap().document,
          );
          const changed = applyDocumentOp(edited.document, {
            ...independent,
            at: compiled.selection,
          }).unwrap();
          const stale = applyDocumentOps(changed.document, edited.inverse);
          expect(stale.isErr()).toBe(true);
          if (stale.isErr())
            expect(stale.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
          expect(before).toStrictEqual(original);
        },
      ),
      {
        id: "table-side paragraph insertion preserves source owners and exact JSON undo redo",
        numRuns: 40,
      },
    );
  },
  propertyTestTimeout(30_000),
);

for (const pending of [false, true])
  test(`tracked table-side insertion accepts and rejects with source tables pending=${pending}`, () => {
    let before = source("body");
    if (pending)
      before = applyDocumentOp(before, {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "10000022",
        revision: { id: 10, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
        newIds: {
          revision: Array.from({ length: 20 }, (_, index) => index + 11),
        },
      }).unwrap().document;
    const original = structuredClone(before);
    const intent = {
      type: "insertParagraphs",
      story: OP_STORIES.MAIN,
      at: { type: "beforeTable", blockId: "10000022" },
      paragraphs: [
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Inserted" }] },
          ],
        },
      ],
    } as const satisfies EditorIntent;
    const ids = createEditorIntentIdAllocator()(before, intent);
    const compiled = compileEditorIntent(before, {
      intent,
      mode: {
        type: "suggesting",
        revision: { id: 100, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
        newIds: ids.newIds,
      },
    }).unwrap();
    const edited = applyDocumentOps(before, compiled.ops).unwrap();
    const rejected = applyDocumentOp(edited.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: edited.revisions,
      decision: REVISION_DECISIONS.REJECT,
    }).unwrap();
    expect(rejected.document).toStrictEqual(before);
    const accepted = applyDocumentOp(edited.document, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: edited.revisions,
      decision: REVISION_DECISIONS.ACCEPT,
    }).unwrap();
    const direct = compileEditorIntent(before, {
      intent,
      mode: { type: "editing", newIds: ids.newIds },
    }).unwrap();
    expect(accepted.document).toStrictEqual(
      applyDocumentOps(before, direct.ops).unwrap().document,
    );
    expect(before).toStrictEqual(original);
  });

for (const kind of ["mark", "properties"] as const)
  test(`empty tail replacement restores a producer-owned prior paragraph ${kind} through JSON`, () => {
    const initial = normalizeForOps({
      package: {
        document: {
          content: [paragraph({ paraId: "10000001", text: "Kept" })],
        },
      },
    });
    const tail = paragraph({ paraId: "10000002", text: "Tail" });
    const stamp = { id: 7, author: "Reviewer", date: "2026-01-02T03:04:05Z" };
    let before: Document;
    if (kind === "mark")
      before = applyDocumentOp(initial, {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId: "10000001" },
        blocks: [tail],
        revision: stamp,
        newIds: {
          revision: Array.from({ length: 20 }, (_, index) => index + 8),
        },
      }).unwrap().document;
    else {
      const formatted = applyDocumentOp(initial, {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: OP_STORIES.MAIN,
        blockId: "10000001",
        patch: { alignment: "center" },
        revision: stamp,
      }).unwrap();
      before = applyDocumentOp(formatted.document, {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId: "10000001" },
        blocks: [tail],
      }).unwrap().document;
    }
    const removed = before.package.document.content.at(-1);
    if (removed?.type !== "paragraph")
      panic("The actual producer needs its tail paragraph.");
    const edit = applyDocumentOp(before, {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: OP_STORIES.MAIN,
      expected: [removed],
      blocks: [],
    }).unwrap();
    let restored = edit.document;
    const redo = [];
    for (const envelope of JSON.parse(
      JSON.stringify(edit.inverse.map(toOpEnvelope)),
    )) {
      const undo = applyDocumentOpEnvelope(restored, envelope).unwrap();
      restored = undo.document;
      redo.unshift(...undo.inverse);
    }
    expect(restored).toStrictEqual(before);
    expect(applyDocumentOps(restored, redo).unwrap().document).toStrictEqual(
      edit.document,
    );
  });

for (const formatting of [
  { type: "empty", value: {} },
  { type: "undefined", value: undefined },
  { type: "authored", value: { alignment: "center" } },
] as const)
  for (const boundary of ["paragraph", "table"] as const)
    test(`empty replacement preserves owned formatting ${formatting.type} beside ${boundary} or refuses atomically`, () => {
      const removed = {
        ...paragraph({ paraId: "10000002", text: "Removed" }),
        formatting: formatting.value,
        textId: undefined,
      } satisfies Paragraph;
      const before = normalizeForOps({
        package: {
          document: {
            content: [
              boundary === "paragraph"
                ? paragraph({ paraId: "10000001", text: "Kept" })
                : table("10000011"),
              removed,
            ],
          },
        },
      });
      const original = structuredClone(before);
      expect(contractViolation(before)).toBeUndefined();
      const result = applyDocumentOp(before, {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: OP_STORIES.MAIN,
        expected: [removed],
        blocks: [],
      });
      if (formatting.type !== "authored") {
        expect(result.isErr()).toBe(true);
        if (result.isOk())
          panic("Owned empty formatting must refuse before removal.");
        expect(result.error.reason).toBe(
          DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
        );
        expect(before).toStrictEqual(original);
        return;
      }
      const edit = result.unwrap();
      let restored = edit.document;
      const redo = [];
      for (const envelope of JSON.parse(
        JSON.stringify(edit.inverse.map(toOpEnvelope)),
      )) {
        const undo = applyDocumentOpEnvelope(restored, envelope).unwrap();
        restored = undo.document;
        redo.unshift(...undo.inverse);
      }
      expect(restored).toStrictEqual(before);
      expect(applyDocumentOps(restored, redo).unwrap().document).toStrictEqual(
        edit.document,
      );
      expect(before).toStrictEqual(original);
    });

for (const container of ["body", "sdt", "cell"] as const)
  test(`empty replacement refuses erasing the sole ${container} block list`, () => {
    const sole = paragraph({ paraId: "10000001", text: "Kept" });
    const before = normalizeForOps({
      package: { document: { content: wrap([sole], container) } },
    });
    const original = structuredClone(before);
    const refusal = applyDocumentOp(before, {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: OP_STORIES.MAIN,
      expected: [sole],
      blocks: [],
    });
    expect(refusal.isErr()).toBe(true);
    if (refusal.isErr())
      expect(refusal.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      );
    expect(before).toStrictEqual(original);
  });

test("empty replacement refuses leaving a table as its cell terminator", () => {
  const tail = paragraph({ paraId: "10000001", text: "Terminator" });
  const before = normalizeForOps({
    package: { document: { content: wrap([table("10000011"), tail], "cell") } },
  });
  const original = structuredClone(before);
  const refusal = applyDocumentOp(before, {
    type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
    story: OP_STORIES.MAIN,
    expected: [tail],
    blocks: [],
  });
  expect(refusal.isErr()).toBe(true);
  if (refusal.isErr())
    expect(refusal.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.CONTAINER_FINAL_MARK,
    );
  expect(before).toStrictEqual(original);
});

for (const boundary of [
  "nested-structure",
  "review-structure",
  "structure",
  "terminal",
  "range",
  "no-carrier",
] as const)
  test(`new tracked table-side insertion refuses ${boundary} before changing source`, () => {
    let before = structuredClone(source("body"));
    const target = before.package.document.content.at(2);
    if (target?.type !== "table") panic("The source target table is required.");
    const cell = target.rows.at(0)?.cells.at(0);
    if (cell?.content.at(0)?.type !== "paragraph")
      panic("The typed source table needs its receiver.");
    const control = {
      type: "blockSdt",
      properties: { sdtType: "richText", id: 88 },
      content: [paragraph({ paraId: "10000030", text: "Nested" })],
    } satisfies BlockContent;
    if (boundary === "structure" || boundary === "review-structure")
      cell.content.unshift(control);
    if (boundary === "nested-structure")
      cell.content.unshift({
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [
                  control,
                  paragraph({ paraId: "10000031", text: "End" }),
                ],
              },
            ],
          },
        ],
      });
    if (boundary === "terminal") before.package.document.content.pop();
    if (boundary === "range")
      before.package.document.content.splice(3, 0, {
        type: "commentRangeStart",
        id: 55,
      });
    if (boundary === "no-carrier")
      before.package.document.content = before.package.document.content.filter(
        (block) => block.type === "table",
      );
    if (boundary === "review-structure")
      before = applyDocumentOp(before, {
        type: DOCUMENT_OP_TYPES.DELETE_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "10000022",
        revision: { id: 10, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
        newIds: {
          revision: Array.from({ length: 20 }, (_, index) => index + 11),
        },
      }).unwrap().document;
    const original = structuredClone(before);
    const at = {
      type:
        boundary === "terminal" || boundary === "range"
          ? "afterTable"
          : "beforeTable",
      blockId: "10000022",
    } as const;
    const intent = {
      type: "insertParagraphs",
      story: OP_STORIES.MAIN,
      at,
      paragraphs: [
        {
          type: "paragraph",
          content: [
            { type: "run", content: [{ type: "text", text: "Inserted" }] },
          ],
        },
      ],
    } as const satisfies EditorIntent;
    const ids = createEditorIntentIdAllocator()(before, intent);
    const compiled = compileEditorIntent(before, {
      intent,
      mode: {
        type: "suggesting",
        revision: { id: 100, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
        newIds: ids.newIds,
      },
    }).unwrap();
    const rejected = applyDocumentOps(before, compiled.ops);
    expect(rejected.isErr()).toBe(true);
    if (rejected.isErr())
      expect(rejected.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
      );
    expect(before).toStrictEqual(original);
    const direct = compileEditorIntent(before, {
      intent,
      mode: { type: "editing", newIds: ids.newIds },
    }).unwrap();
    const applied = applyDocumentOps(before, direct.ops).unwrap();
    expect(
      applyDocumentOps(applied.document, applied.inverse).unwrap().document,
    ).toStrictEqual(before);
  });

test("a pending-table insertion keeps a nonempty rejected mark refused until its content is rejected", () => {
  const pending = applyDocumentOp(source("body"), {
    type: DOCUMENT_OP_TYPES.DELETE_TABLE,
    story: OP_STORIES.MAIN,
    blockId: "10000022",
    revision: { id: 10, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
    newIds: { revision: Array.from({ length: 20 }, (_, index) => index + 11) },
  }).unwrap().document;
  const intent = {
    type: "insertParagraphs",
    story: OP_STORIES.MAIN,
    at: { type: "beforeTable", blockId: "10000022" },
    paragraphs: [
      {
        type: "paragraph",
        content: [
          { type: "run", content: [{ type: "text", text: "Inserted" }] },
        ],
      },
    ],
  } as const satisfies EditorIntent;
  const ids = createEditorIntentIdAllocator()(pending, intent);
  const compiled = compileEditorIntent(pending, {
    intent,
    mode: {
      type: "suggesting",
      revision: { id: 100, author: "Reviewer", date: "2026-01-02T03:04:05Z" },
      newIds: ids.newIds,
    },
  }).unwrap();
  const edit = applyDocumentOps(pending, compiled.ops).unwrap();
  const inserted = storyParagraphs(edit.document.package.document).find(
    ({ paragraph: item }) => item.paraId === compiled.selection.blockId,
  )?.paragraph;
  const mark = inserted?.pPrMark?.info.id;
  if (mark === undefined)
    panic("The actual tracked producer needs its mark ID.");
  const refused = applyDocumentOp(edit.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [mark],
    decision: REVISION_DECISIONS.REJECT,
  });
  expect(refused.isErr()).toBe(true);
  if (refused.isErr())
    expect(refused.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
  const content = edit.revisions.filter((id) => id !== mark);
  expect(content.length).toBeGreaterThan(0);
  const emptied = applyDocumentOp(edit.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: content,
    decision: REVISION_DECISIONS.REJECT,
  }).unwrap();
  const removed = applyDocumentOp(emptied.document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [mark],
    decision: REVISION_DECISIONS.REJECT,
  }).unwrap();
  expect(removed.document).toStrictEqual(pending);
});
