import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
} from "../../model/document";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOpEnvelope,
  applyDocumentOps,
} from "../apply";
import { storyParagraphs } from "../blocks";
import { normalizeForOps } from "../contract";
import {
  compileEditorIntent,
  createEditorIntentIdAllocator,
} from "../editorIntent";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type EditorIntent,
  OP_STORIES,
  REVISION_DECISIONS,
  toOpEnvelope,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
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
        cells: [{ type: "tableCell", content: [paragraph(paraId, "Cell")] }],
      },
    ],
  }) satisfies Table;
const CONTAINERS = ["body", "cell", "sdt", "customXml"] as const;
const SIDES = ["beforeTable", "afterTable"] as const;
const INTERVENING_TABLES = [0, 1, 2] as const;
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
          properties: { id: 77, sdtType: "richText" },
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
const source = (content: BlockContent[]): Document =>
  normalizeForOps({ package: { document: { content } } });
const compile = (
  document: Document,
  intent: EditorIntent,
  revision?: number,
) => {
  const ids = createEditorIntentIdAllocator()(document, intent);
  return compileEditorIntent(document, {
    intent,
    mode:
      revision === undefined
        ? { type: "editing", newIds: ids.newIds }
        : {
            type: "suggesting",
            revision: { id: revision, author: "Reviewer" },
            newIds: ids.newIds,
          },
  }).unwrap();
};
const exactUndoRedo = (before: Document, edit: AppliedDocumentOp) => {
  let restored = edit.document;
  const redo: DocumentOp[] = [];
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
};
const addSourceComments = (document: Document) => {
  const root = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: {
      id: 7,
      author: "Reviewer",
      content: [paragraph("20000001", "Source comment")],
    },
    anchor: {
      kind: "range",
      from: { story: OP_STORIES.MAIN, blockId: "10000001", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "10000001", offset: 3 },
    },
  }).unwrap();
  return applyDocumentOp(root.document, {
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: {
      id: 8,
      author: "Reviewer",
      content: [paragraph("20000002", "Reply")],
    },
    anchor: { kind: "reply", parentId: 7 },
  }).unwrap().document;
};
const pendingSource = (container: Container) => {
  const initial = addSourceComments(
    source(
      wrap(
        [
          {
            ...paragraph("10000001", "Source"),
            formatting: { alignment: "center" },
          },
        ],
        container,
      ),
    ),
  );
  const prepared = compile(
    initial,
    {
      type: "tableBlock",
      operation: {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story: OP_STORIES.MAIN,
        blockId: "10000001",
        side: "after",
        table: table("10000011"),
      },
    },
    10,
  );
  const pending = applyDocumentOps(initial, prepared.ops).unwrap();
  const owner = storyParagraphs(pending.document.package.document).find(
    ({ paragraph: item }) => item.paraId === "10000001",
  )?.paragraph;
  expect(owner?.pPrMark?.survivor).toBe("first");
  return { initial, pending, tableId: prepared.selection.blockId };
};

type SourceOwnerCaseOptions = {
  container: Container;
  side: (typeof SIDES)[number];
  intervening: number;
  texts: readonly string[];
};
const assertSourceOwnerRefusal = ({
  container,
  side,
  intervening,
  texts,
}: SourceOwnerCaseOptions) => {
  const { initial, pending, tableId } = pendingSource(container);
  let before = pending.document;
  for (let index = 0; index < intervening; index += 1) {
    const prepared = compile(
      before,
      {
        type: "tableBlock",
        operation: {
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          story: OP_STORIES.MAIN,
          anchor: "table",
          blockId: tableId,
          side: "before",
          table: table("30000011"),
        },
      },
      40 + index * 20,
    );
    before = applyDocumentOps(before, prepared.ops).unwrap().document;
  }
  const original = structuredClone(before);
  const intent = {
    type: "insertParagraphs",
    story: OP_STORIES.MAIN,
    at: { type: side, blockId: tableId },
    paragraphs: texts.map((text) => ({
      type: "paragraph",
      formatting: { alignment: "right" },
      content: paragraph("00000000", text).content,
    })),
  } as const satisfies EditorIntent;
  const refused = applyDocumentOps(before, compile(before, intent, 100).ops);
  expect(refused.isErr()).toBe(true);
  if (refused.isOk())
    panic(
      "The prior first-source owner must prevent tracked table-side insertion.",
    );
  expect(refused.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
  expect(refused.error.message).toContain(
    "earlier pending first-paragraph owner",
  );
  expect(before).toStrictEqual(original);
  const direct = applyDocumentOps(before, compile(before, intent).ops).unwrap();
  exactUndoRedo(before, direct);
  if (intervening === 0) {
    const rejected = applyDocumentOp(before, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: pending.revisions,
      decision: REVISION_DECISIONS.REJECT,
    }).unwrap();
    expect(rejected.document).toStrictEqual(initial);
    exactUndoRedo(before, rejected);
  }
};

const sourceOwnerCases = CONTAINERS.flatMap((container) =>
  SIDES.flatMap((side) =>
    INTERVENING_TABLES.map((intervening) => ({
      container,
      side,
      intervening,
      texts: [""],
    })),
  ),
);
test.each(sourceOwnerCases)(
  "declared source-owner boundary $container/$side/$intervening refuses with exact direct undo",
  (entry) => {
    assertSourceOwnerRefusal(entry);
  },
);

test(
  "new tracked table-side insertions refuse prior source owners across pending tables in their containing list",
  () => {
    assertProperty(
      fc.property(
        fc.constantFrom(...CONTAINERS),
        fc.constantFrom(...SIDES),
        fc.integer({ min: 0, max: 2 }),
        fc.array(fc.constantFrom("", "New", "Úprava", "表"), {
          minLength: 1,
          maxLength: 2,
        }),
        (container, side, intervening, texts) =>
          assertSourceOwnerRefusal({ container, side, intervening, texts }),
      ),
      {
        id: "new tracked table-side insertions refuse prior source owners across pending tables in their containing list",
        numRuns: 24,
      },
    );
  },
  propertyTestTimeout(30_000),
);

for (const side of ["beforeTable", "afterTable"] as const)
  test(`${side}: a prior source owner in another cell does not block local insertion`, () => {
    const initial = source([
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              { type: "tableCell", content: [paragraph("10000001", "Source")] },
              {
                type: "tableCell",
                content: [
                  paragraph("30000001", "Left"),
                  table("30000011"),
                  paragraph("30000002", "Right"),
                ],
              },
            ],
          },
        ],
      },
    ]);
    const pending = compile(
      initial,
      {
        type: "tableBlock",
        operation: {
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          story: OP_STORIES.MAIN,
          blockId: "10000001",
          side: "after",
          table: table("10000011"),
        },
      },
      10,
    );
    const before = applyDocumentOps(initial, pending.ops).unwrap().document;
    const owner = storyParagraphs(before.package.document).find(
      ({ paragraph: item }) => item.paraId === "10000001",
    )?.paragraph;
    expect(owner?.pPrMark?.survivor).toBe("first");
    const intent = {
      type: "insertParagraphs",
      story: OP_STORIES.MAIN,
      at: { type: side, blockId: "30000011" },
      paragraphs: [
        {
          type: "paragraph",
          content: paragraph("00000000", "Inserted").content,
        },
      ],
    } as const satisfies EditorIntent;
    const edited = applyDocumentOps(
      before,
      compile(before, intent, 100).ops,
    ).unwrap();
    exactUndoRedo(before, edited);
    expect(
      applyDocumentOp(edited.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: edited.revisions,
        decision: REVISION_DECISIONS.REJECT,
      }).unwrap().document,
    ).toStrictEqual(before);
  });

for (const side of SIDES)
  test(`${side}: an ancestor-list source owner does not cross into the nested table's list`, () => {
    const initial = source([
      paragraph("10000001", "Source"),
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [
                  paragraph("30000001", "Left"),
                  table("30000011"),
                  paragraph("30000002", "Right"),
                ],
              },
            ],
          },
        ],
      },
    ]);
    const ancestorIntent = {
      type: "insertParagraphs",
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "10000001" },
      paragraphs: [
        { type: "paragraph", content: paragraph("00000000", "Peer").content },
      ],
    } as const satisfies EditorIntent;
    const ancestor = applyDocumentOps(
      initial,
      compile(initial, ancestorIntent, 10).ops,
    ).unwrap();
    const owner = storyParagraphs(ancestor.document.package.document).find(
      ({ paragraph: item }) => item.paraId === "10000001",
    )?.paragraph;
    expect(owner?.pPrMark?.survivor).toBe("first");
    const intent = {
      type: "insertParagraphs",
      story: OP_STORIES.MAIN,
      at: { type: side, blockId: "30000011" },
      paragraphs: [
        {
          type: "paragraph",
          content: paragraph("00000000", "Inserted").content,
        },
      ],
    } as const satisfies EditorIntent;
    const edited = applyDocumentOps(
      ancestor.document,
      compile(ancestor.document, intent, 100).ops,
    ).unwrap();
    exactUndoRedo(ancestor.document, edited);
    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const groups = [edited.revisions, ancestor.revisions];
      const finish = (order: typeof groups) => {
        let document = edited.document;
        for (const revisionIds of order) {
          const next = applyDocumentOp(document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds,
            decision,
          }).unwrap();
          exactUndoRedo(document, next);
          document = next.document;
        }
        return document;
      };
      const first = finish(groups);
      expect(finish(groups.toReversed())).toStrictEqual(first);
      if (decision === REVISION_DECISIONS.REJECT)
        expect(first).toStrictEqual(initial);
    }
  });
