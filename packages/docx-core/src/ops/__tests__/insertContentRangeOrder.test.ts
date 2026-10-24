import { expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type InsertContentOp, OP_STORIES } from "../types";

type RangeMarkerType =
  | "bookmarkStart"
  | "bookmarkEnd"
  | "moveFromRangeStart"
  | "moveFromRangeEnd"
  | "moveToRangeStart"
  | "moveToRangeEnd";

const paragraph = (paraId: string, content: ParagraphContent[]): Paragraph => ({
  type: "paragraph",
  paraId,
  content,
});

const documentOf = (
  ...content: Document["package"]["document"]["content"]
): Document => ({ package: { document: { content } } });

const textRun = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});

const at = (blockId: string) => ({
  story: OP_STORIES.MAIN,
  blockId,
  offset: 0,
});

const marker = (type: RangeMarkerType, id: number): ParagraphContent => {
  switch (type) {
    case "bookmarkStart":
      return { type, id, name: `_Ref${id}` };
    case "bookmarkEnd":
      return { type, id };
    case "moveFromRangeStart":
      return { type, id, name: `_Move${id}`, author: "Reviewer" };
    case "moveFromRangeEnd":
      return { type, id };
    case "moveToRangeStart":
      return { type, id, name: `_Move${id}`, author: "Reviewer" };
    case "moveToRangeEnd":
      return { type, id };
    default: {
      const unreachable: never = type;
      return unreachable;
    }
  }
};

const inverse = (document: Document, op: InsertContentOp): Document => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  const undone = applyDocumentOps(
    applied.value.document,
    applied.value.inverse,
  );
  if (undone.isErr()) throw undone.error;
  return undone.value.document;
};

const cases = [
  ["bookmarkStart", "bookmarkEnd"],
  ["bookmarkEnd", "bookmarkStart"],
  ["moveFromRangeStart", "moveFromRangeEnd"],
  ["moveFromRangeEnd", "moveFromRangeStart"],
  ["moveToRangeStart", "moveToRangeEnd"],
  ["moveToRangeEnd", "moveToRangeStart"],
] as const;

for (const [incomingType, existingType] of cases) {
  test(`${incomingType} refuses an insertion that completes an inverted range`, () => {
    const incoming = marker(incomingType, 7);
    const existing = marker(existingType, 7);
    const incomingIsEnd =
      incomingType === "bookmarkEnd" ||
      incomingType === "moveFromRangeEnd" ||
      incomingType === "moveToRangeEnd";
    const target = incomingIsEnd ? "00000001" : "00000002";
    const document = incomingIsEnd
      ? documentOf(
          paragraph("00000001", [textRun("a")]),
          paragraph("00000002", [existing]),
        )
      : documentOf(
          paragraph("00000001", [existing]),
          paragraph("00000002", [textRun("a")]),
        );
    const before = structuredClone(document);
    const result = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: at(target),
      slice: { content: [incoming], openStart: 0, openEnd: 0 },
    });

    expect(result.isErr()).toBe(true);
    if (result.isErr())
      expect(result.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
      );
    expect(document).toStrictEqual(before);
  });

  for (const placement of ["paragraph", "table"] as const) {
    test(`deleting ${incomingType} from an imported inverted range in a ${placement} undoes exactly`, () => {
      const incomingIsEnd =
        incomingType === "bookmarkEnd" ||
        incomingType === "moveFromRangeEnd" ||
        incomingType === "moveToRangeEnd";
      const ending = paragraph("00000001", [
        marker(incomingIsEnd ? incomingType : existingType, 7),
      ]);
      const starting = paragraph("00000002", [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText" },
          content: incomingType.startsWith("bookmark")
            ? [
                {
                  type: "hyperlink",
                  href: "https://example.org",
                  children: [{ type: "bookmarkStart", id: 7, name: "_Ref7" }],
                },
              ]
            : [marker(incomingIsEnd ? existingType : incomingType, 7)],
        },
      ]);
      const document =
        placement === "paragraph"
          ? documentOf(ending, starting)
          : documentOf(
              {
                type: "table",
                rows: [
                  {
                    type: "tableRow",
                    cells: [{ type: "tableCell", content: [ending] }],
                  },
                  {
                    type: "tableRow",
                    cells: [
                      {
                        type: "tableCell",
                        content: [paragraph("00000003", [])],
                      },
                    ],
                  },
                ],
              },
              starting,
            );
      const original = structuredClone(document);
      const blockId = incomingIsEnd ? "00000001" : "00000002";
      const deleted = applyDocumentOp(document, {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: { ...at(blockId), zeroWidthBefore: 0 },
        to: { ...at(blockId), zeroWidthBefore: 1 },
      });
      expect(deleted.isOk()).toBe(true);
      if (deleted.isErr()) return;
      expect(deleted.value.inverse.at(0)?.type).toBe(
        DOCUMENT_OP_TYPES.REPLACE_INLINE,
      );
      const restored = applyDocumentOps(
        deleted.value.document,
        JSON.parse(JSON.stringify(deleted.value.inverse)),
      );
      expect(restored.isOk()).toBe(true);
      if (restored.isOk())
        expect(restored.value.document).toStrictEqual(original);
      expect(document).toStrictEqual(original);

      const changed = applyDocumentOp(deleted.value.document, {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at(blockId),
        text: "changed",
        runProps: {},
      });
      expect(changed.isOk()).toBe(true);
      if (changed.isErr()) return;
      const stale = applyDocumentOps(
        changed.value.document,
        deleted.value.inverse,
      );
      expect(stale.isErr()).toBe(true);
      if (stale.isErr())
        expect(stale.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
    });
  }
}

test("ordinary deletions retain compact slice inverses", () => {
  for (const content of [
    [marker("bookmarkEnd", 7)],
    [marker("bookmarkStart", 7), marker("bookmarkEnd", 7)],
    [
      marker("bookmarkEnd", 7),
      marker("bookmarkEnd", 7),
      marker("bookmarkStart", 7),
    ],
  ]) {
    const document = documentOf(paragraph("00000001", content));
    const deleted = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: { ...at("00000001"), zeroWidthBefore: 0 },
      to: { ...at("00000001"), zeroWidthBefore: 1 },
    });
    expect(deleted.isOk()).toBe(true);
    if (deleted.isErr()) continue;
    expect(deleted.value.inverse.at(0)?.type).toBe(
      DOCUMENT_OP_TYPES.INSERT_CONTENT,
    );
    const restored = applyDocumentOps(
      deleted.value.document,
      deleted.value.inverse,
    );
    expect(restored.isOk()).toBe(true);
    if (restored.isOk())
      expect(restored.value.document).toStrictEqual(document);
  }
});

test("a block-level range marker participates in insertion order", () => {
  const document = documentOf(
    { type: "bookmarkEnd", id: 8 },
    paragraph("00000001", [textRun("a")]),
  );
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at("00000001"),
    slice: {
      content: [{ type: "bookmarkStart", id: 8, name: "_Ref8" }],
      openStart: 0,
      openEnd: 0,
    },
  });

  expect(result.isErr()).toBe(true);
  if (result.isErr())
    expect(result.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
});

test("positioned table and nested inline markers participate in story order", () => {
  const document = documentOf(
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [{ type: "tableCell", content: [paragraph("00000001", [])] }],
        },
      ],
      bookmarks: [{ index: 1, marker: { type: "bookmarkEnd", id: 9 } }],
    },
    paragraph("00000002", [textRun("a")]),
  );
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at("00000002"),
    slice: {
      content: [{ type: "bookmarkStart", id: 9, name: "_Ref9" }],
      openStart: 0,
      openEnd: 0,
    },
  });

  expect(result.isErr()).toBe(true);
});

test("a positioned bookmark beyond the table tail is ignored", () => {
  const document = documentOf(
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [{ type: "tableCell", content: [paragraph("00000001", [])] }],
        },
      ],
      bookmarks: [{ index: 2, marker: { type: "bookmarkEnd", id: 15 } }],
    },
    paragraph("00000002", [textRun("a")]),
  );
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at("00000002"),
    slice: {
      content: [{ type: "bookmarkStart", id: 15, name: "_Ref15" }],
      openStart: 0,
      openEnd: 0,
    },
  });

  expect(result.isOk()).toBe(true);
});

test("a positioned row bookmark at the cell-list tail participates in order", () => {
  const document = documentOf(
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [{ type: "tableCell", content: [paragraph("00000001", [])] }],
          bookmarks: [
            { index: 1, marker: { type: "moveFromRangeEnd", id: 16 } },
          ],
        },
      ],
    },
    paragraph("00000002", [
      { type: "run", content: [{ type: "text", text: "a" }] },
    ]),
  );
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at("00000002"),
    slice: {
      content: [
        {
          type: "moveFromRangeStart",
          id: 16,
          name: "_Move16",
          author: "Reviewer",
        },
      ],
      openStart: 0,
      openEnd: 0,
    },
  });

  expect(result.isErr()).toBe(true);
});

test("an inverted marker nested in a hyperlink refuses the completing insertion", () => {
  const document = documentOf(
    paragraph("00000001", [
      {
        type: "hyperlink",
        href: "https://example.org",
        children: [{ type: "bookmarkEnd", id: 14 }],
      },
    ]),
    paragraph("00000002", [textRun("a")]),
  );
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at("00000002"),
    slice: {
      content: [{ type: "bookmarkStart", id: 14, name: "_Ref14" }],
      openStart: 0,
      openEnd: 0,
    },
  });

  expect(result.isErr()).toBe(true);
});

test("nested markers are ordered and valid inserted ranges undo exactly", () => {
  const document = documentOf(paragraph("00000001", [textRun("ab")]));
  const op: InsertContentOp = {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: { ...at("00000001"), offset: 1 },
    slice: {
      content: [
        {
          type: "hyperlink",
          href: "https://example.org",
          children: [
            { type: "bookmarkStart", id: 10, name: "_Ref10" },
            textRun("x"),
            { type: "bookmarkEnd", id: 10 },
          ],
        },
      ],
      openStart: 0,
      openEnd: 0,
    },
  };

  expect(inverse(document, op)).toStrictEqual(document);
});

test("unpaired and ambiguous markers remain insertable", () => {
  const unpaired = documentOf(paragraph("00000001", []));
  expect(
    applyDocumentOp(unpaired, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: at("00000001"),
      slice: {
        content: [{ type: "bookmarkStart", id: 11, name: "_Ref11" }],
        openStart: 0,
        openEnd: 0,
      },
    }).isOk(),
  ).toBe(true);

  const ambiguous = documentOf(
    paragraph("00000001", [
      { type: "bookmarkEnd", id: 12 },
      { type: "bookmarkEnd", id: 12 },
    ]),
    paragraph("00000002", [{ type: "bookmarkStart", id: 12, name: "_Ref12" }]),
  );
  expect(
    applyDocumentOp(ambiguous, {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at: at("00000002"),
      slice: {
        content: [{ type: "bookmarkStart", id: 12, name: "_Ref12" }],
        openStart: 0,
        openEnd: 0,
      },
    }).isOk(),
  ).toBe(true);
});

test("an unrelated insertion leaves an imported inverted range untouched", () => {
  const document = documentOf(
    paragraph("00000001", [{ type: "bookmarkEnd", id: 13 }]),
    paragraph("00000002", [{ type: "bookmarkStart", id: 13, name: "_Ref13" }]),
  );
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at("00000001"),
    slice: {
      content: [textRun("x")],
      openStart: 0,
      openEnd: 0,
    },
  });

  expect(result.isOk()).toBe(true);
});

test("the third tracked operation refuses the exact inverted-range completion", () => {
  const document = documentOf(
    paragraph("00000001", []),
    paragraph("00000002", []),
  );
  const original = structuredClone(document);
  const first = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at("00000001"),
    slice: {
      content: [{ type: "bookmarkEnd", id: 18 }],
      openStart: 0,
      openEnd: 0,
    },
    revision: { id: 101, author: "Reviewer" },
  });
  expect(first.isOk()).toBe(true);
  if (first.isErr()) return;

  const second = applyDocumentOp(first.value.document, {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: at("00000001"),
    text: "x",
    runProps: {},
    revision: { id: 102, author: "Reviewer" },
  });
  expect(second.isOk()).toBe(true);
  if (second.isErr()) return;

  const beforeThird = structuredClone(second.value.document);
  const third = applyDocumentOp(second.value.document, {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at("00000002"),
    slice: {
      content: [{ type: "bookmarkStart", id: 18, name: "_Ref18" }],
      openStart: 0,
      openEnd: 0,
    },
    revision: { id: 103, author: "Reviewer" },
  });

  expect(third.isErr()).toBe(true);
  if (third.isErr())
    expect(third.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
  expect(document).toStrictEqual(original);
  expect(second.value.document).toStrictEqual(beforeThird);
});
