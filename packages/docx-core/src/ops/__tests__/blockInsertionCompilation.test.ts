import { expect, spyOn, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { normalizeForOps } from "../contract";
import {
  beginTrackedInsertionCompilation,
  compileEditorIntent,
  type EditorIntent,
  type EditorIntentMode,
} from "../documentOps";
import * as idLease from "../idLease";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const paragraph = (id: string): Paragraph => ({
  type: "paragraph",
  paraId: id,
  content: [{ type: "run", content: [{ type: "text", text: id }] }],
});

test("one monotone paragraph allocation produces the exact per-member native operations", () => {
  const count = 64;
  const reservedBlockIds = new Set(["00000002", "0000000a", "0000000B"]);
  const baseline: Document = {
    package: {
      document: { content: [paragraph("00000001"), paragraph("00000004")] },
    },
  };
  const members = Array.from({ length: count }, (_, index) => ({
    intent: {
      type: "insertParagraphs",
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000001" },
      paragraphs: [paragraph("FFFFFF00")],
    } satisfies EditorIntent,
    mode: {
      type: "suggesting",
      reservedBlockIds,
      revision: {
        id: 1000 + index * 5,
        author: "Author",
        date: "2026-01-11T00:00:00.000Z",
      },
      newIds: {
        revision: [1003 + index * 5, 1001 + index * 5],
        control: [40, 6],
      },
    } satisfies EditorIntentMode,
  }));

  let document = baseline;
  const conventional = members.map(({ intent, mode }) => {
    const [op] = compileEditorIntent(document, { intent, mode }).unwrap().ops;
    if (op === undefined) throw new Error("Missing sequential operation");
    document = applyDocumentOp(document, op).unwrap().document;
    return op;
  });

  const allocations = spyOn(idLease, "takeParagraphIds");
  try {
    const group = beginTrackedInsertionCompilation({
      document: baseline,
      count,
      reservedBlockIds,
    });
    if (group.type !== "ready") throw new Error("Unexpected compile fallback");
    const first = members.at(0);
    if (first === undefined) throw new Error("Missing test member");
    expect(group.compile({ ...first, index: 1 }).unwrap().type).toBe(
      "unsupported",
    );
    expect(
      group
        .compile({
          ...first,
          index: 0,
          mode: { ...first.mode, reservedBlockIds: new Set() },
        })
        .unwrap().type,
    ).toBe("unsupported");
    const batch = members.map((member, index) => {
      const result = group.compile({ index, ...member }).unwrap();
      if (result.type !== "operation")
        throw new Error("Unexpected member fallback");
      expect(group.compile({ index, ...member }).unwrap().type).toBe(
        "unsupported",
      );
      return result.op;
    });
    expect(batch).toEqual(conventional);
    expect(
      batch.every(({ type }) => type === DOCUMENT_OP_TYPES.INSERT_BLOCKS),
    ).toBe(true);
    expect(group.compile({ ...first, index: count }).unwrap().type).toBe(
      "unsupported",
    );
    expect(
      beginTrackedInsertionCompilation({ document: baseline, count: -1 }).type,
    ).toBe("unsupported");
    expect(allocations).toHaveBeenCalledTimes(1);
  } finally {
    allocations.mockRestore();
  }
});

test("a main-story address routed into a text box stays on the ordinary path", () => {
  const baseline = normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "run",
                content: [
                  {
                    type: "shape",
                    shape: {
                      type: "shape",
                      shapeType: "textBox",
                      id: "1",
                      size: { width: 2_000_000, height: 800_000 },
                      textBody: { content: [paragraph("00000010")] },
                    },
                  },
                ],
              },
            ],
          },
        ],
      },
    },
  });
  const intent = {
    type: "insertParagraphs",
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId: "00000010" },
    paragraphs: [paragraph("00000040")],
  } satisfies EditorIntent;
  const mode = {
    type: "suggesting",
    revision: { id: 601, author: "Author", date: "2026-01-11T00:00:00.000Z" },
    newIds: { revision: [602] },
  } satisfies EditorIntentMode;
  const group = beginTrackedInsertionCompilation({
    document: baseline,
    count: 1,
  });
  if (group.type !== "ready") throw new Error("Missing compiler");
  expect(group.compile({ index: 0, intent, mode }).unwrap().type).toBe(
    "unsupported",
  );
  const [ordinary] = compileEditorIntent(baseline, { intent, mode }).unwrap()
    .ops;
  if (ordinary === undefined) throw new Error("Missing ordinary operation");
  expect(applyDocumentOp(baseline, ordinary).isOk()).toBe(true);
  expect(
    group
      .compile({
        index: 0,
        intent: { ...intent, at: { type: "after", blockId: "00000001" } },
        mode,
      })
      .unwrap().type,
  ).toBe("operation");
});
