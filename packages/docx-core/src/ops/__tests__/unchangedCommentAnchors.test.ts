import { expect, spyOn, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import { formatParaId } from "../ids";
import * as routing from "../textBoxRouting";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES } from "../types";

const paragraph = (text: string, paraId?: string): Paragraph => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId }),
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

type FixtureOptions = { width: number; unnamedParagraphs?: number };
const fixture = ({
  width,
  unnamedParagraphs = 0,
}: FixtureOptions): Document => {
  const content = Array.from({ length: width }, (_, index) =>
    paragraph("Clause", formatParaId(0x10000000 + index)),
  );
  if (unnamedParagraphs > 0)
    content.push({
      type: "paragraph",
      paraId: "1000EEEE",
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
                textBody: {
                  content: Array.from(
                    { length: unnamedParagraphs },
                    (_, index) => paragraph(`Unnamed ${index}`),
                  ),
                },
              },
            },
          ],
        },
      ],
    });
  return { package: { document: { content } } };
};

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const add = (document: Document, id = 1) => {
  const from = { story: OP_STORIES.MAIN, blockId: "10000000", offset: 1 };
  return apply(document, {
    type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
    comment: {
      id,
      author: "Reader",
      content: [paragraph("Note", formatParaId(0x1000fff0 + id))],
    },
    anchor: { kind: "range", from, to: { ...from, offset: 3 } },
  });
};

const status = (resolved: boolean): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.SET_COMMENT_RESOLUTION,
  id: 1,
  status: resolved ? "resolved" : "open",
});

test.each([16, 64])("metadata edits compare a width %i story once", (width) => {
  let document = add(fixture({ width })).document;
  const content = document.package.document.content;
  const enumerate = routing.paragraphsThroughTextBoxes;
  let listed = 0;
  const walk = spyOn(routing, "paragraphsThroughTextBoxes").mockImplementation(
    (current, story) => {
      const found = enumerate(current, story);
      listed += found.length;
      return found;
    },
  );
  try {
    for (let iteration = 0; iteration < 6; iteration += 1) {
      const edit = apply(document, status(iteration % 2 === 0));
      expect(edit.touched.modified).toEqual([]);
      expect(edit.document.package.document.content).toBe(content);
      document = edit.document;
      expect(walk).toHaveBeenCalledTimes(1);
    }
    expect({ calls: walk.mock.calls.length, listed }).toEqual({
      calls: 1,
      listed: width,
    });
  } finally {
    walk.mockRestore();
  }
});

test.each(["raw", "normalized"] as const)(
  "unchanged %s story keeps duplicate unnamed touched entries and exact inverse",
  (kind) => {
    const document = fixture({ width: 2, unnamedParagraphs: 2 });
    const original =
      kind === "normalized" ? normalizeForOps(document) : document;
    const created = add(original);
    expect(created.touched.modified).toEqual(["10000000", ""]);
    const creation = created.inverse.at(0);
    expect(creation?.type).toBe(DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE);
    if (creation?.type === DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE) {
      expect(creation.expected.anchors.map(({ blockId }) => blockId)).toEqual([
        "10000000",
      ]);
      expect(creation.state.anchors.map(({ blockId }) => blockId)).toEqual([
        "10000000",
      ]);
    }
    const before = created.document;
    const resolution = apply(before, status(true));
    expect(resolution.touched.modified).toEqual([""]);
    const inverse = resolution.inverse.at(0);
    expect(inverse?.type).toBe(DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE);
    if (inverse?.type === DOCUMENT_OP_TYPES.RESTORE_COMMENT_STATE) {
      expect(inverse.expected.anchors).toEqual([]);
      expect(inverse.state.anchors).toEqual([]);
    }
    const reverted = applyDocumentOps(resolution.document, resolution.inverse);
    if (reverted.isErr()) throw reverted.error;
    assertExactModel(reverted.value.document, before);
    const reopened = apply(resolution.document, status(false));
    expect(reopened.touched.modified).toEqual([""]);
  },
);

test("a changed story still compares both lists after a metadata edit", () => {
  const current = add(fixture({ width: 16 })).document;
  const warmed = apply(current, status(true)).document;
  const walk = spyOn(routing, "paragraphsThroughTextBoxes");
  try {
    const changed = add(warmed, 2);
    expect(changed.document.package.document.content).not.toBe(
      warmed.package.document.content,
    );
    expect(changed.touched.modified).toEqual(["10000000"]);
    expect(walk).toHaveBeenCalledTimes(2);
    expect(walk.mock.calls.at(0)?.at(0)?.package.document.content).toBe(
      changed.document.package.document.content,
    );
    expect(walk.mock.calls.at(1)?.at(0)?.package.document.content).toBe(
      warmed.package.document.content,
    );
  } finally {
    walk.mockRestore();
  }
});

test("a shared block list is read under each document's current story address", () => {
  const first = add(fixture({ width: 2, unnamedParagraphs: 2 })).document;
  const warm = apply(first, status(true));
  expect(warm.touched.modified).toEqual([""]);
  const secondContent = fixture({
    width: 2,
    unnamedParagraphs: 2,
  }).package.document.content;
  for (const [index, block] of secondContent.entries())
    if (block.type === "paragraph")
      block.paraId = formatParaId(0x20000000 + index);
  const second = {
    ...first,
    package: {
      ...first.package,
      document: { ...first.package.document, content: secondContent },
      headers: new Map([
        [
          "rIdHeader",
          {
            type: "header" as const,
            content: first.package.document.content,
          },
        ],
      ]),
    },
  } satisfies Document;
  const walk = spyOn(routing, "paragraphsThroughTextBoxes");
  try {
    const resolution = apply(second, status(true));
    expect(resolution.touched.modified).toEqual(["", ""]);
    expect(walk).toHaveBeenCalledTimes(1);
    const reopened = apply(resolution.document, status(false));
    expect(reopened.touched.modified).toEqual(["", ""]);
    expect(apply(warm.document, status(false)).touched.modified).toEqual([""]);
    expect(walk).toHaveBeenCalledTimes(1);
  } finally {
    walk.mockRestore();
  }
});

test("a story with many unnamed differences does not retain a large result", () => {
  const document = add(fixture({ width: 2, unnamedParagraphs: 16 })).document;
  const walk = spyOn(routing, "paragraphsThroughTextBoxes");
  try {
    const closed = apply(document, status(true));
    expect(closed.touched.modified).toEqual([""]);
    const reopened = apply(closed.document, status(false));
    expect(reopened.touched.modified).toEqual([""]);
    expect(walk).toHaveBeenCalledTimes(2);
  } finally {
    walk.mockRestore();
  }
});
