import { expect, spyOn, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import * as blocks from "../blocks";
import { formatParaId } from "../idLease";
import { paragraphLogicalText } from "../offsets";
import {
  planTrackedReplace,
  type PlanTrackedReplaceOptions,
} from "../rangeReplacement";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  INHERIT_RUN_PROPS,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const paragraph = (id: number, content: Paragraph["content"]): Paragraph => ({
  type: "paragraph",
  paraId: formatParaId(id),
  content,
});

const operation = ({
  id,
  from,
  to,
}: {
  id: string;
  from: number;
  to: number;
}): PlanTrackedReplaceOptions => ({
  from: { story: OP_STORIES.MAIN, blockId: id.toLowerCase(), offset: from },
  to: { story: OP_STORIES.MAIN, blockId: id, offset: to },
  revision: { id: 1000, author: "Author", date: "2026-01-11T00:00:00.000Z" },
  newIds: { revision: [1001, 1002, 1003, 1004] },
  replacement: {
    paragraphs: [],
    tail: { content: [run("NEW")], openStart: 0, openEnd: 0 },
  },
});

test.each(["body", "cell"] as const)(
  "same-paragraph replacement uses the index on cold and edited %s bodies",
  (container) => {
    const count = 96;
    const anchor = paragraph(175, [run("abc"), run("def")]);
    const prefix = Array.from({ length: count }, (_, index) =>
      paragraph(index + 1, [run(String(index))]),
    );
    const enclosed: BlockContent =
      container === "body"
        ? anchor
        : {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [anchor, paragraph(176, [run("next")])],
                  },
                ],
              },
            ],
          };
    const baseline: Document = {
      package: { document: { content: [...prefix, enclosed] } },
    };
    const id = formatParaId(175);
    const walks = spyOn(blocks, "storyParagraphs");
    try {
      const cold = planTrackedReplace(
        baseline,
        operation({ id, from: 3, to: 3 }),
      ).unwrap();
      expect(cold.map(({ type }) => type)).toEqual([
        DOCUMENT_OP_TYPES.INSERT_CONTENT,
      ]);
      // A cold source still runs the native contract's two full walks.
      expect(
        walks.mock.calls.filter(([body]) => body.content.length >= count),
      ).toHaveLength(2);

      const changed = applyDocumentOp(baseline, {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: { story: OP_STORIES.MAIN, blockId: formatParaId(1), offset: 0 },
        text: "intervening-",
        runProps: INHERIT_RUN_PROPS,
      }).unwrap().document;
      walks.mockClear();

      const replacement = planTrackedReplace(
        changed,
        operation({ id, from: 1, to: 5 }),
      ).unwrap();
      expect(replacement.map(({ type }) => type)).toEqual([
        DOCUMENT_OP_TYPES.DELETE_RANGE,
        DOCUMENT_OP_TYPES.INSERT_CONTENT,
      ]);
      expect(
        walks.mock.calls.every(([body]) => body.content.length < count),
      ).toBe(true);
      expect(
        walks.mock.calls.reduce(
          (slots, [body]) => slots + body.content.length,
          0,
        ),
      ).toBeLessThan(count);

      const tracked = applyDocumentOps(changed, replacement).unwrap();
      const accepted = applyDocumentOp(tracked.document, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        decision: REVISION_DECISIONS.ACCEPT,
        revisionIds: tracked.revisions,
      }).unwrap().document;
      const acceptedAnchor = blocks
        .storyParagraphs(accepted.package.document)
        .find(
          ({ paragraph: acceptedParagraph }) => acceptedParagraph.paraId === id,
        );
      expect(
        acceptedAnchor && paragraphLogicalText(acceptedAnchor.paragraph),
      ).toBe("aNEWf");
      expect(
        applyDocumentOps(tracked.document, tracked.inverse).unwrap().document,
      ).toEqual(changed);
    } finally {
      walks.mockRestore();
    }
  },
);

test("same-paragraph endpoint and offset refusals preserve their order", () => {
  const id = formatParaId(17);
  const baseline: Document = {
    package: { document: { content: [paragraph(17, [run("abc")])] } },
  };
  const missing = planTrackedReplace(
    baseline,
    operation({ id: "000000AA", from: 9, to: 9 }),
  );
  const outOfRange = planTrackedReplace(
    baseline,
    operation({ id, from: 9, to: 9 }),
  );
  expect(missing.isErr() && missing.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
  );
  expect(outOfRange.isErr() && outOfRange.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
  );
});
