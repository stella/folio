import { expect, spyOn, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import * as blocks from "../blocks";
import { validateOpsDocument } from "../contract";
import { formatParaId } from "../idLease";
import { paragraphLogicalText } from "../offsets";
import { planTrackedDeletion, selectedParagraphRuns } from "../plan";
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
  sourceContainerPolicy: "join",
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

test.each([64, 256])(
  "tracked deletion locates adjacent paragraphs in a %i-paragraph story",
  (width) => {
    const document: Document = {
      package: {
        document: {
          content: Array.from({ length: width }, (_, i) =>
            paragraph(i + 1, [run("text")]),
          ),
        },
      },
    };
    validateOpsDocument(document).unwrap();
    const options = {
      from: {
        story: OP_STORIES.MAIN,
        blockId: formatParaId(width - 2),
        offset: 2,
      },
      to: { story: OP_STORIES.MAIN, blockId: formatParaId(width), offset: 1 },
      revision: { id: 1000, author: "Author", date: "2026-01-11T00:00:00Z" },
      newIds: { revision: [1001, 1002, 1003, 1004, 1005] },
    };
    const walks = spyOn(blocks, "storyParagraphs");
    try {
      const runs = selectedParagraphRuns(
        document,
        options.from,
        options.to,
      ).unwrap();
      expect(runs.map((group) => group.map(({ index }) => index))).toEqual([
        [width - 3, width - 2, width - 1],
      ]);
      expect(
        walks.mock.calls.filter(([body]) => body.content.length === width),
      ).toHaveLength(0);
      const ops = planTrackedDeletion(document, options).unwrap();
      expect(
        walks.mock.calls.filter(([body]) => body.content.length >= width - 2),
      ).toHaveLength(0);
      const applied = applyDocumentOps(document, ops).unwrap();
      expect(
        applyDocumentOps(applied.document, applied.inverse).unwrap().document,
      ).toEqual(document);
    } finally {
      walks.mockRestore();
    }
  },
);

test("ranges with unnamed endpoints or nested blocks preserve their story-order walk", () => {
  const unnamed: Paragraph = { type: "paragraph", content: [] };
  const inside = paragraph(2, [run("cell")]);
  const last = paragraph(3, [run("last")]);
  const first = paragraph(1, [run("first")]);
  const document: Document = {
    package: {
      document: {
        content: [
          unnamed,
          first,
          {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [{ type: "tableCell", content: [inside] }],
              },
            ],
          },
          last,
        ],
      },
    },
  };
  const at = (blockId: string) => ({
    story: OP_STORIES.MAIN,
    blockId,
    offset: 0,
  });
  const walks = spyOn(blocks, "storyParagraphs");
  try {
    for (const [left, right, expected] of [
      ["", formatParaId(1), [[unnamed, first]]],
      [formatParaId(1), formatParaId(3), [[first], [inside], [last]]],
      [formatParaId(2), formatParaId(3), [[inside], [last]]],
    ] as const) {
      walks.mockClear();
      const groups = selectedParagraphRuns(
        document,
        at(left),
        at(right),
      ).unwrap();
      expect(
        groups.map((group) => group.map((location) => location.paragraph)),
      ).toEqual(expected);
      expect(
        walks.mock.calls.filter(
          ([body]) => body.content === document.package.document.content,
        ),
      ).toHaveLength(1);
    }
    walks.mockClear();
    const reversed = selectedParagraphRuns(
      document,
      at(formatParaId(3)),
      at(formatParaId(1)),
    );
    expect(reversed.isErr() && reversed.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
    );
    expect(
      walks.mock.calls.filter(
        ([body]) => body.content === document.package.document.content,
      ),
    ).toHaveLength(0);
  } finally {
    walks.mockRestore();
  }
});

test.each([32, 128])(
  "tracked paste visits only its seam paragraphs in a %i-paragraph body",
  (width) => {
    const seed = Array.from({ length: width }, (_, index) =>
      paragraph(index + 1, [run("abcdef")]),
    );
    const anchor = formatParaId(width - 2);
    for (const { fromId, toId, offset, fallbackWalks } of [
      { fromId: anchor, toId: anchor, offset: 1, fallbackWalks: 0 },
      { fromId: anchor, toId: anchor, offset: 3, fallbackWalks: 0 },
      {
        fromId: anchor,
        toId: formatParaId(width),
        offset: 3,
        fallbackWalks: 0,
      },
      {
        fromId: anchor,
        toId: formatParaId(width + 1),
        offset: 3,
        fallbackWalks: 1,
      },
      { fromId: "", toId: "", offset: 3, fallbackWalks: 1 },
    ]) {
      const cell: BlockContent = {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [paragraph(width + 1, [run("cell text")])],
              },
            ],
          },
        ],
      };
      const last = paragraph(width + 2, [run("last text")]);
      if (fromId === "") last.paraId = "";
      const document: Document = {
        package: { document: { content: [...seed, cell, last] } },
      };
      validateOpsDocument(document).unwrap();
      const input = {
        ...operation({ id: fromId, from: 1, to: offset }),
        to: { story: OP_STORIES.MAIN, blockId: toId, offset: offset },
        newIds: { revision: Array.from({ length: 32 }, (_, i) => 1001 + i) },
      };
      const walks = spyOn(blocks, "storyParagraphs");
      try {
        const count = (
          policy: PlanTrackedReplaceOptions["sourceContainerPolicy"],
        ) => {
          walks.mockClear();
          const planned = planTrackedReplace(document, {
            ...input,
            sourceContainerPolicy: policy,
          }).unwrap();
          return {
            planned,
            count: walks.mock.calls.filter(
              ([body]) => body.content.length >= width - 3,
            ).length,
          };
        };
        const joined = count("join");
        const separate = count("separate");
        if (fromId === anchor && toId === anchor && offset === 3) {
          expect(
            separate.planned.some(
              (op) => op.type === DOCUMENT_OP_TYPES.REPLACE_INLINE,
            ),
          ).toBe(true);
        }
        expect(separate.count - joined.count).toBe(fallbackWalks);
        const applied = applyDocumentOps(document, separate.planned).unwrap();
        expect(
          applyDocumentOps(applied.document, applied.inverse).unwrap().document,
        ).toEqual(document);
      } finally {
        walks.mockRestore();
      }
    }
  },
);
