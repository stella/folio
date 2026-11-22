/**
 * Tracked deletions of paragraphs a pending join runs through, on the
 * canonical session. Where the paragraph's own mark goes too, the join runs on
 * past it and the deletion takes its words alone, as the direct deletion does.
 * Where its mark stays, it is the paragraph the join makes: the words joined
 * into it go with it, as deleting it directly once the join is accepted does.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Paragraph } from "../types/document";

type Mark = NonNullable<Paragraph["pPrMark"]>;

const INFO = { author: "Earlier", date: "2026-01-02T03:04:05Z" };

const paragraph = (
  paraId: string,
  text: string,
  pPrMark?: Mark,
): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
  ...(pPrMark === undefined ? {} : { pPrMark }),
});

const GOING_MARKS = {
  "pending deletion": { kind: "del", info: { id: 2, ...INFO } },
  "pending a move away": { kind: "moveFrom", info: { id: 2, ...INFO } },
  "added, then pending deletion": {
    kind: "ins",
    info: { id: 2, ...INFO },
    deletion: { info: { id: 3, ...INFO } },
  },
} as const satisfies Record<string, Mark>;

const bytesOf = (content: BlockContent[]): Promise<ArrayBuffer> =>
  createDocx({ package: { document: { content } } });

type AcceptedOptions = {
  bytes: ArrayBuffer;
  mode: "direct" | "tracked-changes";
  deleted: string;
  /** Whether the join is accepted before the deletion. */
  joinsFirst?: boolean;
};

/** The words once the deletion and every other change are accepted. */
const accepted = async ({
  bytes,
  mode,
  deleted,
  joinsFirst = false,
}: AcceptedOptions): Promise<string> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
    author: "Editor",
    session: "canonical",
  });
  if (joinsFirst) reviewer.acceptAll();
  const { skipped } = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [{ id: "op", type: "deleteBlock", blockId: deleted }],
    },
    { undefinedReferences: "refuse" },
  );
  expect(skipped).toEqual([]);
  reviewer.acceptAll();
  return reviewer
    .getContent()
    .map(({ text }) => text)
    .join("");
};

test.each(
  Object.entries(GOING_MARKS).map(([state, mark]) => ({ state, mark })),
)(
  "a tracked deletion of a paragraph whose mark is $state accepts as the direct one",
  async ({ mark }) => {
    const bytes = await bytesOf([
      paragraph("00000001", "First.", {
        kind: "del",
        info: { id: 1, ...INFO },
      }),
      paragraph("00000002", "Second.", mark),
      paragraph("00000003", "Third."),
    ]);
    const deleted = "00000002";
    const direct = await accepted({ bytes, mode: "direct", deleted });
    expect(direct).toBe("First.Third.");
    expect(await accepted({ bytes, mode: "tracked-changes", deleted })).toBe(
      direct,
    );
  },
);

const JOINED: Record<string, { content: BlockContent[]; deleted: string }> = {
  "a join through a break moved away": {
    content: [
      paragraph("00000001", "First.", {
        kind: "moveFrom",
        info: { id: 1, ...INFO },
      }),
      paragraph("00000002", "Second.", {
        kind: "del",
        info: { id: 2, ...INFO },
      }),
      paragraph("00000003", "Third."),
      paragraph("00000004", "Fourth."),
    ],
    deleted: "00000003",
  },
  "a paragraph with an added mark before a table": {
    content: [
      paragraph("00000001", "First.", {
        kind: "del",
        info: { id: 1, ...INFO },
      }),
      paragraph("00000002", "Second.", {
        kind: "ins",
        info: { id: 2, ...INFO },
      }),
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              { type: "tableCell", content: [paragraph("00000003", "Cell.")] },
            ],
          },
        ],
      },
      paragraph("00000004", ""),
    ],
    deleted: "00000002",
  },
};

test.each(Object.entries(JOINED).map(([shape, joined]) => ({ shape, joined })))(
  "a tracked deletion of $shape takes the words joined into it",
  async ({ joined: { content, deleted } }) => {
    const bytes = await bytesOf(content);
    const direct = await accepted({
      bytes,
      mode: "direct",
      deleted,
      joinsFirst: true,
    });
    expect(direct).not.toContain("First.");
    expect(await accepted({ bytes, mode: "tracked-changes", deleted })).toBe(
      direct,
    );
  },
);
