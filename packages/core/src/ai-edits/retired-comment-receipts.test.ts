import { expect, test } from "bun:test";

import {
  paragraphsDocx,
  reviewComment,
  textRun,
} from "../__tests__/operationBatchDocuments";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { Document, ParagraphContent, TableCell } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const anchor = (id: number, text: string): ParagraphContent[] => [
  { type: "commentRangeStart", id },
  textRun(text),
  { type: "commentRangeEnd", id },
  { type: "commentReference", id },
];

const open = (session: "legacy" | "canonical") =>
  paragraphsDocx(
    [
      [textRun("First "), ...anchor(3, "clause"), textRun(".")],
      anchor(8, "Second clause."),
      [textRun("Third "), ...anchor(12, "clause"), textRun(".")],
      [textRun("End.")],
    ],
    [reviewComment(3), reviewComment(8), reviewComment(12)],
  ).then((buffer) => FolioDocxReviewer.fromBuffer(buffer, { session }));

const edits = [
  { id: "delete-first", type: "deleteBlock", blockId: "10000000" },
  {
    id: "delete-part",
    type: "replaceInBlock",
    blockId: "10000001",
    find: "Second ",
    replace: "",
  },
  { id: "delete-third", type: "deleteBlock", blockId: "10000002" },
  {
    id: "replace-end",
    type: "replaceInBlock",
    blockId: "10000003",
    find: "End.",
    replace: "Done.",
  },
] as const satisfies readonly FolioDocumentOperation[];

const inlineEdits = [
  {
    id: "delete-first",
    type: "replaceInBlock",
    blockId: "10000000",
    find: "clause",
    replace: "",
  },
  edits[1],
  {
    id: "delete-third",
    type: "replaceInBlock",
    blockId: "10000002",
    find: "clause",
    replace: "",
  },
  edits[3],
] as const satisfies readonly FolioDocumentOperation[];

for (const session of ["legacy", "canonical"] as const) {
  test(`${session} attributes only complete anchor losses in preview and commit`, async () => {
    const reviewer = await open(session);
    const before = reviewer.toDocument();
    for (const dryRun of [true, false]) {
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: session === "canonical" ? inlineEdits : edits,
          dryRun,
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      const expected = [
        ["delete-first", [3]],
        ["delete-part", undefined],
        ["delete-third", [12]],
        ["replace-end", undefined],
      ] as const;
      expect(
        new Map(
          result.applied.map(({ id, retiredCommentIds }) => [
            id,
            retiredCommentIds,
          ]),
        ),
      ).toEqual(new Map(expected));
      expect(
        result.receipts.map(({ operationId, retiredCommentIds }) => [
          operationId,
          retiredCommentIds,
        ]),
      ).toEqual(expected);
      if (dryRun) expect(reviewer.toDocument()).toEqual(before);
    }

    const saved = await reviewer.toBuffer();
    const accepted = await FolioDocxReviewer.fromBuffer(saved, { session });
    expect(accepted.acceptAll().skipped).toEqual([]);
    expect(accepted.getComments().map(({ id }) => id)).toEqual([8]);
    const rejected = await FolioDocxReviewer.fromBuffer(saved, { session });
    expect(rejected.rejectAll().skipped).toEqual([]);
    expect(
      rejected.getComments().map(({ id, anchoredText }) => [id, anchoredText]),
    ).toEqual([
      [3, "clause"],
      [8, "Second clause."],
      [12, "clause"],
    ]);
  });

  test(`${session} attributes retirement to the split separator deletion`, async () => {
    const bytes = await paragraphsDocx(
      [
        [textRun("Left"), ...anchor(3, "|"), textRun("Right")],
        anchor(8, "Keep."),
        [textRun("End.")],
      ],
      [reviewComment(3), reviewComment(8)],
    );
    const reviewer = await FolioDocxReviewer.fromBuffer(bytes, { session });
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "tracked-changes",
        operations: [
          {
            id: "split",
            type: "splitBlock",
            blockId: "10000000",
            offset: 4,
            separator: "|",
          },
          { ...edits[3], blockId: "10000002" },
        ],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    expect(
      result.receipts.map(({ operationId, retiredCommentIds }) => [
        operationId,
        retiredCommentIds,
      ]),
    ).toEqual([
      ["split", [3]],
      ["replace-end", undefined],
    ]);

    const saved = await reviewer.toBuffer();
    const accepted = await FolioDocxReviewer.fromBuffer(saved, { session });
    expect(accepted.acceptAll().skipped).toEqual([]);
    expect(accepted.getComments().map(({ id }) => id)).toEqual([8]);
    const rejected = await FolioDocxReviewer.fromBuffer(saved, { session });
    expect(rejected.rejectAll().skipped).toEqual([]);
    expect(
      rejected.getComments().map(({ id, anchoredText }) => [id, anchoredText]),
    ).toEqual([
      [3, "|"],
      [8, "Keep."],
    ]);
  });

  test(`${session} retires an anchor whose last words a later deletion removes`, async () => {
    const reviewer = await open(session);
    const apply = (operation: FolioDocumentOperation) =>
      reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations: [operation],
        },
        { undefinedReferences: "refuse" },
      );
    expect(apply(edits[1]).receipts.at(0)).not.toHaveProperty(
      "retiredCommentIds",
    );
    // The second deletion meets the first one's words, pending already.
    const rest = apply({
      id: "delete-rest",
      type: "replaceInBlock",
      blockId: "10000001",
      find: "clause.",
      replace: "",
    });
    expect(rest.skipped).toEqual([]);
    expect(rest.receipts.at(0)?.retiredCommentIds).toEqual([8]);
    expect(reviewer.acceptAll().skipped).toEqual([]);
    expect(reviewer.getComments().map(({ id }) => id)).toEqual([3, 12]);
  });

  test(`${session} refuses a jointly exhausted anchor without assigning an operation receipt`, async () => {
    const reviewer = await open(session);
    const before = structuredClone(reviewer.toDocument());
    const apply = (operations: readonly FolioDocumentOperation[]) =>
      reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode: "tracked-changes",
          operations,
        },
        { undefinedReferences: "refuse" },
      );
    const joint = apply([
      edits[1],
      {
        id: "delete-rest",
        type: "replaceInBlock",
        blockId: "10000001",
        find: "clause.",
        replace: "",
      },
    ]);
    expect(joint.status).toBe("rejected");
    expect(joint.applied).toEqual([]);
    expect(joint.receipts).toEqual([]);
    expect(joint.skipped).toEqual(
      ["delete-part", "delete-rest"].map((id) => ({
        id,
        reason: "invalidResult",
        message: "Accept All would lose a comment.",
      })),
    );
    expect(structuredClone(reviewer.toDocument())).toEqual(before);
    expect(reviewer.getChanges()).toEqual([]);
    const other = apply([edits[3]]);
    expect(other.applied).toHaveLength(1);
    expect(other.receipts.at(0)).not.toHaveProperty("retiredCommentIds");
  });

  for (const mode of ["direct", "suggested"] as const)
    test(`${session} omits retirement metadata in ${mode} mode`, async () => {
      const reviewer = await open(session);
      const result = reviewer.applyDocumentOperations(
        {
          version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
          mode,
          operations: [session === "canonical" ? inlineEdits[0] : edits[0]],
        },
        { undefinedReferences: "refuse" },
      );
      expect(result.skipped).toEqual([]);
      expect(result.receipts.at(0)).not.toHaveProperty("retiredCommentIds");
    });
}

for (const type of ["deleteTableRow", "deleteTableColumn"] as const)
  test(`legacy reports retirement for ${type}`, async () => {
    const cell = (paraId: string, content: ParagraphContent[]) =>
      ({
        type: "tableCell",
        content: [{ type: "paragraph", paraId, content }],
      }) satisfies TableCell;
    const document = {
      package: {
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  cells: [
                    cell("10000000", [
                      ...anchor(3, "Remove."),
                      { type: "commentReference", id: 4 },
                    ]),
                    cell("10000001", [textRun("Keep column.")]),
                  ],
                },
                {
                  type: "tableRow",
                  cells: [
                    cell("10000002", [textRun("Keep row.")]),
                    cell("10000003", [textRun("Keep both.")]),
                  ],
                },
              ],
            },
            {
              type: "paragraph",
              paraId: "10000004",
              content: [textRun("End.")],
            },
          ],
          comments: [reviewComment(3), reviewComment(4)],
        },
      },
    } satisfies Document;
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(document),
      { session: "legacy" },
    );
    expect(reviewer.getComments().map(({ id }) => id)).toEqual([3, 4]);
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations: [{ id: "delete-cells", type, blockId: "10000000" }],
      },
      { undefinedReferences: "refuse" },
    );
    expect(result.skipped).toEqual([]);
    expect(reviewer.acceptAll().skipped).toEqual([]);
    const survived = reviewer.getComments().map(({ id }) => id);
    expect(survived).not.toContain(3);
    expect(result.receipts.at(0)?.retiredCommentIds).toEqual(
      [3, 4].filter((id) => !survived.includes(id)),
    );
  });
