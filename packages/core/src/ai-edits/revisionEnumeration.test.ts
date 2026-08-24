import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";

import { createDocx } from "../docx/rezip";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { acceptAIEditRevision, rejectAIEditRevision } from "../prosemirror/commands/comments";
import type { BlockContent, Paragraph, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import {
  FOLIO_RESOLVED_REVIEWED_VIEWS,
  type FolioEditableDocumentStoryHandle,
  FolioDocxReviewer,
} from "./headless";
import { FOLIO_REVIEW_CHANGE_KINDS, getTrackedChangesFromDoc } from "./read";

const AUTHOR = "Reviewer";
const DATE = "2026-08-16T10:00:00Z";

const REVISION = {
  insertion: 1,
  deletion: 2,
  formatting: 3,
  moveFrom: 4,
  moveTo: 5,
  paragraphMarkInserted: 101,
  paragraphMarkDeleted: 102,
  paragraphPropertiesChanged: 103,
  sectionPropertiesChanged: 104,
  tablePropertiesChanged: 105,
  rowPropertiesChanged: 106,
  cellPropertiesChanged: 107,
  rowInserted: 201,
  rowDeleted: 202,
  cellInserted: 203,
  cellDeleted: 204,
  cellMerged: 205,
} as const;

const revisionInfo = (id: number) => ({ id, author: AUTHOR, date: DATE });

const paragraphContent = (text: string): Paragraph["content"] => [
  { type: "run", content: [{ type: "text", text }] },
];

const revisionContent = (): BlockContent[] => {
  const paragraphs = [
    {
      type: "paragraph",
      paraId: "A0000100",
      content: [
        { type: "run", content: [{ type: "text", text: "Stable <&> 日本語 " }] },
        {
          type: "insertion",
          info: revisionInfo(REVISION.insertion),
          content: [{ type: "run", content: [{ type: "text", text: "Inserted" }] }],
        },
        {
          type: "deletion",
          info: revisionInfo(REVISION.deletion),
          content: [{ type: "run", content: [{ type: "text", text: "Deleted" }] }],
        },
        {
          type: "run",
          formatting: { italic: true },
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: revisionInfo(REVISION.formatting),
              previousFormatting: { bold: true },
            },
          ],
          content: [{ type: "text", text: " Formatted" }],
        },
        {
          type: "moveFrom",
          info: revisionInfo(REVISION.moveFrom),
          content: [{ type: "run", content: [{ type: "text", text: " Moved from" }] }],
        },
        {
          type: "moveTo",
          info: revisionInfo(REVISION.moveTo),
          content: [{ type: "run", content: [{ type: "text", text: " Moved to" }] }],
        },
      ],
    },
    {
      type: "paragraph",
      paraId: "A0000101",
      pPrMark: { kind: "ins", info: revisionInfo(REVISION.paragraphMarkInserted) },
      content: paragraphContent("Inserted boundary"),
    },
    {
      type: "paragraph",
      paraId: "A0000102",
      content: paragraphContent("Insertion join target"),
    },
    {
      type: "paragraph",
      paraId: "A0000103",
      pPrMark: { kind: "del", info: revisionInfo(REVISION.paragraphMarkDeleted) },
      content: paragraphContent("Deleted boundary"),
    },
    {
      type: "paragraph",
      paraId: "A0000104",
      content: paragraphContent("Deletion join target"),
    },
    {
      type: "paragraph",
      paraId: "A0000105",
      formatting: { alignment: "center" },
      propertyChanges: [
        {
          type: "paragraphPropertyChange",
          info: revisionInfo(REVISION.paragraphPropertiesChanged),
          previousFormatting: { alignment: "left" },
        },
      ],
      sectionProperties: {
        sectionStart: "nextPage",
        propertyChanges: [
          {
            type: "sectionPropertyChange",
            info: revisionInfo(REVISION.sectionPropertiesChanged),
            previousProperties: { sectionStart: "continuous" },
          },
        ],
      },
      content: paragraphContent("Property changes"),
    },
  ] satisfies Paragraph[];
  const table = {
    type: "table",
    formatting: { justification: "center" },
    propertyChanges: [
      {
        type: "tablePropertyChange",
        info: revisionInfo(REVISION.tablePropertiesChanged),
        previousFormatting: { justification: "left" },
      },
    ],
    rows: [
      {
        type: "tableRow",
        formatting: { height: { value: 500, type: "dxa" } },
        propertyChanges: [
          {
            type: "tableRowPropertyChange",
            info: revisionInfo(REVISION.rowPropertiesChanged),
            previousFormatting: { height: { value: 300, type: "dxa" } },
          },
        ],
        cells: [
          {
            type: "tableCell",
            formatting: { verticalAlign: "center" },
            propertyChanges: [
              {
                type: "tableCellPropertyChange",
                info: revisionInfo(REVISION.cellPropertiesChanged),
                previousFormatting: { verticalAlign: "top" },
              },
            ],
            content: [
              {
                type: "paragraph",
                paraId: "A0000106",
                content: paragraphContent("Cell"),
              },
            ],
          },
        ],
      },
      {
        type: "tableRow",
        structuralChange: {
          type: "tableRowInsertion",
          info: revisionInfo(REVISION.rowInserted),
        },
        cells: [
          {
            type: "tableCell",
            content: [
              {
                type: "paragraph",
                paraId: "A0000107",
                content: paragraphContent("Inserted row"),
              },
            ],
          },
        ],
      },
      {
        type: "tableRow",
        structuralChange: {
          type: "tableRowDeletion",
          info: revisionInfo(REVISION.rowDeleted),
        },
        cells: [
          {
            type: "tableCell",
            content: [
              {
                type: "paragraph",
                paraId: "A0000108",
                content: paragraphContent("Deleted row"),
              },
            ],
          },
        ],
      },
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            structuralChange: {
              type: "tableCellInsertion",
              info: revisionInfo(REVISION.cellInserted),
            },
            content: [
              {
                type: "paragraph",
                paraId: "A0000109",
                content: paragraphContent("Inserted cell"),
              },
            ],
          },
          {
            type: "tableCell",
            structuralChange: {
              type: "tableCellDeletion",
              info: revisionInfo(REVISION.cellDeleted),
            },
            content: [
              {
                type: "paragraph",
                paraId: "A0000110",
                content: paragraphContent("Deleted cell"),
              },
            ],
          },
        ],
      },
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            formatting: { vMerge: "restart" },
            content: [
              {
                type: "paragraph",
                paraId: "A0000111",
                content: paragraphContent("Merge origin"),
              },
            ],
          },
        ],
      },
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            formatting: { vMerge: "continue" },
            structuralChange: {
              type: "tableCellMerge",
              info: revisionInfo(REVISION.cellMerged),
              verticalMerge: "continue",
              verticalMergeOriginal: "rest",
            },
            content: [
              {
                type: "paragraph",
                paraId: "A0000112",
                content: paragraphContent("Merge continuation"),
              },
            ],
          },
        ],
      },
    ],
  } satisfies Table;
  return [...paragraphs, table];
};

const revisionDocument = () => {
  const document = createEmptyDocument();
  document.package.document.content = revisionContent();
  return document;
};

const expectedChanges = [
  { id: REVISION.insertion, type: "insertion", text: "Inserted" },
  { id: REVISION.deletion, type: "deletion", text: "Deleted" },
  { id: REVISION.formatting, type: "formatting", text: " Formatted" },
  { id: REVISION.moveFrom, type: "deletion", text: " Moved from" },
  { id: REVISION.moveTo, type: "insertion", text: " Moved to" },
  { id: REVISION.paragraphMarkInserted, type: "paragraphMarkInserted", text: "Inserted boundary" },
  { id: REVISION.paragraphMarkDeleted, type: "paragraphMarkDeleted", text: "Deleted boundary" },
  {
    id: REVISION.paragraphPropertiesChanged,
    type: "paragraphPropertiesChanged",
    text: "Property changes",
  },
  {
    id: REVISION.sectionPropertiesChanged,
    type: "sectionPropertiesChanged",
    text: "Property changes",
  },
  {
    id: REVISION.tablePropertiesChanged,
    type: "tablePropertiesChanged",
    text: "CellInserted rowDeleted rowInserted cellDeleted cellMerge originMerge continuation",
  },
  { id: REVISION.rowPropertiesChanged, type: "rowPropertiesChanged", text: "Cell" },
  { id: REVISION.cellPropertiesChanged, type: "cellPropertiesChanged", text: "Cell" },
  { id: REVISION.rowInserted, type: "rowInserted", text: "Inserted row" },
  { id: REVISION.rowDeleted, type: "rowDeleted", text: "Deleted row" },
  { id: REVISION.cellInserted, type: "cellInserted", text: "Inserted cell" },
  { id: REVISION.cellDeleted, type: "cellDeleted", text: "Deleted cell" },
  { id: REVISION.cellMerged, type: "cellMerged", text: "Merge continuation" },
] as const;

const HEADER_RELATIONSHIP_ID = "rId_revision_header";
const FOOTER_RELATIONSHIP_ID = "rId_revision_footer";
const PRESERVATION_SENTINEL_PART = "customXml/revision-preservation.xml";
const PRESERVATION_SENTINEL = new TextEncoder().encode(
  '<?xml version="1.0" encoding="UTF-8"?><preserve value="&lt;&amp; 日本語"/>',
);

const REVIEW_STORY_HANDLES = {
  main: { type: "main" },
  header: { type: "header", relationshipId: HEADER_RELATIONSHIP_ID },
  footer: { type: "footer", relationshipId: FOOTER_RELATIONSHIP_ID },
  footnote: { type: "footnote", noteId: 2 },
  endnote: { type: "endnote", noteId: 3 },
} as const satisfies Record<
  FolioEditableDocumentStoryHandle["type"],
  FolioEditableDocumentStoryHandle
>;

const REVIEW_STORIES = Object.values(REVIEW_STORY_HANDLES);

const makeRevisionMatrixDocument = () => {
  const document = revisionDocument();
  document.package.headers = new Map([
    [
      HEADER_RELATIONSHIP_ID,
      { type: "header", hdrFtrType: "default", content: structuredClone(revisionContent()) },
    ],
  ]);
  document.package.footers = new Map([
    [
      FOOTER_RELATIONSHIP_ID,
      { type: "footer", hdrFtrType: "default", content: structuredClone(revisionContent()) },
    ],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER_RELATIONSHIP_ID }],
    footerReferences: [{ type: "default", rId: FOOTER_RELATIONSHIP_ID }],
  };
  document.package.footnotes = [
    { type: "footnote", id: 2, noteType: "normal", content: structuredClone(revisionContent()) },
  ];
  document.package.endnotes = [
    { type: "endnote", id: 3, noteType: "normal", content: structuredClone(revisionContent()) },
  ];
  return document;
};

const makeRevisionMatrixDocx = async (paragraphIds: "with" | "without"): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createDocx(makeRevisionMatrixDocument()));
  if (paragraphIds === "without") {
    for (const path of [
      "word/document.xml",
      "word/header1.xml",
      "word/footer1.xml",
      "word/footnotes.xml",
      "word/endnotes.xml",
    ]) {
      const file = zip.file(path);
      if (!file) {
        throw new Error(`revision matrix is missing ${path}`);
      }
      zip.file(path, (await file.async("text")).replace(/ w14:paraId="[^"]+"/gu, ""));
    }
  }
  zip.file(PRESERVATION_SENTINEL_PART, PRESERVATION_SENTINEL);
  return zip.generateAsync({ type: "arraybuffer" });
};

const storyKey = (story: FolioEditableDocumentStoryHandle): string => JSON.stringify(story);

const commentProjection = (reviewer: FolioDocxReviewer) =>
  reviewer.getComments().map(({ id, author, text, anchoredText, blockId, replies, done }) => ({
    id,
    author,
    text,
    anchoredText,
    blockId,
    replies: replies.map(({ id: replyId, author: replyAuthor, text: replyText }) => ({
      id: replyId,
      author: replyAuthor,
      text: replyText,
    })),
    done,
  }));

const storyProjection = (reviewer: FolioDocxReviewer, story: FolioEditableDocumentStoryHandle) => {
  const reviewed = reviewer.readReviewedStory({ story, view: "current-markup" });
  if (!reviewed) {
    throw new Error(`revision matrix is missing ${storyKey(story)}`);
  }
  return {
    blocks: reviewed.snapshot.blocks.map(({ kind, text, previewRuns }) => ({
      kind,
      text,
      previewRuns,
    })),
    changes: reviewed.changes,
  };
};

const partBytes = async (buffer: ArrayBuffer, path: string): Promise<Uint8Array> => {
  const file = (await JSZip.loadAsync(buffer)).file(path);
  if (!file) {
    throw new Error(`revision matrix is missing ${path}`);
  }
  return file.async("uint8array");
};

const MATRIX_CASES = FOLIO_RESOLVED_REVIEWED_VIEWS.flatMap((view) =>
  (["with", "without"] as const).flatMap((paragraphIds) =>
    (["none", "thread"] as const).map((comments) => ({ view, paragraphIds, comments })),
  ),
);

describe("body revision enumeration", () => {
  test("enumerates every property and paragraph-mark carrier the resolver supports", () => {
    const changes = getTrackedChangesFromDoc(toProseDoc(revisionDocument()));

    expect(changes.map(({ id, type, text }) => ({ id, type, text }))).toEqual(expectedChanges);
    expect(
      changes.every(({ author, date, blockId }) => author === AUTHOR && date === DATE && blockId),
    ).toBe(true);
  });

  test.each(["accept", "reject"] as const)(
    "%s resolves each enumerated carrier without consuming the others",
    (mode) => {
      for (const selected of expectedChanges) {
        let state = EditorState.create({ doc: toProseDoc(revisionDocument()) });
        const command =
          mode === "accept" ? acceptAIEditRevision(selected.id) : rejectAIEditRevision(selected.id);

        expect(
          command(state, (transaction) => {
            state = state.apply(transaction);
          }),
        ).toBe(true);

        const roundtripped = toProseDoc(fromProseDoc(state.doc));
        const remainingIds = getTrackedChangesFromDoc(roundtripped).map(({ id }) => id);
        expect(remainingIds).not.toContain(selected.id);
        expect(remainingIds.toSorted((left, right) => left - right)).toEqual(
          expectedChanges
            .filter(({ id }) => id !== selected.id)
            .map(({ id }) => id)
            .toSorted((left, right) => left - right),
        );
      }
    },
  );

  test.each(["accept", "reject"] as const)(
    "headless %s persists a fully resolved revision census",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(revisionDocument()));
      const changes = reviewer.getChanges();
      expect(changes.map(({ id }) => id).toSorted((left, right) => left - right)).toEqual(
        expectedChanges.map(({ id }) => id).toSorted((left, right) => left - right),
      );

      for (const change of changes) {
        expect(
          mode === "accept" ? reviewer.acceptChange(change) : reviewer.rejectChange(change),
        ).toBe(true);
      }

      const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
      expect(reopened.getChanges()).toEqual([]);
    },
  );
});

describe("resolved story serialization structural matrix", () => {
  test.each(MATRIX_CASES)(
    "preserves every revision kind in $view view, $paragraphIds paragraph ids, comments: $comments",
    async ({ view, paragraphIds, comments }) => {
      const baseline = await makeRevisionMatrixDocx(paragraphIds);
      const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: AUTHOR });

      if (comments === "thread") {
        const target = reviewer.snapshot().blocks.find(({ text }) => text.includes("Stable <&>"));
        if (!target) {
          throw new Error("revision matrix is missing the stable comment target");
        }
        const result = reviewer.applyOperations([
          {
            id: "matrix-comment",
            type: "commentOnBlock",
            blockId: target.id,
            comment: { text: "Parent <&> 日本語" },
          },
        ]);
        expect(result.applied).toHaveLength(1);
        const parent = reviewer.getComments().at(0);
        if (!parent) {
          throw new Error("revision matrix did not create the parent comment");
        }
        expect(
          reviewer.replyTo(parent, { author: "Second reviewer", text: "Reply <&> العربية" }),
        ).not.toBeNull();
        expect(reviewer.resolveComment(String(parent.id))).toBe(true);
      }

      const expectedByStory = new Map<string, ReturnType<typeof storyProjection>>();
      for (const story of REVIEW_STORIES) {
        const current = reviewer.readReviewedStory({ story, view: "current-markup" });
        const resolved = reviewer.readReviewedStory({ story, view });
        if (!current || !resolved) {
          throw new Error(`revision matrix is missing ${storyKey(story)}`);
        }
        expect(new Set(current.changes.map(({ type }) => type))).toEqual(
          new Set(Object.values(FOLIO_REVIEW_CHANGE_KINDS)),
        );
        expect(reviewer.resolveReviewedStory({ story, view })).toBe(true);
        expectedByStory.set(storyKey(story), storyProjection(reviewer, story));
      }
      const expectedComments = commentProjection(reviewer);

      const saved = await reviewer.toBuffer();
      expect(await partBytes(saved, PRESERVATION_SENTINEL_PART)).toEqual(PRESERVATION_SENTINEL);
      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      for (const story of REVIEW_STORIES) {
        const actual = storyProjection(reopened, story);
        expect(actual.changes).toEqual([]);
        expect(actual).toEqual(expectedByStory.get(storyKey(story)));
      }
      expect(commentProjection(reopened)).toEqual(expectedComments);

      for (const story of REVIEW_STORIES) {
        expect(reopened.resolveReviewedStory({ story, view })).toBe(true);
      }
      const savedAgain = await reopened.toBuffer();
      expect(await partBytes(savedAgain, PRESERVATION_SENTINEL_PART)).toEqual(
        PRESERVATION_SENTINEL,
      );
      const reopenedAgain = await FolioDocxReviewer.fromBuffer(savedAgain);
      for (const story of REVIEW_STORIES) {
        expect(storyProjection(reopenedAgain, story)).toEqual(expectedByStory.get(storyKey(story)));
      }
      expect(commentProjection(reopenedAgain)).toEqual(expectedComments);
    },
    30_000,
  );
});
