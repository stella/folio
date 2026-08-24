import { describe, expect, test } from "bun:test";
import JSZip from "jszip";
import { EditorState } from "prosemirror-state";

import { createDocx } from "../docx/rezip";
import {
  getAttributeByNamespaceUri,
  getChildElements,
  getLocalName,
  parseXmlDocument,
} from "../docx/xmlParser";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { acceptAIEditRevision, rejectAIEditRevision } from "../prosemirror/commands/comments";
import type { BlockContent, Paragraph, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import {
  FOLIO_RESOLVED_REVIEWED_VIEWS,
  type FolioEditableDocumentStoryHandle,
  FolioDocxReviewer,
  type FolioReviewComment,
  type FolioReviewCommentReply,
} from "./headless";
import { FOLIO_REVIEW_CHANGE_KINDS, getTrackedChangesFromDoc } from "./read";

const AUTHOR = "Reviewer";
const DATE = "2026-08-16T10:00:00Z";
const WORDPROCESSINGML_NAMESPACES = [
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
] as const;
const WORDPROCESSINGML_NAMESPACE_SET = new Set<string>(WORDPROCESSINGML_NAMESPACES);

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

const REVIEW_STORY_PART_PATHS = {
  main: "word/document.xml",
  header: "word/header1.xml",
  footer: "word/footer1.xml",
  footnote: "word/footnotes.xml",
  endnote: "word/endnotes.xml",
} as const satisfies Record<FolioEditableDocumentStoryHandle["type"], string>;

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

const REVIEW_COMMENT_FIELD_CENSUS = {
  id: "id",
  author: "author",
  date: "date",
  text: "text",
  anchoredText: "anchoredText",
  blockId: "blockId",
  replies: "replies",
  done: "done",
} as const satisfies Record<keyof FolioReviewComment, keyof FolioReviewComment>;

const REVIEW_COMMENT_REPLY_FIELD_CENSUS = {
  id: "id",
  author: "author",
  date: "date",
  text: "text",
} as const satisfies Record<keyof FolioReviewCommentReply, keyof FolioReviewCommentReply>;

const commentProjection = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getComments()
    .map(({ id, author, date, text, anchoredText, blockId, replies, done }) => ({
      id,
      author,
      date,
      text,
      anchoredText,
      blockId,
      replies: replies.map(
        ({ id: replyId, author: replyAuthor, date: replyDate, text: replyText }) => ({
          id: replyId,
          author: replyAuthor,
          date: replyDate,
          text: replyText,
        }),
      ),
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

const assertCommentAnchorsInStoryPart = async (
  buffer: ArrayBuffer,
  story: FolioEditableDocumentStoryHandle,
  commentIds: readonly number[],
): Promise<void> => {
  const xml = new TextDecoder().decode(
    await partBytes(buffer, REVIEW_STORY_PART_PATHS[story.type]),
  );
  for (const commentId of commentIds) {
    for (const localName of ["commentRangeStart", "commentRangeEnd"] as const) {
      const ids = commentAnchorIds(xml, localName);
      expect(ids).toContain(String(commentId));
    }
  }
};

const commentAnchorIds = (
  xml: string,
  localName: "commentRangeStart" | "commentRangeEnd",
): string[] => {
  const root = parseXmlDocument(xml);
  if (!root) {
    throw new Error("revision matrix could not parse a story part");
  }
  const ids: string[] = [];
  const pending = [root];
  while (pending.length > 0) {
    const element = pending.pop();
    if (!element) {
      continue;
    }
    if (
      element.namespaceUri !== undefined &&
      WORDPROCESSINGML_NAMESPACE_SET.has(element.namespaceUri) &&
      getLocalName(element.name) === localName
    ) {
      const id = getAttributeByNamespaceUri(element, WORDPROCESSINGML_NAMESPACE_SET, "id");
      if (id !== null) {
        ids.push(id);
      }
    }
    pending.push(...getChildElements(element));
  }
  return ids;
};

const PARAGRAPH_ID_STATES = ["with", "without"] as const;
const COMMENT_ROOT_COUNTS = [1, 2] as const;
const COMMENT_REPLY_STATES = ["without-reply", "with-reply"] as const;
const COMMENT_RESOLUTION_STATES = ["open", "resolved"] as const;
const REVIEW_COMMENT_SCENARIOS = [
  {
    name: "none",
    rootCount: 0,
    replyState: "without-reply",
    resolutionState: "open",
  } as const,
  ...COMMENT_ROOT_COUNTS.flatMap((rootCount) =>
    COMMENT_REPLY_STATES.flatMap((replyState) =>
      COMMENT_RESOLUTION_STATES.map((resolutionState) => ({
        name: `${rootCount === 1 ? "single" : "overlapping"}-${replyState}-${resolutionState}`,
        rootCount,
        replyState,
        resolutionState,
      })),
    ),
  ),
];

const MATRIX_CASES = FOLIO_RESOLVED_REVIEWED_VIEWS.flatMap((view) =>
  PARAGRAPH_ID_STATES.flatMap((paragraphIds) =>
    REVIEW_COMMENT_SCENARIOS.map((commentScenario) => ({
      view,
      paragraphIds,
      comments: commentScenario.name,
      commentScenario,
    })),
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
  test.each(WORDPROCESSINGML_NAMESPACES)(
    "recognizes comment anchors in the %s namespace independent of prefix",
    (namespace) => {
      const xml = `<alt:document xmlns:alt="${namespace}"><alt:commentRangeStart alt:id="42"/><alt:commentRangeEnd alt:id="42"/></alt:document>`;
      for (const localName of ["commentRangeStart", "commentRangeEnd"] as const) {
        expect(commentAnchorIds(xml, localName)).toContain("42");
      }
    },
  );

  test("enumerates the declared Cartesian product exactly once", () => {
    const keys = MATRIX_CASES.map(({ view, paragraphIds, comments }) =>
      JSON.stringify({ view, paragraphIds, comments }),
    );
    expect(MATRIX_CASES).toHaveLength(
      FOLIO_RESOLVED_REVIEWED_VIEWS.length *
        PARAGRAPH_ID_STATES.length *
        REVIEW_COMMENT_SCENARIOS.length,
    );
    expect(new Set(keys).size).toBe(keys.length);
    expect(Object.values(REVIEW_COMMENT_FIELD_CENSUS)).toHaveLength(8);
    expect(Object.values(REVIEW_COMMENT_REPLY_FIELD_CENSUS)).toHaveLength(4);
  });

  test.each(MATRIX_CASES)(
    "preserves every revision kind in $view view, $paragraphIds paragraph ids, comments: $comments",
    async ({ view, paragraphIds, commentScenario }) => {
      const baseline = await makeRevisionMatrixDocx(paragraphIds);
      const reviewer = await FolioDocxReviewer.fromBuffer(baseline, { author: AUTHOR });

      const commentIdsByStory = new Map<string, readonly number[]>();

      if (commentScenario.rootCount > 0) {
        for (const story of REVIEW_STORIES) {
          const commentIds: number[] = [];
          const parentIds: number[] = [];
          for (let rootIndex = 0; rootIndex < commentScenario.rootCount; rootIndex += 1) {
            const snapshot = reviewer.snapshotStory(story);
            const target = snapshot?.blocks.find(({ text }) => text.includes("Stable <&>"));
            if (!snapshot || !target) {
              throw new Error(
                `revision matrix is missing the comment target in ${storyKey(story)}`,
              );
            }
            const previousIds = new Set(reviewer.getComments().map(({ id }) => id));
            const result = reviewer.applyDocumentOperationsToStory({
              story,
              snapshot,
              batch: {
                version: 1,
                mode: "direct",
                operations: [
                  {
                    id: `matrix-comment-${story.type}-${rootIndex}`,
                    type: "commentOnBlock",
                    blockId: target.id,
                    comment: { text: `Parent ${story.type} ${rootIndex} <&> 日本語` },
                  },
                ],
              },
            });
            expect(result.status).toBe("committed");
            const parent = reviewer.getComments().find(({ id }) => !previousIds.has(id));
            if (!parent) {
              throw new Error(`revision matrix did not create a comment in ${storyKey(story)}`);
            }
            parentIds.push(parent.id);
            commentIds.push(parent.id);
          }
          if (commentScenario.replyState === "with-reply") {
            const parentId = parentIds.at(0);
            const reply =
              parentId === undefined
                ? null
                : reviewer.replyTo(parentId, {
                    author: "Second reviewer",
                    text: `Reply ${story.type} <&> العربية`,
                  });
            if (!reply) {
              throw new Error(`revision matrix did not create a reply in ${storyKey(story)}`);
            }
            commentIds.push(reply.id);
          }
          if (commentScenario.resolutionState === "resolved") {
            for (const parentId of parentIds) {
              expect(reviewer.resolveComment(String(parentId))).toBe(true);
            }
          }
          commentIdsByStory.set(storyKey(story), commentIds);
        }
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
      for (const story of REVIEW_STORIES) {
        await assertCommentAnchorsInStoryPart(
          saved,
          story,
          commentIdsByStory.get(storyKey(story)) ?? [],
        );
      }
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
        await assertCommentAnchorsInStoryPart(
          savedAgain,
          story,
          commentIdsByStory.get(storyKey(story)) ?? [],
        );
      }
      for (const story of REVIEW_STORIES) {
        expect(storyProjection(reopenedAgain, story)).toEqual(expectedByStory.get(storyKey(story)));
      }
      expect(commentProjection(reopenedAgain)).toEqual(expectedComments);
    },
    30_000,
  );
});
