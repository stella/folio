import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import { createDocx } from "../docx/rezip";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { acceptAIEditRevision, rejectAIEditRevision } from "../prosemirror/commands/comments";
import type { Paragraph, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { getTrackedChangesFromDoc } from "./read";

const AUTHOR = "Reviewer";
const DATE = "2026-08-16T10:00:00Z";

const REVISION = {
  paragraphMarkInserted: 101,
  paragraphMarkDeleted: 102,
  paragraphPropertiesChanged: 103,
  sectionPropertiesChanged: 104,
  tablePropertiesChanged: 105,
  rowPropertiesChanged: 106,
  cellPropertiesChanged: 107,
} as const;

const revisionInfo = (id: number) => ({ id, author: AUTHOR, date: DATE });

const paragraphContent = (text: string): Paragraph["content"] => [
  { type: "run", content: [{ type: "text", text }] },
];

const revisionDocument = () => {
  const document = createEmptyDocument();
  const paragraphs = [
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
    ],
  } satisfies Table;
  document.package.document.content = [...paragraphs, table];
  return document;
};

const expectedChanges = [
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
  { id: REVISION.tablePropertiesChanged, type: "tablePropertiesChanged", text: "Cell" },
  { id: REVISION.rowPropertiesChanged, type: "rowPropertiesChanged", text: "Cell" },
  { id: REVISION.cellPropertiesChanged, type: "cellPropertiesChanged", text: "Cell" },
] as const;

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
