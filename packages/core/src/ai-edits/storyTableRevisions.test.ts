import { expect, test } from "bun:test";

import { createDocx } from "../docx/rezip";
import type { BlockContent, Document, Table } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import {
  FolioDocxReviewer,
  type FolioEditableDocumentStoryHandle,
} from "./headless";

const info = (id: number) => ({
  id,
  author: "Reviewer",
  date: "2026-11-20T00:00:00Z",
});

const paragraph = (paraId: string, text: string): BlockContent => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/** A story holding a table whose own, row and cell properties are tracked. */
const storyContent = (prefix: string): BlockContent[] => {
  const table: Table = {
    type: "table",
    formatting: { justification: "center" },
    propertyChanges: [
      {
        type: "tablePropertyChange",
        info: info(Number.parseInt(`${prefix}1`, 10)),
        previousFormatting: { justification: "left" },
      },
    ],
    rows: [
      {
        type: "tableRow",
        formatting: { cantSplit: true },
        propertyChanges: [
          {
            type: "tableRowPropertyChange",
            info: info(Number.parseInt(`${prefix}2`, 10)),
            previousFormatting: {},
          },
        ],
        cells: [
          {
            type: "tableCell",
            formatting: { noWrap: true },
            propertyChanges: [
              {
                type: "tableCellPropertyChange",
                info: info(Number.parseInt(`${prefix}3`, 10)),
                previousFormatting: {},
              },
            ],
            content: [paragraph(`${prefix}000002`, "Cell")],
          },
        ],
      },
    ],
  };
  return [
    paragraph(`${prefix}000001`, "Before"),
    table,
    paragraph(`${prefix}000003`, "After"),
  ];
};

const HEADER = "rIdStoryTableHeader";
const FOOTER = "rIdStoryTableFooter";

const storyTableDocument = (): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [paragraph("10000001", "Body")];
  document.package.headers = new Map([
    [
      HEADER,
      { type: "header", hdrFtrType: "default", content: storyContent("2") },
    ],
  ]);
  document.package.footers = new Map([
    [
      FOOTER,
      { type: "footer", hdrFtrType: "default", content: storyContent("3") },
    ],
  ]);
  document.package.document.finalSectionProperties = {
    ...document.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: HEADER }],
    footerReferences: [{ type: "default", rId: FOOTER }],
  };
  return document;
};

const STORIES = [
  { type: "header", relationshipId: HEADER },
  { type: "footer", relationshipId: FOOTER },
] as const satisfies readonly FolioEditableDocumentStoryHandle[];

for (const session of ["legacy", "canonical"] as const)
  for (const view of ["final", "original"] as const)
    for (const story of STORIES)
      test(`a ${story.type} table's property revisions resolve to the ${view} view (${session})`, async () => {
        const reviewer = await FolioDocxReviewer.fromBuffer(
          await createDocx(storyTableDocument()),
          { session },
        );
        const pending = reviewer.readReviewedStory({
          story,
          view: "current-markup",
        });
        expect(pending?.changes.map(({ type }) => type).toSorted()).toEqual([
          "cellPropertiesChanged",
          "rowPropertiesChanged",
          "tablePropertiesChanged",
        ]);
        expect(reviewer.resolveReviewedStory({ story, view })).toBe(true);
        expect(
          reviewer.readReviewedStory({ story, view: "current-markup" })
            ?.changes,
        ).toEqual([]);
        const [table] = reviewer.storyTables({ story, view: "current-markup" });
        expect(table?.node.attrs["justification"]).toBe(
          view === "final" ? "center" : "left",
        );
      });
