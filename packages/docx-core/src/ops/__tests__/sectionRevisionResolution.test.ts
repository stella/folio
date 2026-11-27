import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const info = (id: number) => ({
  id,
  author: "Reviewer",
  date: "2026-11-20T00:00:00Z",
});

const text = (paraId: string, value: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: value }] }],
});

const headerReferences = [{ type: "default" as const, rId: "rIdNow" }];

/** A body whose first section's end carries a tracked section change. */
const sectionDocument = (): Document => ({
  package: {
    document: {
      content: [
        text("00000001", "Opening"),
        {
          ...text("00000002", "Section end"),
          sectionProperties: {
            sectionStart: "nextPage",
            headerReferences,
            propertyChanges: [
              {
                type: "sectionPropertyChange",
                info: info(7),
                previousProperties: { sectionStart: "continuous" },
              },
            ],
          },
        },
        {
          type: "paragraph",
          paraId: "00000003",
          content: [
            {
              type: "insertion",
              info: info(8),
              content: [
                { type: "run", content: [{ type: "text", text: "Added" }] },
              ],
            },
          ],
        },
      ],
    },
  },
});

const sectionEnd = (document: Document) => {
  const block = document.package.document.content.at(1);
  if (block?.type !== "paragraph")
    throw new Error("The section end stays a paragraph.");
  return block.sectionProperties;
};

const resolve = (revisionIds: number[], decision: RevisionDecision) => {
  const before = sectionDocument();
  const result = applyDocumentOp(before, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
  if (result.isErr()) throw result.error;
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toEqual(before);
  return result.value.document;
};

test("accepting a body section's change keeps its properties", () => {
  expect(sectionEnd(resolve([7], REVISION_DECISIONS.ACCEPT))).toEqual({
    sectionStart: "nextPage",
    headerReferences,
  });
});

test("rejecting a body section's change restores its properties and keeps its references", () => {
  expect(sectionEnd(resolve([7], REVISION_DECISIONS.REJECT))).toEqual({
    sectionStart: "continuous",
    headerReferences,
  });
});

for (const decision of Object.values(REVISION_DECISIONS))
  test(`a section change resolves with the inline changes beside it (${decision})`, () => {
    const resolved = resolve([7, 8], decision);
    expect(sectionEnd(resolved)?.propertyChanges).toBeUndefined();
    expect(JSON.stringify(resolved).includes("Added")).toBe(
      decision === REVISION_DECISIONS.ACCEPT,
    );
    expect(JSON.stringify(resolved)).not.toContain('"insertion"');
  });

/** A paragraph of a story with no sections that still carries section properties with a record. */
const storyParagraphWithSection = (paraId: string, id: number): Paragraph => ({
  ...text(paraId, "Story paragraph"),
  sectionProperties: {
    sectionStart: "nextPage",
    propertyChanges: [
      {
        type: "sectionPropertyChange",
        info: info(id),
        previousProperties: { sectionStart: "continuous" },
      },
    ],
  },
});

const STORIES = [
  { story: { kind: "header", rId: "rIdHeader" }, paraId: "00000011", id: 9 },
  { story: { kind: "footnote", id: 2 }, paraId: "00000012", id: 10 },
] as const;

const storyDocument = (): Document => {
  const document = sectionDocument();
  document.package.headers = new Map([
    [
      "rIdHeader",
      {
        type: "header",
        hdrFtrType: "default",
        content: [storyParagraphWithSection("00000011", 9)],
      },
    ],
  ]);
  document.package.footnotes = [
    {
      type: "footnote",
      id: 2,
      noteType: "normal",
      content: [storyParagraphWithSection("00000012", 10)],
    },
  ];
  return document;
};

for (const { story, paraId, id } of STORIES)
  for (const [decision, sectionStart] of [
    [REVISION_DECISIONS.ACCEPT, "nextPage"],
    [REVISION_DECISIONS.REJECT, "continuous"],
  ] as const)
    test(`a section change on a ${story.kind} paragraph resolves on that paragraph (${decision})`, () => {
      const before = storyDocument();
      const result = applyDocumentOp(before, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story,
        revisionIds: [id],
        decision,
      });
      if (result.isErr()) throw result.error;
      const content =
        story.kind === "header"
          ? result.value.document.package.headers?.get(story.rId)?.content
          : result.value.document.package.footnotes?.at(0)?.content;
      expect(content).toEqual([
        {
          ...text(paraId, "Story paragraph"),
          sectionProperties: { sectionStart },
        },
      ]);
      // The body's own section change belongs to another story: it stays.
      expect(sectionEnd(result.value.document)?.propertyChanges).toHaveLength(
        1,
      );
      const undone = applyDocumentOps(
        result.value.document,
        result.value.inverse,
      );
      if (undone.isErr()) throw undone.error;
      expect(undone.value.document).toEqual(before);
    });

type NestedCase = {
  name: string;
  wrap: (paragraph: Paragraph) => BlockContent;
  inner: (document: Document) => BlockContent | undefined;
};

/** Main-story containers whose paragraphs end no section. */
const NESTED: readonly NestedCase[] = [
  {
    name: "a block content control",
    wrap: (paragraph) => ({
      type: "blockSdt",
      properties: { sdtType: "richText", id: 40 },
      content: [paragraph],
    }),
    inner: (document) => {
      const control = document.package.document.content.at(0);
      return control?.type === "blockSdt" ? control.content.at(0) : undefined;
    },
  },
  {
    name: "a table cell",
    wrap: (paragraph) => ({
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [{ type: "tableCell", content: [paragraph] }],
        },
      ],
    }),
    inner: (document) => {
      const table = document.package.document.content.at(0);
      return table?.type === "table"
        ? table.rows.at(0)?.cells.at(0)?.content.at(0)
        : undefined;
    },
  },
];

// A main-story paragraph nested in a container ends no section, but a file
// may still give it section properties with a record: they resolve on it.
for (const { name, wrap, inner } of NESTED)
  for (const [decision, sectionStart] of [
    [REVISION_DECISIONS.ACCEPT, "nextPage"],
    [REVISION_DECISIONS.REJECT, "continuous"],
  ] as const)
    test(`a section change on a paragraph in ${name} resolves on that paragraph (${decision})`, () => {
      const before: Document = {
        package: {
          document: {
            content: [
              wrap(storyParagraphWithSection("00000021", 11)),
              text("00000022", "After"),
            ],
          },
        },
      };
      const result = applyDocumentOp(before, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [11],
        decision,
      });
      if (result.isErr()) throw result.error;
      expect(inner(result.value.document)).toEqual({
        ...text("00000021", "Story paragraph"),
        sectionProperties: { sectionStart },
      });
      const undone = applyDocumentOps(
        result.value.document,
        result.value.inverse,
      );
      if (undone.isErr()) throw undone.error;
      expect(undone.value.document).toEqual(before);
    });
