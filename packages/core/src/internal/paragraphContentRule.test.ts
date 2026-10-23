import { describe, expect, test } from "bun:test";

import {
  applyDocumentOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
} from "@stll/docx-core/ops";

import { createHarnessState } from "../__tests__/editorHarness";
import { acceptAllChanges } from "../prosemirror/commands/comments";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { BlockContent, Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";

// The editor and the engine decide a paragraph-mark outcome by one content
// rule: accepting the same deleted mark before a table must leave the same
// blocks, with any range boundary the paragraph held still in the document.

const info = { id: 7, author: "Reviewer", date: "2026-10-19" };

const documentWith = (content: Paragraph["content"]): Document => {
  const document = createEmptyDocument();
  document.package.document.content = [
    {
      type: "paragraph",
      paraId: "00000001",
      content,
      pPrMark: { kind: "del", info },
    },
    {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [
                {
                  type: "paragraph",
                  paraId: "00000002",
                  content: [
                    { type: "run", content: [{ type: "text", text: "cell" }] },
                  ],
                },
              ],
            },
          ],
        },
      ],
    },
    {
      type: "paragraph",
      paraId: "00000003",
      content: [{ type: "run", content: [{ type: "text", text: "after" }] }],
    },
  ];
  return document;
};

/** What the outcome decides: the block kinds left, and which range boundaries survive where. */
const outcome = (blocks: readonly BlockContent[]) => ({
  kinds: blocks.map((block) => block.type),
  boundariesInTable: [
    ...JSON.stringify(blocks[0]).matchAll(/"type":"(bookmark(?:Start|End))"/gu),
  ]
    .map((match) => match[1])
    .toSorted(),
});

const acceptInEditor = (document: Document) => {
  const state = createHarnessState(document, "editing");
  let resolved = state;
  expect(
    acceptAllChanges()(state, (tr) => {
      resolved = state.apply(tr);
    }),
  ).toBe(true);
  return fromProseDoc(resolved.doc, document, {
    stylesheetSource: { type: "package" },
  }).package.document.content;
};

const acceptInEngine = (document: Document) => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [info.id],
    decision: REVISION_DECISIONS.ACCEPT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document.package.document.content;
};

describe("editor and engine share one paragraph-content rule", () => {
  test("a bookmark-only paragraph whose mark is accepted before a table runs into the cell in both", () => {
    const document = documentWith([
      { type: "bookmarkStart", id: 1, name: "kept" },
      { type: "bookmarkEnd", id: 1 },
    ]);
    const editor = outcome(acceptInEditor(document));
    expect(editor).toEqual(outcome(acceptInEngine(document)));
    expect(editor).toEqual({
      kinds: ["table", "paragraph"],
      boundariesInTable: ["bookmarkEnd", "bookmarkStart"],
    });
  });

  // A content control is counted in place, as the engine counts it: holding
  // nothing, or only a marker, it is no width in either.
  for (const [label, content] of [
    ["an empty content control", []],
    [
      "a content control holding only a bookmark",
      [
        { type: "bookmarkStart", id: 2, name: "inside" },
        { type: "bookmarkEnd", id: 2 },
      ],
    ],
  ] as const) {
    test(`${label} decides the same outcome in both`, () => {
      const document = documentWith([
        {
          type: "inlineSdt",
          properties: { sdtType: "richText" },
          content: [...content],
        },
      ]);
      expect(outcome(acceptInEditor(document))).toEqual(
        outcome(acceptInEngine(document)),
      );
    });
  }

  test("a paragraph holding only a cached page break is removed in both", () => {
    const document = documentWith([
      { type: "run", content: [{ type: "renderedPageBreak" }] },
    ]);
    const editor = outcome(acceptInEditor(document));
    expect(editor).toEqual(outcome(acceptInEngine(document)));
    expect(editor.kinds).toEqual(["table", "paragraph"]);
  });
});
