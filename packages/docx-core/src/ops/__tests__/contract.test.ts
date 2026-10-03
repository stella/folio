import { describe, expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph, Run } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps, validateOpsDocument } from "../contract";
import { documentStories, storyBody } from "../stories";
import { storyParagraphs } from "../blocks";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp, INHERIT_RUN_PROPS, OP_STORIES } from "../types";

const run = (text: string): Run => ({ type: "run", content: [{ type: "text", text }] });

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [run(text)],
});

const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });

const typing = (blockId: string): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.INSERT_TEXT,
  at: at(blockId, 0),
  text: "x",
  runProps: INHERIT_RUN_PROPS,
});

const reasonOf = (document: Document, op: DocumentOp): string | undefined => {
  const result = applyDocumentOp(document, op);
  return result.isErr() ? result.error.reason : undefined;
};

/** A body whose section view holds the body's own records, as the parser builds it. */
const parsedShape = (...content: Paragraph[]): Document => ({
  package: {
    document: {
      content,
      sections: [{ properties: { pageWidth: 12240 }, content }],
      finalSectionProperties: { pageWidth: 12240 },
    },
  },
});

describe("the id census", () => {
  test("counts a paragraph the section view also holds once", () => {
    const document = parsedShape(paragraph("00000001", "a"), paragraph("00000002", "b"));
    const [first] = document.package.document.content;
    if (first?.type !== "paragraph") throw new Error("fixture");
    const replaced = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: OP_STORIES.MAIN,
      expected: [first],
      blocks: [paragraph("00000001", "kept id")],
    });
    expect(replaced.isOk()).toBe(true);
  });

  test("finds a paragraph by its id in either case", () => {
    const document = parsedShape(paragraph("0000ABCD", "a"));
    expect(reasonOf(document, typing("0000abcd"))).toBeUndefined();
  });
});

describe("the seed contract", () => {
  test("refuses a paragraph id another story also uses, before any operation", () => {
    const document: Document = {
      package: {
        document: { content: [paragraph("00000001", "a"), paragraph("00000002", "b")] },
        footnotes: [{ type: "footnote", id: 2, content: [paragraph("00000002", "note")] }],
      },
    };
    // Joining would retire 00000002, and its inverse could not create it again.
    expect(
      reasonOf(document, {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000002",
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_BLOCK_ID);
  });

  test("refuses a main-story paragraph without an id", () => {
    const document = parsedShape({ type: "paragraph", content: [run("a")] });
    expect(validateOpsDocument(document).isErr()).toBe(true);
    expect(reasonOf(document, typing("00000001"))).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.MISSING_BLOCK_ID,
    );
  });

  test("refuses two records carrying one revision id", () => {
    const insertion = (text: string) => ({
      type: "insertion" as const,
      info: { id: 5, author: "A" },
      content: [run(text)],
    });
    const document = parsedShape({
      type: "paragraph",
      paraId: "00000001",
      content: [insertion("a"), insertion("b")],
    });
    expect(reasonOf(document, typing("00000001"))).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_RECORD_ID,
    );
  });

  test("counts a row-level content control recorded on each of its rows once", () => {
    const control = { sdtType: "richText" as const, id: 9, tag: "rows" };
    const tableOf = (first: typeof control, second: typeof control): Document => ({
      package: {
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  type: "tableRow",
                  cells: [{ type: "tableCell", content: [paragraph("00000001", "a")] }],
                  contentControls: [first],
                },
                {
                  type: "tableRow",
                  cells: [{ type: "tableCell", content: [paragraph("00000002", "b")] }],
                  contentControls: [second],
                },
              ],
            },
          ],
        },
      },
    });
    // One control over two rows, whether or not its records are shared.
    expect(reasonOf(tableOf(control, control), typing("00000001"))).toBeUndefined();
    expect(reasonOf(tableOf(control, { ...control }), typing("00000001"))).toBeUndefined();
    // Two different controls carrying one id.
    expect(reasonOf(tableOf(control, { ...control, tag: "other" }), typing("00000001"))).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.DUPLICATE_RECORD_ID,
    );
  });

  test("refuses a section view out of step with the body", () => {
    const document: Document = {
      package: {
        document: {
          content: [paragraph("00000001", "a"), paragraph("00000002", "b")],
          sections: [{ properties: { pageWidth: 12240 }, content: [paragraph("00000001", "a")] }],
        },
      },
    };
    expect(reasonOf(document, typing("00000001"))).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.SECTIONS_OUT_OF_STEP,
    );
  });

  test.each(["insertion", "deletion", "moveFrom", "moveTo"] as const)(
    "normalization removes empty %s wrappers recursively and preserves zero-width markup",
    (type) => {
      const document = parsedShape({
        type: "paragraph",
        paraId: "00000001",
        content: [
          { type, info: { id: 1, author: "Reviewer" }, content: [] },
          {
            type,
            info: { id: 2, author: "Reviewer" },
            content: [{ type: "insertion", info: { id: 3, author: "Reviewer" }, content: [] }],
          },
          {
            type,
            info: { id: 4, author: "Reviewer" },
            content: [{ type: "bookmarkStart", id: 5, name: "anchor" }],
          },
          { type: "hyperlink", children: [] },
        ],
      });
      expect(reasonOf(document, typing("00000001"))).toBe(DOCUMENT_OP_REFUSAL_REASONS.EMPTY_RECORD);
      const normalized = normalizeForOps(document);
      expect(normalized.package.document.content).toEqual([
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            {
              type,
              info: { id: 4, author: "Reviewer" },
              content: [{ type: "bookmarkStart", id: 5, name: "anchor" }],
            },
            { type: "hyperlink", children: [] },
          ],
        },
      ]);
      expect(normalizeForOps(normalized)).toEqual(normalized);
      expect(validateOpsDocument(normalized).isOk()).toBe(true);
    },
  );

  test("normalizing removes empty runs and empty text nodes, and nothing else", () => {
    const document = parsedShape({
      type: "paragraph",
      paraId: "00000001",
      content: [
        { type: "run", content: [] },
        { type: "run", content: [{ type: "text", text: "" }] },
        { type: "hyperlink", children: [] },
        run("a"),
      ],
    });
    expect(reasonOf(document, typing("00000001"))).toBe(DOCUMENT_OP_REFUSAL_REASONS.EMPTY_RECORD);
    const normalized = normalizeForOps(document);
    expect(normalized.package.document.content).toEqual([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [{ type: "hyperlink", children: [] }, run("a")],
      },
    ]);
    expect(normalized.package.document.sections?.[0]?.content).toEqual(
      normalized.package.document.content,
    );
    expect(validateOpsDocument(normalized).isOk()).toBe(true);
  });
});

describe("the section view", () => {
  test("is derived from the body after an edit, whatever records it held", () => {
    const content = [paragraph("00000001", "a"), paragraph("00000002", "b")];
    const document: Document = {
      package: {
        document: {
          content,
          // A view built independently of the body: equal records, no sharing.
          sections: [{ properties: { pageWidth: 12240 }, content: structuredClone(content) }],
          finalSectionProperties: { pageWidth: 12240 },
        },
      },
    };
    const edited = applyDocumentOp(document, typing("00000002"));
    if (edited.isErr()) throw edited.error;
    const body = edited.value.document.package.document;
    expect(body.sections?.[0]?.content).toEqual(body.content);
    for (const [index, block] of (body.sections?.[0]?.content ?? []).entries()) {
      expect(block).toBe(body.content[index]!);
    }
    const undone = applyDocumentOps(edited.value.document, edited.value.inverse);
    expect(undone.isOk() ? undone.value.document : undefined).toEqual(document);
  });
});

test("paragraph property cardinality is enforced for seeds and operation outputs in every story", () => {
  const nested = (block: Paragraph): BlockContent => ({
    type: "blockSdt",
    properties: { sdtType: "richText" },
    content: [
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
                    type: "blockCustomXml",
                    openingXml: '<w:customXml w:element="record" w:uri="urn:fixture">',
                    closingXml: "</w:customXml>",
                    content: [block],
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  });
  const makeDocument = () =>
    ({
      package: {
        document: { content: [nested(paragraph("00000001", "main"))] },
        headers: new Map([
          [
            "rIdHeader",
            {
              type: "header",
              hdrFtrType: "default",
              content: [nested(paragraph("00000002", "header"))],
            },
          ],
        ]),
        footers: new Map([
          [
            "rIdFooter",
            {
              type: "footer",
              hdrFtrType: "default",
              content: [nested(paragraph("00000003", "footer"))],
            },
          ],
        ]),
        footnotes: [
          { type: "footnote", id: 1, content: [nested(paragraph("00000004", "footnote"))] },
        ],
        endnotes: [{ type: "endnote", id: 2, content: [nested(paragraph("00000005", "endnote"))] }],
      },
    }) satisfies Document;
  const propertyChanges = [101, 102].map((id) => ({
    type: "paragraphPropertyChange" as const,
    info: { id, author: "Reviewer" },
  }));
  for (const story of documentStories(makeDocument())) {
    const document = makeDocument();
    const target = storyParagraphs(storyBody(document, story)).at(0)?.paragraph;
    if (target?.paraId === undefined) throw new Error("Missing story fixture paragraph");
    const badSeed = structuredClone(document);
    const badParagraph = storyParagraphs(storyBody(badSeed, story)).at(0)?.paragraph;
    if (badParagraph === undefined) throw new Error("Missing story fixture paragraph");
    badParagraph.propertyChanges = propertyChanges;
    expect(validateOpsDocument(badSeed).isErr()).toBe(true);
    const operations = [
      {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
        story,
        blockId: target.paraId,
        expected: {},
        review: { propertyChanges },
      },
      {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        story,
        at: { type: "before", blockId: target.paraId },
        blocks: [{ ...paragraph("00000006", "inserted"), propertyChanges }],
      },
    ] as const satisfies readonly DocumentOp[];
    for (const op of operations) {
      expect(reasonOf(document, op)).toBe(DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH);
    }
    expect(document).toEqual(makeDocument());
  }
});
