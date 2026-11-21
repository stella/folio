import { expect, test } from "bun:test";

import {
  type BlockContent,
  type Document,
  type Paragraph,
  REVIEW_CARRIERS,
  type Table,
} from "../../model/document";
import {
  type ResolvedDocumentView,
  resolvedDocumentView,
} from "../resolvedView";

const info = { id: 3, author: "Reviewer", date: "2026-11-18" };
const text = (paraId: string, value: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: value }] }],
});
// A host model writes an inserted paragraph's mark and words under one id.
const inserted: Paragraph = {
  type: "paragraph",
  paraId: "00000002",
  content: [
    {
      type: "insertion",
      info,
      content: [{ type: "run", content: [{ type: "text", text: "Added" }] }],
    },
  ],
  pPrMark: { kind: "ins", info },
};
const document: Document = {
  package: {
    document: {
      content: [text("00000001", "Kept"), inserted, text("00000003", "Tail")],
    },
  },
};

for (const decision of ["accept", "reject"] as const)
  test(`records sharing a revision id resolve together (${decision})`, () => {
    const view = resolvedDocumentView(document, decision);
    expect(view.unresolved).toEqual([]);
    const texts = view.document.package.document.content.map((block) =>
      JSON.stringify(block).includes("Added") ? "Added" : "other",
    );
    expect(texts).toEqual(
      decision === "accept" ? ["other", "Added", "other"] : ["other", "other"],
    );
    // The host's model is read, never changed.
    expect(document.package.document.content[1]).toBe(inserted);
  });

test("one mark object shared by several paragraphs resolves as a record in each", () => {
  const deletedMark = { kind: "del" as const, info: { id: 91, author: "A" } };
  const shared: Document = {
    package: {
      document: {
        content: [
          { ...text("00000001", "One "), pPrMark: deletedMark },
          { ...text("00000002", "One "), pPrMark: deletedMark },
          text("00000003", "Tail"),
        ],
      },
    },
  };
  const view = resolvedDocumentView(shared, "accept");
  expect(view.unresolved).toEqual([]);
  expect(view.document.package.document.content).toHaveLength(1);
  expect(JSON.stringify(view.document)).toContain("One One Tail");
});

for (const decision of ["accept", "reject"] as const)
  test(`a paragraph id repeated across stories resolves in each and keeps its id (${decision})`, () => {
    const story = (): Paragraph[] => [
      text("0000000A", "Kept"),
      {
        type: "paragraph",
        paraId: "0000000B",
        content: [
          {
            type: "insertion",
            info,
            content: [
              { type: "run", content: [{ type: "text", text: "Added" }] },
            ],
          },
        ],
      },
    ];
    const repeated: Document = {
      package: {
        document: { content: story() },
        headers: new Map([
          ["rId1", { type: "header", hdrFtrType: "default", content: story() }],
        ]),
      },
    };
    const view = resolvedDocumentView(repeated, decision);
    expect(view.unresolved).toEqual([]);
    const read = (blocks: readonly unknown[]) =>
      blocks.map((block) => {
        const { paraId, content } = block as Paragraph;
        return { paraId, added: JSON.stringify(content).includes("Added") };
      });
    const expected = [
      { paraId: "0000000A", added: false },
      { paraId: "0000000B", added: decision === "accept" },
    ];
    expect(read(view.document.package.document.content)).toEqual(expected);
    expect(
      read(view.document.package.headers?.get("rId1")?.content ?? []),
    ).toEqual(expected);
  });

const table = (paraId: string, deleted: boolean): Table => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      ...(deleted
        ? { structuralChange: { type: "tableRowDeletion" as const, info } }
        : {}),
      cells: [{ type: "tableCell", content: [text(paraId, "Cell")] }],
    },
  ],
});
const bodyOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});
/** Each body block: a paragraph's id and what it holds, or its type. */
const shape = (view: ResolvedDocumentView): string[] =>
  view.document.package.document.content.map((block) => {
    if (block.type !== "paragraph") return block.type;
    const held = JSON.stringify(block.content).includes("Added")
      ? "Added"
      : String(block.content.length);
    return `${block.paraId ?? "unnamed"}:${held}`;
  });

for (const decision of ["accept", "reject"] as const)
  test(`a body that ends in a table keeps that ending where resolution leaves its table (${decision})`, () => {
    const view = resolvedDocumentView(
      bodyOf([inserted, table("00000005", false)]),
      decision,
    );
    expect(view.unresolved).toEqual([]);
    expect(shape(view)).toEqual(
      decision === "accept" ? ["00000002:Added", "table"] : ["table"],
    );
  });

test("a body that is one deleted table resolves to the paragraph it closes with", () => {
  const view = resolvedDocumentView(
    bodyOf([table("00000005", true)]),
    "accept",
  );
  expect(view.unresolved).toEqual([]);
  // The closing paragraph's id is the view's own: no package names it.
  expect(shape(view)).toEqual(["unnamed:0"]);
});

test("a deleted table after a paragraph resolves to that paragraph", () => {
  const view = resolvedDocumentView(
    bodyOf([text("00000001", "Kept"), table("00000005", true)]),
    "accept",
  );
  expect(view.unresolved).toEqual([]);
  expect(shape(view)).toEqual(["00000001:1"]);
});

for (const decision of ["accept", "reject"] as const)
  test(`a terminal-table review carrier resolves as its table does (${decision})`, () => {
    const carrier: Paragraph = {
      ...inserted,
      paraId: "00000006",
      reviewCarrier: REVIEW_CARRIERS.TERMINAL_TABLE,
    };
    const view = resolvedDocumentView(
      bodyOf([text("00000001", "Kept"), table("00000005", true), carrier]),
      decision,
    );
    expect(view.unresolved).toEqual([]);
    // Accepted, the table goes and the carrier stays a plain paragraph;
    // rejected, the carrier goes and the body ends in its table again.
    expect(shape(view)).toEqual(
      decision === "accept"
        ? ["00000001:1", "00000006:Added"]
        : ["00000001:1", "table"],
    );
    expect(JSON.stringify(view.document)).not.toContain("reviewCarrier");
  });
