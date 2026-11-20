import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { resolvedDocumentView } from "../resolvedView";

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
