import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { withoutPendingSuggestions } from "../suggestions";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});
const info = (id: number, suggestionId?: string) => ({
  id,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
  ...(suggestionId === undefined ? {} : { suggestionId }),
});
const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: "00000001", content }],
    },
  },
});
const contentOf = (document: Document) => {
  const [paragraph] = document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : [];
};

describe("a save without pending suggestions", () => {
  test("rejects every suggested revision and keeps ordinary tracked changes", () => {
    const tracked: ParagraphContent = {
      type: "insertion",
      info: info(3),
      content: [run(" kept")],
    };
    const saved = withoutPendingSuggestions(
      documentOf([
        run("Hello"),
        { type: "insertion", info: info(1, "s1"), content: [run(" there")] },
        { type: "deletion", info: info(2, "s1"), content: [run(" world")] },
        tracked,
      ]),
    ).unwrap();
    expect(contentOf(saved)).toEqual([run("Hello world"), tracked]);
  });

  test("leaves a document without suggestions as it is", () => {
    const document = documentOf([
      run("Hello"),
      { type: "insertion", info: info(1), content: [run("!")] },
    ]);
    expect(withoutPendingSuggestions(document).unwrap()).toBe(document);
  });
});

describe("setting revision info", () => {
  const adopted = { id: 1, author: "Adopter", date: "2026-02-02T00:00:00Z" };
  const suggested = documentOf([
    run("Hello"),
    { type: "insertion", info: info(1, "s1"), content: [run(" there")] },
    {
      type: "run",
      content: [{ type: "text", text: "!" }],
    },
    { type: "insertion", info: info(2, "s1"), content: [run("?")] },
  ]);
  const adopt = {
    type: DOCUMENT_OP_TYPES.SET_REVISION_INFO,
    story: OP_STORIES.MAIN,
    infos: [adopted, { ...adopted, id: 2 }],
  } as const;

  test("writes each revision's info and undoes exactly", () => {
    const applied = applyDocumentOp(suggested, adopt).unwrap();
    expect(
      contentOf(applied.document).flatMap((item) =>
        item.type === "insertion" ? [item.info] : [],
      ),
    ).toEqual([adopted, { ...adopted, id: 2 }]);
    expect(
      applyDocumentOps(applied.document, applied.inverse).unwrap().document,
    ).toEqual(suggested);
  });

  test("refuses an id the story does not hold", () => {
    expect(
      applyDocumentOp(suggested, {
        ...adopt,
        infos: [{ ...adopted, id: 9 }],
      }).isErr(),
    ).toBe(true);
  });
});
