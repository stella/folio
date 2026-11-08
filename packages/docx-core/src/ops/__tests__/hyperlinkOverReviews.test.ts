import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import type { HyperlinkEditorIntent } from "../hyperlinkIntent";
import { OP_STORIES } from "../types";

const info = (id: number) => ({
  id,
  author: "Reviewer",
  date: "2026-01-01T00:00:00Z",
});
const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: "12345678", content }],
    },
  },
});
const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "12345678",
  offset,
});
const linked = (from: number, to: number): HyperlinkEditorIntent => ({
  type: "setHyperlink",
  from: at(from),
  to: at(to),
  href: "https://example.org/",
});

/** Apply the intent the way an editor does, with the ids it asks for, and undo it. */
const applied = (document: Document, intent: HyperlinkEditorIntent) => {
  const ids = allocateEditorIntentIds(document, intent);
  const compiled = compileEditorIntent(document, {
    intent,
    mode: { type: "editing", newIds: ids.newIds },
  }).unwrap();
  const result = applyDocumentOps(document, compiled.ops).unwrap();
  expect(contractViolation(result.document)).toBeUndefined();
  expect(
    applyDocumentOps(result.document, result.inverse).unwrap().document,
  ).toEqual(document);
  const [paragraph] = result.document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : [];
};

describe("a link edit over pending reviews", () => {
  test("cuts a run's pending formatting change, giving the later half a fresh id", () => {
    const content = applied(
      documentOf([
        {
          type: "run",
          formatting: { bold: true },
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: info(5),
              previousFormatting: {},
            },
          ],
          content: [{ type: "text", text: "abcdef" }],
        },
      ]),
      linked(2, 4),
    );
    expect(content.map(({ type }) => type)).toEqual([
      "run",
      "hyperlink",
      "run",
    ]);
    const ids = JSON.stringify(content).match(/"id":\d+/gu) ?? [];
    // Three records of the one change, each under its own id.
    expect(ids).toHaveLength(3);
    expect(new Set(ids).size).toBe(3);
    expect(ids[0]).toBe('"id":5');
  });

  test("links text inside a tracked insertion, the link inside the change", () => {
    const content = applied(
      documentOf([
        { type: "run", content: [{ type: "text", text: "ab" }] },
        {
          type: "insertion",
          info: info(7),
          content: [{ type: "run", content: [{ type: "text", text: "cdef" }] }],
        },
      ]),
      linked(3, 6),
    );
    const insertions = content.filter(({ type }) => type === "insertion");
    // The change is cut where the link starts; the linked half holds the link.
    expect(insertions).toHaveLength(2);
    const second = insertions[1];
    expect(
      second?.type === "insertion" && second.content.map(({ type }) => type),
    ).toEqual(["hyperlink"]);
    expect(
      new Set(
        insertions.map((item) => item.type === "insertion" && item.info.id),
      ).size,
    ).toBe(2);
  });
});
