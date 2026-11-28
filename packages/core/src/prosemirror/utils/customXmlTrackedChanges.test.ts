/**
 * A custom XML element's own record on the review surface: listed as a
 * change of its own, or carried by the card of the content it was suggested
 * with, so accepting or rejecting that card resolves it too.
 */
import { describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import type {
  BlockCustomXml,
  Document,
  Paragraph,
  TrackedChangeInfo,
} from "../../types/document";
import { toProseDoc } from "../conversion/toProseDoc";
import { extractTrackedChanges } from "./extractTrackedChanges";

const SUGGESTED: TrackedChangeInfo = {
  id: 0,
  author: "Reviewer",
  date: "2026-02-01T00:00:00Z",
};

const inserted = (text: string, id: number): Paragraph => ({
  type: "paragraph",
  paraId: `0000000${String(id)}`,
  content: [
    {
      type: "insertion",
      info: { ...SUGGESTED, id },
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

const deleted = (text: string, id: number): Paragraph => ({
  type: "paragraph",
  paraId: `0000000${String(id)}`,
  content: [
    {
      type: "deletion",
      info: { ...SUGGESTED, id },
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

const element = (
  info: TrackedChangeInfo,
  content: Paragraph[],
): BlockCustomXml => ({
  type: "blockCustomXml",
  openingXml: '<w:customXml w:element="clause">',
  closingXml: "</w:customXml>",
  content,
  insertion: { info },
});

const entriesOf = (content: Document["package"]["document"]["content"]) =>
  extractTrackedChanges(
    EditorState.create({
      doc: toProseDoc({ package: { document: { content } } }),
    }),
  ).entries;

describe("a custom XML element's own record on the review surface", () => {
  test("rides on the card of the content suggested with it", () => {
    const entries = entriesOf([
      element({ ...SUGGESTED, id: 9 }, [inserted("Clause", 1)]),
      { type: "paragraph", paraId: "00000002", content: [] },
    ]);
    expect(entries.map(({ type }) => type)).toEqual(["insertion"]);
    const [card] = entries;
    if (card === undefined) throw new TypeError("No card.");
    expect([card.revisionId, ...(card.coalescedRevisionIds ?? [])]).toContain(
      9,
    );
  });

  test("is a change of its own beside another author's content", () => {
    const entries = entriesOf([
      element({ id: 9, author: "Tagger", date: "2026-03-01T00:00:00Z" }, [
        inserted("Clause", 1),
      ]),
      { type: "paragraph", paraId: "00000002", content: [] },
    ]);
    expect(
      entries.map(({ type, revisionId, author }) => ({
        type,
        revisionId,
        author,
      })),
    ).toEqual([
      { type: "customXmlInserted", revisionId: 9, author: "Tagger" },
      { type: "insertion", revisionId: 1, author: "Reviewer" },
    ]);
  });

  test("stays separate from unrelated content with the same metadata", () => {
    const entries = entriesOf([
      element({ ...SUGGESTED, id: 9 }, [
        { type: "paragraph", paraId: "00000001", content: [] },
      ]),
      inserted("Outside", 1),
    ]);
    expect(
      entries.map(({ type, revisionId }) => ({ type, revisionId })),
    ).toEqual([
      { type: "customXmlInserted", revisionId: 9 },
      { type: "insertion", revisionId: 1 },
    ]);
  });

  test("stays separate from an opposite-direction change inside it", () => {
    const entries = entriesOf([
      element({ ...SUGGESTED, id: 9 }, [deleted("Removed", 1)]),
      { type: "paragraph", paraId: "00000002", content: [] },
    ]);
    expect(
      entries.map(({ type, revisionId }) => ({ type, revisionId })),
    ).toEqual([
      { type: "customXmlInserted", revisionId: 9 },
      { type: "deletion", revisionId: 1 },
    ]);
  });
});
