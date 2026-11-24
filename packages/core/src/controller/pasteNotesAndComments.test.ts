/**
 * Copying text with a footnote and a comment, then pasting it: the paste
 * copies the note and the comment under fresh ids, directly or
 * as a suggestion that, accepted, leaves what the direct paste does.
 */
import { expect, test } from "bun:test";
import { TextSelection } from "prosemirror-state";

import {
  createCanonicalEditorHarness,
  resolveCanonicalHarnessDocument,
} from "../../../../test/canonicalEditorHarness";
import type { Document, Paragraph } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";

const text = (value: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text: value }],
});

const source = (): Document => {
  const document = createEmptyDocument();
  const annotated: Paragraph = {
    type: "paragraph",
    paraId: "10000001",
    content: [
      { type: "commentRangeStart", id: 1 },
      text("Noted"),
      { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
      { type: "commentRangeEnd", id: 1 },
      { type: "commentReference", id: 1 },
    ],
  };
  document.package.document.content = [
    annotated,
    { type: "paragraph", paraId: "10000002", content: [text("Target")] },
  ];
  document.package.document.comments = [
    {
      id: 1,
      author: "Reviewer",
      content: [
        { type: "paragraph", paraId: "1000C001", content: [text("A remark")] },
      ],
    },
  ];
  document.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      content: [
        { type: "paragraph", paraId: "1000F001", content: [text("A source")] },
      ],
    },
  ];
  return document;
};

const stories = (document: Document) => ({
  notes: (document.package.footnotes ?? []).map(({ id, content }) => [
    id,
    JSON.stringify(content).includes("A source"),
  ]),
  // Comment ids share the revision id space, so a copy's id is only fresh.
  comments: (document.package.document.comments ?? []).map(
    ({ id, content }) => [
      id === 1 ? 1 : "fresh",
      JSON.stringify(content).includes("A remark"),
    ],
  ),
});

const copyAndPaste = (mode: "editing" | "suggesting") => {
  const driver = createCanonicalEditorHarness(source(), mode);
  try {
    // The first paragraph's content, copied, then pasted at the second's end.
    const first = driver.state.doc.child(0);
    driver.dispatch(
      driver.state.tr.setSelection(
        TextSelection.create(driver.state.doc, 1, 1 + first.content.size),
      ),
    );
    driver.copy();
    const end = driver.state.doc.content.size - 1;
    driver.dispatch(
      driver.state.tr.setSelection(TextSelection.create(driver.state.doc, end)),
    );
    driver.pasteCopied();
    expect(driver.refusals).toEqual([]);
    return driver.snapshot();
  } finally {
    driver.dispose();
  }
};

test.each(["editing", "suggesting"] as const)(
  "a copied footnote and comment are copied under fresh ids (%s)",
  (mode) => {
    const pasted = copyAndPaste(mode);
    expect(stories(pasted)).toEqual({
      notes: [
        [1, true],
        [2, true],
      ],
      comments: [
        [1, true],
        ["fresh", true],
      ],
    });
    const copied = pasted.package.document.comments?.at(1)?.id;
    const target = JSON.stringify(pasted.package.document.content.at(1));
    expect(target).toContain('"type":"footnoteRef","id":2');
    expect(target).toContain(
      `"type":"commentReference","id":${String(copied)}`,
    );
  },
);

test("a slice like the copy, but not the clipboard the editor wrote, refuses", () => {
  const driver = createCanonicalEditorHarness(source(), "editing");
  try {
    const first = driver.state.doc.child(0);
    driver.dispatch(
      driver.state.tr.setSelection(
        TextSelection.create(driver.state.doc, 1, 1 + first.content.size),
      ),
    );
    driver.copy();
    const slice = driver.state.selection.content();
    const before = driver.state.doc;
    driver.paste(slice);
    expect(driver.refusals).toHaveLength(1);
    expect(driver.state.doc.eq(before)).toBe(true);
  } finally {
    driver.dispose();
  }
});

test("the suggested copy, accepted, leaves what the direct copy does", () => {
  const direct = copyAndPaste("editing");
  const accepted = resolveCanonicalHarnessDocument(
    copyAndPaste("suggesting"),
    "accept",
  ).model;
  expect(stories(accepted)).toEqual(stories(direct));
});
