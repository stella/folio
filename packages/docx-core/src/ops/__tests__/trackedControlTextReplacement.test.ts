/**
 * A tracked replacement of exactly a content control's text, typed just
 * inside its edges, strikes and inserts inside the control: the control stays,
 * and accepting or rejecting the change leaves it with the new or the old
 * text.
 */

import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { compileEditorIntent } from "../editorIntent";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const document: Document = {
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000002",
          content: [
            { type: "run", content: [{ type: "text", text: "Pay " }] },
            {
              type: "inlineSdt",
              properties: { sdtType: "richText", id: 7, tag: "amount" },
              content: [
                { type: "run", content: [{ type: "text", text: "100" }] },
              ],
            },
            { type: "run", content: [{ type: "text", text: " euros." }] },
          ],
        },
      ],
    },
  },
};

const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000002",
  offset,
});

const replaced = () => {
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "replaceText",
      from: at(4),
      to: at(7),
      text: "250",
      inside: { start: true, end: true },
    },
    mode: {
      type: "suggesting",
      revision: { id: 2, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
      newIds: { revision: [3, 4, 5], control: [] },
    },
  });
  if (compiled.isErr()) throw compiled.error;
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const paragraphContent = (current: Document): readonly ParagraphContent[] => {
  const [paragraph] = current.package.document.content;
  if (paragraph?.type !== "paragraph") throw new Error("Expected a paragraph");
  return paragraph.content;
};

/** The control's text, and the text outside it. */
const reading = (content: readonly ParagraphContent[]) => {
  const texts = (value: unknown): string =>
    (JSON.stringify(value).match(/"text":"[^"]*"/gu) ?? [])
      .map((match) => match.slice(8, -1))
      .join("");
  const controls = content.filter((node) => node.type === "inlineSdt");
  return {
    controls: controls.map((control) => texts(control)),
    outside: texts(content.filter((node) => node.type !== "inlineSdt")),
  };
};

test("the deletion and the insertion stand inside the control", () => {
  const content = paragraphContent(replaced());
  const control = content.find((node) => node.type === "inlineSdt");
  expect(control?.type === "inlineSdt" && control.properties.tag).toBe(
    "amount",
  );
  expect(
    control?.type === "inlineSdt" && control.content.map(({ type }) => type),
  ).toEqual(["deletion", "insertion"]);
  expect(
    content.some(
      (node) => node.type === "deletion" || node.type === "insertion",
    ),
  ).toBe(false);
});

const DECISIONS: readonly [RevisionDecision, string][] = [
  [REVISION_DECISIONS.ACCEPT, "250"],
  [REVISION_DECISIONS.REJECT, "100"],
];
for (const [decision, kept] of DECISIONS)
  test(`resolving it (${decision}) keeps the control with ${kept}`, () => {
    const resolved = applyDocumentOps(replaced(), [
      {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: [2, 5],
        decision,
      },
    ]);
    if (resolved.isErr()) throw resolved.error;
    expect(reading(paragraphContent(resolved.value.document))).toEqual({
      controls: [kept],
      outside: "Pay  euros.",
    });
  });

// An inline control lives in one paragraph: a break inside its text is
// refused, naming it; one at its edge goes beside it.
test("a paragraph break inside the control is refused, one at its edge is not", () => {
  const split = (offset: number) =>
    compileEditorIntent(document, {
      intent: {
        type: "splitParagraph",
        at: at(offset),
        newBlockId: "00000009",
      },
      mode: { type: "editing" },
    });
  const inside = split(5);
  expect(inside.isErr() && inside.error.reason).toBe(
    DOCUMENT_OP_REFUSAL_REASONS.INSIDE_INLINE_CONTROL,
  );
  expect(inside.isErr() && inside.error.message).toContain('"amount"');
  expect(split(4).isOk()).toBe(true);
  expect(split(7).isOk()).toBe(true);
});
