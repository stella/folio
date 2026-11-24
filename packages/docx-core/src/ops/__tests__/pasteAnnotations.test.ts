/**
 * Pasted note references and comment markers: each note and comment is copied
 * under a fresh id where its reference lands, directly or as a suggestion, and
 * the copy never takes the id of a standing one.
 */
import { expect, test } from "bun:test";

import type {
  Comment,
  Document,
  Footnote,
  Paragraph,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const text = (value: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text: value }],
});

const note = (id: number, value: string): Footnote => ({
  type: "footnote",
  id,
  content: [
    {
      type: "paragraph",
      paraId: `0000F00${String(id)}`,
      content: [text(value)],
    },
  ],
});

const comment = (id: number, value: string): Comment => ({
  id,
  author: "Reviewer",
  date: "2026-01-02T00:00:00Z",
  content: [
    {
      type: "paragraph",
      paraId: `0000C00${String(id)}`,
      content: [text(value)],
    },
  ],
});

/** A body that already holds footnote 1 and comment 1. */
const destination = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [
            text("Before"),
            { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
            text(" after"),
          ],
        },
        {
          type: "paragraph",
          paraId: "00000002",
          content: [
            { type: "commentRangeStart", id: 1 },
            text("Commented"),
            { type: "commentRangeEnd", id: 1 },
            { type: "commentReference", id: 1 },
          ],
        },
      ],
      comments: [comment(1, "Standing comment")],
    },
    footnotes: [note(1, "Standing note")],
  },
});

/** Clipboard text naming footnote 1 and comment 1 of its source. */
const clipboard: Paragraph = {
  type: "paragraph",
  content: [
    { type: "commentRangeStart", id: 1 },
    text("Copied"),
    { type: "run", content: [{ type: "footnoteRef", id: 1 }] },
    { type: "commentRangeEnd", id: 1 },
    { type: "commentReference", id: 1 },
  ],
};

const SUGGESTING = {
  type: "suggesting",
  revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
  newIds: { revision: Array.from({ length: 80 }, (_, index) => 91 + index) },
} as const satisfies EditorIntentMode;

const paste = (
  mode: EditorIntentMode,
  document: Document = destination(),
): Document => {
  const at = { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 };
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "pasteBlocks",
      from: at,
      to: at,
      blocks: [clipboard],
      openStart: 1,
      openEnd: 1,
      annotations: {
        notes: [note(1, "Copied note")],
        comments: [comment(1, "Copied comment")],
      },
    },
    mode,
  });
  if (compiled.isErr()) throw compiled.error;
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const noteTexts = (document: Document) =>
  (document.package.footnotes ?? []).map(({ id, content }) => [
    id,
    JSON.stringify(content).match(/"text":"([^"]*)"/u)?.[1],
  ]);
/** Comment ids share the revision id space, so a copy's id is only fresh. */
const commentTexts = (document: Document) =>
  (document.package.document.comments ?? []).map(({ id, content }) => [
    id === 1 ? 1 : "fresh",
    JSON.stringify(content).match(/"text":"([^"]*)"/u)?.[1],
  ]);

for (const mode of [{ type: "editing" }, SUGGESTING] as const)
  test(`notes and comments are copied under fresh ids (${mode.type})`, () => {
    const pasted = paste(mode);
    expect(noteTexts(pasted)).toEqual([
      [1, "Standing note"],
      [2, "Copied note"],
    ]);
    expect(commentTexts(pasted)).toEqual([
      [1, "Standing comment"],
      ["fresh", "Copied comment"],
    ]);
    const copied = pasted.package.document.comments?.at(1)?.id;
    const second = pasted.package.document.content.at(1);
    if (second?.type !== "paragraph") throw new TypeError("Lost paragraph");
    const json = JSON.stringify(second.content);
    expect(json).toContain('"type":"footnoteRef","id":2');
    expect(json).toContain(`"type":"commentRangeStart","id":${String(copied)}`);
    expect(json).toContain(`"type":"commentReference","id":${String(copied)}`);
    expect(json).not.toContain("_folioPasteAnchor");
  });

test("a suggested paste with a note, accepted, reads as the direct paste", () => {
  const direct = paste({ type: "editing" });
  const suggested = paste(SUGGESTING);
  const ids = new Set<number>();
  JSON.stringify(suggested, (key, value: unknown) => {
    if (key === "info" && typeof value === "object" && value !== null) {
      const id: unknown = Reflect.get(value, "id");
      if (typeof id === "number") ids.add(id);
    }
    return value;
  });
  const accepted = applyDocumentOp(suggested, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...ids],
    decision: REVISION_DECISIONS.ACCEPT,
  });
  if (accepted.isErr()) throw accepted.error;
  expect(noteTexts(accepted.value.document)).toEqual(noteTexts(direct));
  expect(commentTexts(accepted.value.document)).toEqual(commentTexts(direct));
});

test("a reference naming no copied note refuses", () => {
  const document = destination();
  const at = { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 };
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "pasteBlocks",
      from: at,
      to: at,
      blocks: [clipboard],
      openStart: 1,
      openEnd: 1,
      annotations: { notes: [], comments: [comment(1, "Copied comment")] },
    },
    mode: { type: "editing" },
  });
  expect(compiled.isErr()).toBe(true);
});

test("two copied comments take two fresh ids while suggesting", () => {
  const document = destination();
  const at = { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 };
  const twoComments: Paragraph = {
    type: "paragraph",
    content: [
      { type: "commentReference", id: 1 },
      text("Both"),
      { type: "commentReference", id: 2 },
    ],
  };
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "pasteBlocks",
      from: at,
      to: at,
      blocks: [twoComments],
      openStart: 1,
      openEnd: 1,
      annotations: {
        notes: [],
        comments: [comment(1, "First copy"), comment(2, "Second copy")],
      },
    },
    mode: SUGGESTING,
  });
  if (compiled.isErr()) throw compiled.error;
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  const ids = (applied.value.document.package.document.comments ?? []).map(
    ({ id }) => id,
  );
  expect(ids).toHaveLength(3);
  expect(new Set(ids).size).toBe(3);
});

test("a destination bookmark named like a placeholder takes no copy's place", () => {
  const document = destination();
  const first = document.package.document.content.at(0);
  if (first?.type !== "paragraph") throw new TypeError("Lost paragraph");
  // Every name a placeholder could take, standing in the destination.
  first.content.push(
    { type: "bookmarkStart", id: 40, name: "_folioPasteAnchor0_0" },
    { type: "bookmarkEnd", id: 40 },
  );
  const pasted = paste({ type: "editing" }, document);
  const second = pasted.package.document.content.at(1);
  if (second?.type !== "paragraph") throw new TypeError("Lost paragraph");
  expect(JSON.stringify(second.content)).toContain(
    '"type":"footnoteRef","id":2',
  );
});
