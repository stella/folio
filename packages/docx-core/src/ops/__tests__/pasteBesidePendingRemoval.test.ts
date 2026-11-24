/**
 * A suggested paste beside content pending removal, accepted, reads as the
 * direct paste over that content once removed: the struck content keeps no
 * paragraph of its own (a cleared table cell pasted into while suggesting).
 */
import { expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { compileEditorIntent, type EditorIntentMode } from "../editorIntent";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const struck = (text: string, id = 5): ParagraphContent => ({
  type: "deletion",
  info: { id, author: "A", date: "2026-01-01T00:00:00Z" },
  content: [run(text)],
});

const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "00000001", content },
        { type: "paragraph", paraId: "00000002", content: [run("next")] },
      ],
    },
  },
});

const pasted: Paragraph[] = ["Alpha", "Nested", "Omega"].map((text, index) => ({
  type: "paragraph",
  paraId: `0000001${String(index)}`,
  content: [run(text)],
}));

const SUGGESTING = {
  type: "suggesting",
  revision: { id: 90, author: "R", date: "2026-01-02T00:00:00Z" },
  newIds: { revision: [91, 92, 93, 94, 95, 96, 97, 98] },
} as const satisfies EditorIntentMode;

const paste = (
  document: Document,
  offset: number,
  mode: EditorIntentMode,
): Document => {
  const at = { story: OP_STORIES.MAIN, blockId: "00000001", offset };
  const compiled = compileEditorIntent(document, {
    intent: {
      type: "replaceFragment",
      from: at,
      to: at,
      openStart: 0,
      openEnd: 0,
      paragraphs: structuredClone(pasted),
    },
    mode,
  });
  if (compiled.isErr()) throw compiled.error;
  const applied = applyDocumentOps(document, compiled.value.ops);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const revisionIds = (document: Document): number[] => {
  const ids = new Set<number>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const child of value) visit(child);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    const info: unknown = Reflect.get(value, "info");
    if (typeof info === "object" && info !== null) {
      const id: unknown = Reflect.get(info, "id");
      if (typeof id === "number") ids.add(id);
    }
    for (const child of Object.values(value)) visit(child);
  };
  visit(document.package.document.content);
  return [...ids];
};

const resolved = (
  document: Document,
  decision: (typeof REVISION_DECISIONS)[keyof typeof REVISION_DECISIONS],
): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: revisionIds(document),
    decision,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const texts = (document: Document): string[] =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? block.content
          .flatMap((item) =>
            item.type === "run"
              ? item.content.flatMap((piece) =>
                  piece.type === "text" ? [piece.text] : [],
                )
              : [],
          )
          .join("")
      : block.type,
  );

for (const { name, content, offset } of [
  { name: "after the caret", content: [struck("gone")], offset: 0 },
  { name: "before the caret", content: [struck("gone")], offset: 4 },
  {
    name: "on both sides of the caret",
    content: [struck("go"), struck("ne", 6)],
    offset: 2,
  },
]) {
  test(`content pending removal ${name} keeps no paragraph once accepted`, () => {
    const direct = paste(documentOf([]), 0, { type: "editing" });
    const suggested = paste(documentOf(content), offset, SUGGESTING);
    expect(texts(resolved(suggested, REVISION_DECISIONS.ACCEPT))).toEqual(
      texts(direct),
    );
    // Rejected, the paste is gone and the struck content stands again.
    expect(texts(resolved(suggested, REVISION_DECISIONS.REJECT))).toEqual(
      texts(resolved(documentOf(content), REVISION_DECISIONS.REJECT)),
    );
  });
}

test("live content beside the caret keeps its paragraph", () => {
  const suggested = paste(documentOf([run("kept")]), 4, SUGGESTING);
  expect(texts(resolved(suggested, REVISION_DECISIONS.ACCEPT))).toEqual([
    "kept",
    "Alpha",
    "Nested",
    "Omega",
    "next",
  ]);
});
