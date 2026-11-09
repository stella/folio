import { describe, expect, test } from "bun:test";

import type { Document, ParagraphContent, Run } from "../../model/document";
import { applyDocumentOps } from "../apply";
import { allocateEditorIntentIds, compileEditorIntent } from "../editorIntent";
import { identityKeysIn } from "../ids";
import { paragraphLength } from "../offsets";
import { OP_STORIES } from "../types";

/**
 * Typing over a selection beside, or across, the author's own earlier
 * deletions continues one of them. Whichever record the struck words join
 * and whatever dates the records carry, the edit applies, and every record
 * keeps an id of its own.
 */

const DATES = [
  "2026-01-01T00:00:00.001Z",
  "2026-01-01T00:00:00.002Z",
  "2026-01-01T00:00:00.003Z",
] as const;

const run = (text: string): Run => ({
  type: "run",
  formatting: { bold: true },
  content: [{ type: "text", text }],
});

const deletion = (
  id: number,
  date: string,
  text: string,
): ParagraphContent => ({
  type: "deletion",
  info: { id, author: "Author", date },
  content: [run(text)],
});

const SHAPES = {
  "two deletions, then text": (first: string, second: string) => [
    deletion(1, first, "A"),
    deletion(2, second, "😀"),
    run("B"),
  ],
  "text between two deletions": (first: string, second: string) => [
    deletion(1, first, "A"),
    run("B"),
    deletion(2, second, "😀"),
  ],
  "text, then two deletions": (first: string, second: string) => [
    run("B"),
    deletion(1, first, "A"),
    deletion(2, second, "😀"),
  ],
} as const;

const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "12345678", content },
        { type: "paragraph", paraId: "23456789", content: [] },
      ],
    },
  },
});

/** Offsets in the paragraph that do not fall inside a surrogate pair. */
const offsetsOf = (content: ParagraphContent[]): number[] => {
  const text = content
    .flatMap((item) => {
      if (item.type === "run") return [item];
      return item.type === "deletion" ? item.content : [];
    })
    .flatMap((item) => (item.type === "run" ? item.content : []))
    .map((item) => (item.type === "text" ? item.text : ""))
    .join("");
  const offsets = [0];
  let offset = 0;
  for (const character of text) {
    offset += character.length;
    offsets.push(offset);
  }
  return offsets;
};

describe("typing over a selection beside one's own deletions", () => {
  for (const [shape, contentOf] of Object.entries(SHAPES))
    test(`${shape}: every record dating and selection applies with distinct ids`, () => {
      let edits = 0;
      for (const first of DATES)
        for (const second of DATES)
          for (const now of DATES) {
            const content = contentOf(first, second);
            const document = documentOf(content);
            const paragraph = document.package.document.content[0];
            if (paragraph?.type !== "paragraph")
              throw new TypeError("The paragraph is absent.");
            const offsets = offsetsOf(content);
            expect(offsets.at(-1)).toBe(paragraphLength(paragraph));
            for (const [index, from] of offsets.entries())
              for (const to of offsets.slice(index + 1)) {
                const at = (offset: number) => ({
                  story: OP_STORIES.MAIN,
                  blockId: "12345678",
                  offset,
                });
                const intent = {
                  type: "replaceText",
                  from: at(from),
                  to: at(to),
                  text: "é",
                } as const;
                const ids = allocateEditorIntentIds(document, intent);
                const label = JSON.stringify({ first, second, now, from, to });
                const compiled = compileEditorIntent(document, {
                  intent,
                  mode: {
                    type: "suggesting",
                    revision: {
                      id: ids.revisionId,
                      author: "Author",
                      date: now,
                    },
                    newIds: ids.newIds,
                  },
                });
                if (compiled.isErr())
                  throw new Error(`${label}: ${compiled.error.message}`);
                const applied = applyDocumentOps(document, compiled.value.ops);
                if (applied.isErr())
                  throw new Error(`${label}: ${applied.error.message}`);
                const keys = identityKeysIn(
                  applied.value.document.package.document.content,
                );
                expect([label, new Set(keys).size]).toEqual([
                  label,
                  keys.length,
                ]);
                edits += 1;
              }
          }
      expect(edits).toBeGreaterThan(0);
    });
});
