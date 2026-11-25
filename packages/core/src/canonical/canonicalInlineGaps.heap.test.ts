/**
 * Retained-heap budgets for the canonical session's per-paragraph projection.
 *
 * A canonical session keeps the projection of every paragraph of the story,
 * so anything the projection holds per character is held for the whole
 * document: a 500-page document is about 1.4 million characters. These tests
 * read the heap after full collections (object count and bytes, not time),
 * which is stable on a loaded machine, and fail when the projection or the
 * session starts keeping per-character objects again.
 */
import { heapStats } from "bun:jsc";
import { expect, test } from "bun:test";

import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Paragraph } from "../types/document";
import { projectCanonicalInline } from "./canonicalInlineProjection";
import { createCanonicalSession } from "./canonicalSession";

const PARAGRAPHS = 400;
const CHARACTERS_PER_PARAGRAPH = 1000;
const CHARACTERS = PARAGRAPHS * CHARACTERS_PER_PARAGRAPH;

const WORDS = ["lorem", "ipsum", "dolor", "sit", "amet", "consectetur"];
const textOf = (seed: number): string => {
  let text = "";
  for (let word = seed; text.length < CHARACTERS_PER_PARAGRAPH; word += 1)
    text += `${WORDS[word % WORDS.length]} `;
  return text.slice(0, CHARACTERS_PER_PARAGRAPH);
};

/** A body paragraph of plain text in two runs, as a long justified document holds. */
const paragraphOf = (index: number): Paragraph => {
  const text = textOf(index);
  return {
    type: "paragraph",
    paraId: (0x10000000 + index).toString(16).toUpperCase(),
    content: [
      { type: "run", content: [{ type: "text", text: text.slice(0, 600) }] },
      { type: "run", content: [{ type: "text", text: text.slice(600) }] },
    ],
  };
};

const documentOf = (paragraphs: Paragraph[]): Document =>
  ({ package: { document: { content: paragraphs } } }) as unknown as Document;

/** Heap after full collections: what is still reachable. */
const settledHeap = () => {
  Bun.gc(true);
  Bun.gc(true);
  const stats = heapStats();
  return { objects: stats.objectCount, bytes: stats.heapSize };
};

test("a paragraph's projection keeps no objects per character", () => {
  const sources = Array.from({ length: PARAGRAPHS }, (_, index) =>
    paragraphOf(index),
  );
  const nodes = sources.map((source) =>
    toProseDoc(documentOf([source])).child(0),
  );
  const kept: unknown[] = [];
  const before = settledHeap();
  for (const [index, source] of sources.entries())
    kept.push(
      projectCanonicalInline({
        source,
        paragraph: nodes[index] ?? panicMissing(),
        pairedBookmarkIds: new Set(),
      }).unwrap(),
    );
  const after = settledHeap();
  expect(kept).toHaveLength(PARAGRAPHS);
  const objectsPerCharacter = (after.objects - before.objects) / CHARACTERS;
  const bytesPerCharacter = (after.bytes - before.bytes) / CHARACTERS;
  // A gap table holds 16 bytes per offset and a handful of objects per
  // paragraph. Per-character gap and seam lists held 4 objects and over
  // 150 bytes per character.
  expect(objectsPerCharacter).toBeLessThan(0.05);
  expect(bytesPerCharacter).toBeLessThan(40);
});

test("a canonical session over a long document keeps a bounded heap per character", () => {
  const document = documentOf(
    Array.from({ length: PARAGRAPHS }, (_, index) => paragraphOf(index)),
  );
  const before = settledHeap();
  const session = createCanonicalSession(document).unwrap();
  // Every paragraph is addressable: the projection is built.
  expect(session.projection.addressAt(2).isOk()).toBe(true);
  const after = settledHeap();
  const objectsPerCharacter = (after.objects - before.objects) / CHARACTERS;
  const bytesPerCharacter = (after.bytes - before.bytes) / CHARACTERS;
  expect(objectsPerCharacter).toBeLessThan(0.25);
  expect(bytesPerCharacter).toBeLessThan(120);
  // Keep the session reachable through the reading above.
  expect(session.projection).toBeDefined();
});

const panicMissing = (): never => {
  throw new TypeError("Missing projected paragraph");
};
