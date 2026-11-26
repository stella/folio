import { describe, expect, spyOn, test } from "bun:test";
import { Node as PMNode } from "prosemirror-model";

import { createCanonicalSession } from "../canonical/canonicalSession";
import type { Document } from "../types/document";

const idOf = (index: number) =>
  (0x10_00_00_00 + index).toString(16).toUpperCase();

/** A story of `count` plain paragraphs, each with an id of its own. */
const story = (count: number): Document => ({
  package: {
    document: {
      content: Array.from({ length: count }, (_, index) => ({
        type: "paragraph" as const,
        paraId: idOf(index),
        content: [
          {
            type: "run" as const,
            content: [{ type: "text" as const, text: `Paragraph ${index}.` }],
          },
        ],
      })),
    },
  },
});

describe("a session opening a story without text boxes", () => {
  test("serializes none of the blocks it converts to learn where they may be reused", () => {
    const serialized = spyOn(PMNode.prototype, "toJSON");
    try {
      createCanonicalSession(story(400)).unwrap();
      // Telling whether a block holds a text-box group id once serialized
      // every converted block, 800 nodes here.
      expect(serialized.mock.calls.length).toBeLessThan(10);
    } finally {
      serialized.mockRestore();
    }
  });
});
