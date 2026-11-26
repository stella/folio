import { Result } from "better-result";
import { describe, expect, spyOn, test } from "bun:test";

import { OP_STORIES } from "@stll/docx-core/ops";

import * as inlineProjection from "../canonical/canonicalInlineProjection";
import {
  createCanonicalSession,
  projectCanonicalStory,
} from "../canonical/canonicalSession";
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

describe("a session opening a story", () => {
  test("aligns no paragraph's text until it is addressed, however long the story", () => {
    const aligned = spyOn(
      inlineProjection.InlineProjectionCache.prototype,
      "project",
    );
    try {
      const session = createCanonicalSession(story(400)).unwrap();
      const atOpen = aligned.mock.calls.length;
      // Opening may address the paragraph a caret starts in, no more.
      expect(atOpen).toBeLessThanOrEqual(1);
      const position = session.projection.positionAt({
        story: OP_STORIES.MAIN,
        blockId: idOf(200),
        offset: 3,
      });
      expect(position.isOk()).toBe(true);
      expect(aligned.mock.calls.length).toBe(atOpen + 1);
      // A paragraph is aligned once, however often it is addressed.
      session.projection.positionAt({
        story: OP_STORIES.MAIN,
        blockId: idOf(200),
        offset: 1,
      });
      expect(aligned.mock.calls.length).toBe(atOpen + 1);
    } finally {
      aligned.mockRestore();
    }
  });

  test("addresses each paragraph as a projection aligning every one does", () => {
    const document = story(40);
    const session = createCanonicalSession(document).unwrap();
    const eager = projectCanonicalStory({
      document: session.document,
      styles: undefined,
    }).unwrap();
    for (let index = 0; index < 40; index += 1) {
      const lazy = session.projection.paragraph(idOf(index));
      const full = eager.paragraph(idOf(index));
      expect(lazy?.text).toBe(full?.text);
      expect(lazy?.start).toBe(full?.start);
      for (let offset = 0; offset <= (full?.text.length ?? 0); offset += 1)
        expect(lazy?.gaps.at(offset)).toEqual(full?.gaps.at(offset));
    }
  });

  test("still refuses at open an inline construct it cannot place", () => {
    const native = inlineProjection.canonicalNativeCells;
    let calls = 0;
    const refused = spyOn(
      inlineProjection,
      "canonicalNativeCells",
    ).mockImplementation((paragraph) => {
      calls += 1;
      return calls === 2
        ? Result.err(
            new inlineProjection.InlineProjectionError({
              message:
                "Canonical editing has no position mapping for inline x.",
            }),
          )
        : native(paragraph);
    });
    try {
      const session = createCanonicalSession(story(3));
      expect(session.isErr()).toBe(true);
      if (session.isErr())
        expect(session.error.message).toBe(
          "Canonical editing has no position mapping for inline x.",
        );
    } finally {
      refused.mockRestore();
    }
  });
});
