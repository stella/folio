/**
 * Finding a story paragraph by its id through the kept locator equals
 * finding it in a walk of every paragraph of the story.
 *
 * The locator indexes each version's top-level blocks from a version indexed
 * before it, along the blocks that differ, so every version of random runs
 * of direct and tracked edits, in every story, and back to earlier versions
 * as undo brings them, is asked for every paragraph id it holds, in both
 * spellings, and for ids it does not hold.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { BlockContent, Document } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { storyParagraphs } from "../blocks";
import { idKey } from "../ids";
import { findStoryParagraph } from "../paragraphLocator";
import { documentStories, storyBody } from "../stories";
import { OP_STORIES } from "../types";
import {
  documentArbitrary,
  opForStory,
  opSeedArbitrary,
  reviewDocumentArbitrary,
  trackedOpFor,
} from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(240_000));

/** Every story of a version answers as a walk does. */
const expectLocatorOf = (document: Document): number => {
  let asked = 0;
  for (const story of documentStories(document)) {
    const body = storyBody(document, story);
    const walked = storyParagraphs(body);
    const ids = walked.flatMap(({ paragraph }) =>
      paragraph.paraId === undefined ? [] : [paragraph.paraId],
    );
    for (const id of [
      ...ids,
      ...ids.map((each) => each.toLowerCase()),
      "7FFFFFFE",
    ]) {
      const expected = walked.find(
        ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(id),
      );
      const found = findStoryParagraph(body, id);
      expect(found).toEqual(expected);
      expect(found?.paragraph).toBe(expected?.paragraph);
      asked += 1;
    }
  }
  return asked;
};

/** `document` with its main story's blocks reordered: the first block moved last. */
const rotated = (document: Document): Document => {
  const [first, ...rest]: BlockContent[] = [
    ...document.package.document.content,
  ];
  if (first === undefined) return document;
  return {
    ...document,
    package: {
      ...document.package,
      document: { ...document.package.document, content: [...rest, first] },
    },
  };
};

describe("finding a story paragraph through the kept locator", () => {
  test("equals a walk of the story over a run of edits", async () => {
    let asked = 0;
    await assertProperty(
      fc.property(
        fc.oneof(documentArbitrary, reviewDocumentArbitrary),
        fc.array(
          fc.record({
            seed: opSeedArbitrary,
            tracked: fc.boolean(),
            story: fc.nat(),
            // Reading an earlier version again, as undo brings one back.
            back: fc.option(fc.nat(), { freq: 4 }),
            // Moving a block, which shifts every block after it.
            rotate: fc.boolean(),
          }),
          { minLength: 1, maxLength: 12 },
        ),
        (generated, edits) => {
          let document = generated;
          const versions = [document];
          asked += expectLocatorOf(document);
          edits.forEach(({ seed, tracked, story: at, back, rotate }, index) => {
            if (back !== null) {
              document = versions[back % versions.length] ?? document;
              asked += expectLocatorOf(document);
              return;
            }
            if (rotate) {
              document = rotated(document);
              versions.push(document);
              asked += expectLocatorOf(document);
              return;
            }
            const stories = documentStories(document);
            const story = stories[at % stories.length] ?? OP_STORIES.MAIN;
            const op =
              tracked && story === OP_STORIES.MAIN
                ? trackedOpFor(document, seed, index)
                : opForStory({ document, seed, story });
            const applied = applyDocumentOp(document, op);
            if (applied.isErr()) return;
            document = applied.value.document;
            versions.push(document);
            asked += expectLocatorOf(document);
          });
        },
      ),
      {
        id: "kept paragraph locator equals a walk of the story",
        numRuns: 300,
      },
    );
    expect(asked).toBeGreaterThan(1000);
  });
});
