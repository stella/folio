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

import { checkCoverageSample } from "../../../../../test/property-coverage";
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

/** The property's run count, which the threshold was set for. */
const COVERAGE_RUNS = 300;

/**
 * Every story of a version answers as a walk does, for the ids it holds, ids
 * earlier versions held, and an id none holds.
 */
const expectLocatorOf = (document: Document, seen: Set<string>): number => {
  let asked = 0;
  for (const story of documentStories(document)) {
    const body = storyBody(document, story);
    const walked = storyParagraphs(body);
    const ids = walked.flatMap(({ paragraph }) =>
      paragraph.paraId === undefined ? [] : [paragraph.paraId],
    );
    const earlier = [...seen];
    for (const id of ids) seen.add(id);
    for (const id of [
      ...ids,
      ...ids.map((each) => each.toLowerCase()),
      ...earlier,
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

/** `document` with `content` as its main story's blocks. */
const withMainContent = (
  document: Document,
  content: BlockContent[],
): Document => ({
  ...document,
  package: {
    ...document.package,
    document: { ...document.package.document, content },
  },
});

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
    const law = (counts: { asked: number }) =>
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
          const seen = new Set<string>();
          counts.asked += expectLocatorOf(document, seen);
          edits.forEach(({ seed, tracked, story: at, back, rotate }, index) => {
            if (back !== null) {
              document = versions[back % versions.length] ?? document;
              counts.asked += expectLocatorOf(document, seen);
              return;
            }
            if (rotate) {
              document = rotated(document);
              versions.push(document);
              counts.asked += expectLocatorOf(document, seen);
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
            counts.asked += expectLocatorOf(document, seen);
          });
        },
      );
    await assertProperty(law({ asked: 0 }), {
      id: "kept paragraph locator equals a walk of the story",
      numRuns: 300,
    });
    const counts = { asked: 0 };
    checkCoverageSample(law(counts), COVERAGE_RUNS);
    expect(counts.asked).toBeGreaterThan(1000);
  });

  test("equals a walk of the story along a chain longer than one index keeps", () => {
    // Eighty versions in a row, each deriving its index from the one before:
    // the chain settles into a root on the way.
    const [generated] = fc.sample(documentArbitrary, { seed: 7, numRuns: 1 });
    if (generated === undefined) throw new Error("No document was drawn.");
    const seen = new Set<string>();
    let document = generated;
    let asked = expectLocatorOf(document, seen);
    for (let version = 0; version < 80; version += 1) {
      // One block replaced by a copy of itself: one changed block a version.
      const content = [...document.package.document.content];
      const at = version % Math.max(1, content.length);
      const block = content[at];
      if (block !== undefined) content[at] = { ...block };
      document = withMainContent(document, content);
      asked += expectLocatorOf(document, seen);
    }
    expect(asked).toBeGreaterThan(80);
  });
});
