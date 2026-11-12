/**
 * Carrying paragraph source captures over the blocks an edit replaced equals
 * carrying them over every paragraph of every story.
 *
 * Each version a session commits is turned into a target the way an edit
 * leaves one: the blocks it changed are copies that hold no captures, every
 * other block is the object the version before held. Both the carry from the
 * version before and the carry from the owners of each paragraph id must
 * give the same document, the same captures on every paragraph, and keep the
 * same blocks.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  documentStories,
  findStoryBody,
  type OpStory,
} from "@stll/docx-core/ops";

import { checkCoverageSample } from "../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { paragraphPropertySourceState } from "../docx/paragraphPropertySource";
import type { BlockContent, Document, Paragraph } from "../types/document";
import {
  applyEdit,
  editArbitrary,
  parsedSources,
} from "./__tests__/canonicalEditRuns";
import { changedWindow } from "./canonicalSeedContent";
import { createCanonicalSession } from "./canonicalSession";
import {
  ownedSourceCopies,
  preservePropertySources,
  preservePropertySourcesWholeStories,
} from "./propertySourcePreservation";

setDefaultTimeout(propertyTestTimeout(320_000));

/** A block as a new object holding no captures, its paragraphs new too. */
const uncaptured = (block: BlockContent): BlockContent => {
  if (block.type === "paragraph") return { ...block };
  if (block.type !== "table") return block;
  return {
    ...block,
    rows: block.rows.map((row) => ({
      ...row,
      cells: row.cells.map((cell) => ({
        ...cell,
        content: cell.content.map(uncaptured),
      })),
    })),
  };
};

/** The stories whose blocks differ, the way the session finds them. */
const changedStoriesOf = (previous: Document, next: Document): OpStory[] =>
  documentStories(next).filter(
    (story) =>
      findStoryBody(previous, story)?.content !==
      findStoryBody(next, story)?.content,
  );

/**
 * `next` with every block of the main story's run copied, holding no
 * captures; the other stories as `next` holds them.
 */
const targetOf = (previous: Document, next: Document): Document => {
  const after = next.package.document.content;
  const { added } = changedWindow(previous.package.document.content, after);
  const copies = new Map(added.map((block) => [block, uncaptured(block)]));
  return {
    ...next,
    package: {
      ...next.package,
      document: {
        ...next.package.document,
        content: after.map((block) => copies.get(block) ?? block),
      },
    },
  };
};

const paragraphsOf = (blocks: readonly BlockContent[]): Paragraph[] =>
  blocks.flatMap((block): Paragraph[] => {
    if (block.type === "paragraph") return [block];
    if (block.type !== "table") return [];
    return block.rows.flatMap((row) =>
      row.cells.flatMap((cell) => paragraphsOf(cell.content)),
    );
  });

/** Every story's paragraphs with their captures, by value, and which blocks are the target's. */
const outcomeOf = (document: Document, target: Document) =>
  documentStories(document).map((story) => {
    const blocks = findStoryBody(document, story)?.content ?? [];
    const kept = findStoryBody(target, story)?.content ?? [];
    return {
      blocks: JSON.parse(JSON.stringify(blocks)) as unknown,
      captures: JSON.parse(
        JSON.stringify(
          paragraphsOf(blocks).map((paragraph) =>
            paragraphPropertySourceState(paragraph),
          ),
        ),
      ) as unknown,
      kept: blocks.map((block, index) => block === kept[index]),
    };
  });

describe("carrying paragraph source captures over an edit", () => {
  test("equals carrying them over every paragraph", async () => {
    const sources = await parsedSources();
    const law = (counts: { compared: number }) =>
      fc.property(
        fc.constantFrom(...sources),
        fc.array(editArbitrary, { minLength: 1, maxLength: 12 }),
        ({ document }, edits) => {
          const session = createCanonicalSession(
            structuredClone(document),
          ).unwrap();
          let state = EditorState.create({
            schema: session.projection.doc.type.schema,
            doc: session.projection.doc,
          });
          for (const edit of edits) {
            const previous = session.document;
            state = applyEdit(session, state, edit);
            const next = session.document;
            if (next === previous) continue;
            const stories = changedStoriesOf(previous, next);
            const target = targetOf(previous, next);
            const windowed = preservePropertySources({
              target,
              source: previous,
              stories,
            });
            const whole = preservePropertySourcesWholeStories({
              target,
              source: previous,
              stories,
            });
            expect(outcomeOf(windowed, target)).toEqual(
              outcomeOf(whole, target),
            );
            // The owners of each id: every main-story paragraph of `previous`.
            const owners = new Map(
              previous.package.document.content.flatMap((block) =>
                block.type === "paragraph" && block.paraId !== undefined
                  ? [[block.paraId, block] as const]
                  : [],
              ),
            );
            const blocks = target.package.document.content;
            const run = changedWindow(
              previous.package.document.content,
              blocks,
            ).added;
            const fromRun = ownedSourceCopies(owners, run);
            const fromAll = ownedSourceCopies(owners, blocks);
            expect([...fromRun.keys()]).toEqual([...fromAll.keys()]);
            expect(
              [...fromRun.values()].map((copy) => [
                JSON.parse(JSON.stringify(copy)) as unknown,
                JSON.parse(
                  JSON.stringify(paragraphPropertySourceState(copy)),
                ) as unknown,
              ]),
            ).toEqual(
              [...fromAll.values()].map((copy) => [
                JSON.parse(JSON.stringify(copy)) as unknown,
                JSON.parse(
                  JSON.stringify(paragraphPropertySourceState(copy)),
                ) as unknown,
              ]),
            );
            counts.compared += 1;
          }
        },
      );
    await assertProperty(law({ compared: 0 }), {
      id: "property source carry over changed blocks equals the whole walk",
      numRuns: 150,
    });
    // A fixed sample compares enough edits that changed the document. Each
    // case replays up to 12 edits, so a third of the runs, at the same ratio.
    const RUNS = 50;
    const counts = { compared: 0 };
    checkCoverageSample(law(counts), RUNS);
    expect(counts.compared).toBeGreaterThan((RUNS * 2) / 3);
  });
});
