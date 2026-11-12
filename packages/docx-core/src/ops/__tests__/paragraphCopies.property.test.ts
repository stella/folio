/**
 * Paragraph copies that hold the same fields take their paragraphs' places
 * and keep the document's standing against the seed contract.
 *
 * A host replaces a paragraph by a copy when facts it keeps outside the
 * paragraph's fields change. Generated documents get copies of chosen story
 * paragraphs, at any depth: each copy must stand where its paragraph stood,
 * every other block must be the same object, and the result must be known to
 * meet the contract exactly when the document was, with the contract checked
 * afresh agreeing. A copy that changes a field, or names a paragraph the
 * story does not hold, is refused.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Paragraph } from "../../model/document";
import { storyParagraphs } from "../blocks";
import {
  contractViolation,
  knownToMeetContract,
  validateOpsDocument,
} from "../contract";
import { withParagraphCopies } from "../paragraphCopies";
import { documentStories, sameStory, storyBody } from "../stories";
import { OP_STORIES } from "../types";
import { documentArbitrary } from "./documentArbitraries";

setDefaultTimeout(propertyTestTimeout(120_000));

const kept = Symbol("kept outside the fields");

const copyOf = (paragraph: Paragraph): Paragraph => {
  const copy: Paragraph = { ...paragraph };
  Object.defineProperty(copy, kept, { value: {}, enumerable: true });
  return copy;
};

describe("paragraph copies", () => {
  test("stand in their paragraphs' places and keep the contract", () => {
    let copied = 0;
    let unchecked = 0;
    let viewed = 0;
    assertProperty(
      fc.property(
        documentArbitrary,
        fc.nat(),
        fc.array(fc.nat(), { minLength: 1, maxLength: 4 }),
        fc.boolean(),
        (generated, storyChoice, choices, checked) => {
          // A document checked once, or the same content never checked.
          const document: Document = checked
            ? generated
            : { ...generated, package: { ...generated.package } };
          if (checked) expect(validateOpsDocument(document).isOk()).toBe(true);
          const stories = documentStories(document);
          const story = stories[storyChoice % stories.length] ?? stories[0];
          if (story === undefined) return;
          const located = storyParagraphs(storyBody(document, story));
          if (located.length === 0) return;
          const copies = new Map<Paragraph, Paragraph>();
          for (const choice of choices) {
            const paragraph = located[choice % located.length]?.paragraph;
            if (paragraph !== undefined && !copies.has(paragraph))
              copies.set(paragraph, copyOf(paragraph));
          }
          const result = withParagraphCopies({ document, story, copies });
          const after = storyParagraphs(storyBody(result, story));
          expect(after.length).toBe(located.length);
          for (const [index, { paragraph, list }] of located.entries()) {
            const now = after[index];
            expect(now?.list).toEqual(list);
            expect(now?.paragraph).toBe(copies.get(paragraph) ?? paragraph);
          }
          for (const other of documentStories(document))
            if (!sameStory(other, story))
              expect(storyBody(result, other).content).toBe(
                storyBody(document, other).content,
              );
          // A section view holds the body's new block wherever it held the old one.
          const body = storyBody(document, story);
          const bodyAfter = storyBody(result, story);
          const views = body.sections ?? [];
          const viewsAfter = bodyAfter.sections ?? [];
          expect(viewsAfter.length).toBe(views.length);
          for (const [index, view] of views.entries()) {
            const blocks = viewsAfter[index]?.content ?? [];
            for (const [at, block] of view.content.entries()) {
              const held = body.content.indexOf(block);
              expect(blocks[at]).toBe(
                held < 0 ? block : bodyAfter.content[held],
              );
            }
            if (
              view.content.some(
                (block) =>
                  body.content.includes(block) &&
                  bodyAfter.content[body.content.indexOf(block)] !== block,
              )
            )
              viewed += 1;
          }
          expect(knownToMeetContract(result)).toBe(
            knownToMeetContract(document),
          );
          if (checked) expect(contractViolation(result)).toBeUndefined();
          else unchecked += 1;
          copied += copies.size;
        },
      ),
      { id: "paragraph copies keep places and the contract", numRuns: 300 },
    );
    expect(copied).toBeGreaterThan(300);
    expect(unchecked).toBeGreaterThan(50);
    // Section views held copied paragraphs, so their check ran.
    expect(viewed).toBeGreaterThan(20);
  });

  test("a broken unchecked document stays refused, and false copies are refused", () => {
    let refused = 0;
    assertProperty(
      fc.property(
        documentArbitrary,
        fc.nat(),
        fc.nat(),
        (generated, at, other) => {
          const blocks = generated.package.document.content;
          const indexes = blocks.flatMap((block, index) =>
            block.type === "paragraph" ? [index] : [],
          );
          if (indexes.length < 2) return;
          const index = indexes[at % indexes.length] ?? 0;
          const block = blocks[index];
          if (block?.type !== "paragraph") return;
          // The document with one paragraph's identity removed, never checked.
          const anonymous: Paragraph = { ...block };
          delete anonymous.paraId;
          const broken: Document = {
            ...generated,
            package: {
              ...generated.package,
              document: {
                ...generated.package.document,
                content: blocks.map((each) =>
                  each === block ? anonymous : each,
                ),
              },
            },
          };
          const otherIndex =
            indexes.filter((each) => each !== index)[
              other % (indexes.length - 1)
            ] ?? 0;
          const copied = blocks[otherIndex];
          if (copied?.type !== "paragraph") return;
          const result = withParagraphCopies({
            document: broken,
            story: OP_STORIES.MAIN,
            copies: new Map([[copied, copyOf(copied)]]),
          });
          expect(knownToMeetContract(result)).toBe(false);
          expect(validateOpsDocument(result).isErr()).toBe(true);
          refused += 1;
          // A copy that changes a field is no copy.
          expect(() =>
            withParagraphCopies({
              document: generated,
              story: OP_STORIES.MAIN,
              copies: new Map([[block, anonymous]]),
            }),
          ).toThrow("must hold the fields");
          // Nor is one that holds another value under the same fields.
          expect(() =>
            withParagraphCopies({
              document: generated,
              story: OP_STORIES.MAIN,
              copies: new Map([[block, { ...block, content: [] }]]),
            }),
          ).toThrow("must hold the fields");
          // A paragraph the story does not hold has no place to take.
          const stranger: Paragraph = { ...block };
          expect(() =>
            withParagraphCopies({
              document: generated,
              story: OP_STORIES.MAIN,
              copies: new Map([[stranger, copyOf(stranger)]]),
            }),
          ).toThrow("must replace a paragraph of its story");
        },
      ),
      { id: "paragraph copies of a broken document", numRuns: 100 },
    );
    expect(refused).toBeGreaterThan(20);
  });
});
