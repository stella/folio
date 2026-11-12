/**
 * The comment check judges each comment by its own anchors in document order.
 *
 * Paragraphs hold comment range starts, ends and references for a few ids in
 * a generated order. The check must give the verdict read straight off that
 * order: every anchor has a definition, each comment has as many starts as
 * ends, at most one of each and at most one reference, and its start comes
 * before its end.
 */
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document } from "../../model/document";
import { commentDocumentIssue } from "../comments";

setDefaultTimeout(propertyTestTimeout(15_000));

const anchorOf = fc.record({
  type: fc.constantFrom(
    "commentRangeStart" as const,
    "commentRangeEnd" as const,
    "commentReference" as const,
  ),
  id: fc.integer({ min: 0, max: 3 }),
});
type Anchor = fc.RecordValue<typeof anchorOf>;

const documentOf = (
  paragraphs: readonly (readonly Anchor[])[],
  defined: readonly number[],
): Document =>
  ({
    package: {
      document: {
        content: paragraphs.map((anchors, index) => ({
          type: "paragraph" as const,
          paraId: (0x10000000 + index).toString(16).toUpperCase(),
          content: [
            { type: "run", content: [{ type: "text", text: "a" }] },
            ...anchors.map((anchor) =>
              anchor.type === "commentReference"
                ? {
                    type: "run",
                    content: [{ type: "commentReference", id: anchor.id }],
                  }
                : { type: anchor.type, id: anchor.id },
            ),
          ],
        })),
        comments: defined.map((id) => ({ id, author: "A", content: [] })),
      },
    },
  }) as Document;

/** The verdict read off the anchors in the order they were generated. */
const expectedIssue = (
  anchors: readonly Anchor[],
  defined: ReadonlySet<number>,
): string | undefined => {
  if (anchors.some(({ id }) => !defined.has(id)))
    return "A comment anchor has no definition.";
  for (const id of [...defined].toSorted((a, b) => a - b)) {
    const at = (type: Anchor["type"]) =>
      anchors.flatMap((anchor, index) =>
        anchor.id === id && anchor.type === type ? [index] : [],
      );
    const starts = at("commentRangeStart");
    const ends = at("commentRangeEnd");
    if (
      starts.length !== ends.length ||
      starts.length > 1 ||
      at("commentReference").length > 1
    )
      return "Comment anchors are unbalanced or duplicated.";
    const [start] = starts;
    const [end] = ends;
    if (start !== undefined && end !== undefined && start > end)
      return "Comment range boundaries must be ordered in one story.";
  }
  return undefined;
};

test("the comment check gives the verdict the anchors' order gives", () => {
  assertProperty(
    fc.property(
      fc.array(fc.array(anchorOf, { maxLength: 3 }), {
        minLength: 1,
        maxLength: 5,
      }),
      fc.uniqueArray(fc.integer({ min: 0, max: 3 }), { maxLength: 4 }),
      (paragraphs, defined) => {
        const sorted = defined.toSorted((a, b) => a - b);
        expect(
          commentDocumentIssue(documentOf(paragraphs, sorted), "all"),
        ).toBe(expectedIssue(paragraphs.flat(), new Set(sorted)));
      },
    ),
    { id: "comment check verdict follows anchor order", numRuns: 500 },
  );
});
