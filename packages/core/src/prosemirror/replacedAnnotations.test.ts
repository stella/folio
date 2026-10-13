/**
 * The disposition table must stay total over the schema.
 *
 * Boundary inheritance handles every `inclusive: true` mark, so the marks a
 * replacement has to decide about are exactly the non-inclusive ones. A new
 * one that nobody classified would silently follow the text — which is how a
 * comment's range start was lost in the first place — so the omission is a
 * test failure rather than a default.
 */

import { describe, expect, test } from "bun:test";

import {
  NON_INCLUSIVE_MARK_DISPOSITION,
  hasReplacedAnnotations,
  nonInclusiveMarkNames,
  surveyReplacedAnnotations,
  resolvedCommentPoints,
} from "./replacedAnnotations";
import { schema } from "./schema";

describe("NON_INCLUSIVE_MARK_DISPOSITION", () => {
  test("classifies every non-inclusive mark the schema declares", () => {
    expect(nonInclusiveMarkNames(schema).toSorted()).toEqual(
      Object.keys(NON_INCLUSIVE_MARK_DISPOSITION).toSorted(),
    );
  });
});

const paragraphWith = (text: string, markNames: readonly string[]) =>
  schema.node("doc", null, [
    schema.node("paragraph", null, [
      schema.text(
        text,
        markNames.map((name) => {
          const type = schema.marks[name];
          if (!type) {
            throw new Error(`Unknown mark: ${name}`);
          }
          return type.create(
            name === "comment" ? { commentId: 3 } : { href: "https://example.invalid/terms" },
          );
        }),
      ),
    ]),
  ]);

describe("surveyReplacedAnnotations", () => {
  test("reports nothing for a plain span, so the replace keeps its fast path", () => {
    const doc = paragraphWith("plain prose", []);
    const annotations = surveyReplacedAnnotations(doc, 1, doc.content.size - 1);

    expect(annotations).toEqual({ carried: [], leading: [], trailing: [] });
    expect(hasReplacedAnnotations(annotations)).toBe(false);
  });

  test("reports nothing for a formatting-only span", () => {
    const doc = paragraphWith("bold prose", ["bold"]);

    expect(hasReplacedAnnotations(surveyReplacedAnnotations(doc, 1, doc.content.size - 1))).toBe(
      false,
    );
  });

  test.each(["comment", "hyperlink"] as const)("carries a %s mark", (markName) => {
    const doc = paragraphWith("marked prose", [markName]);
    const { carried } = surveyReplacedAnnotations(doc, 1, doc.content.size - 1);

    expect(carried.map((mark) => mark.type.name)).toEqual([markName]);
  });

  test("restores move-range starts before and ends after replacement text", () => {
    const start = schema.node("moveRangeBoundary", {
      marker: {
        type: "moveFromRangeStart",
        id: 5,
        name: "move-5",
        author: "Reviewer",
      },
    });
    const end = schema.node("moveRangeBoundary", {
      marker: { type: "moveFromRangeEnd", id: 5 },
    });
    const doc = schema.node("doc", null, [
      schema.node("paragraph", null, [start, schema.text("replace me"), end]),
    ]);

    const annotations = surveyReplacedAnnotations(doc, 1, doc.content.size - 1);

    expect(annotations.leading.map((node) => node.attrs["marker"])).toEqual([
      start.attrs["marker"],
    ]);
    expect(annotations.trailing.map((node) => node.attrs["marker"])).toEqual([end.attrs["marker"]]);
  });
});

// Boundary deletion can live on either adjacent physical paragraph carrier.
test.each(["del", "moveFrom"] as const)(
  "point comments distinguish retired %s boundaries from surviving text",
  (kind) => {
    const deletion = schema.marks["deletion"];
    const comment = schema.marks["comment"];
    if (!deletion || !comment) throw new Error("Expected revision and comment marks");
    for (const owner of ["own", "previous", "none"] as const) {
      for (const remaining of ["", "tail"]) {
        const boundary = { pPrMark: { kind, info: { id: 12, author: "Reviewer" } } };
        const previous = schema.node("paragraph", owner === "previous" ? boundary : null, [
          schema.text("Prefix"),
        ]);
        const paragraph = schema.node("paragraph", owner === "own" ? boundary : null, [
          schema.text("gone", [
            deletion.create({ revisionId: 11, author: "Reviewer" }),
            comment.create({ commentId: 3 }),
          ]),
          ...(remaining ? [schema.text(remaining)] : []),
        ]);
        const doc = schema.node("doc", null, [previous, paragraph]);
        const points = resolvedCommentPoints({
          doc,
          removes: (node) => deletion.isInSet(node.marks) !== undefined,
        });
        expect([...points.keys()]).toEqual(owner === "none" || remaining ? [3] : []);
      }
    }
  },
);
