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
});
