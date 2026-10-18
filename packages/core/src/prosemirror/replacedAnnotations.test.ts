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
import { EditorState } from "prosemirror-state";
import { acceptAIEditRevision, acceptChange } from "./commands/comments";
import { readRangeAnchorAttrs } from "./rangeAnchorAttrs";

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
  "ranged comments are removed with all covered content across %s boundaries",
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
          retiresStructure: () => true,
        });
        expect([...points.keys()]).toEqual([3]);
      }
    }
  },
);

// Bulk-only fixtures missed independently pending structural revisions.
test.each(["paragraph", "precedingParagraph", "tableRow", "tableCell", "tableHeader"] as const)(
  "accepting text alone removes a ranged comment inside pending %s deletion",
  (owner) => {
    const deletion = schema.marks["deletion"];
    const comment = schema.marks["comment"];
    if (!deletion || !comment) throw new Error("Expected revision and comment marks");
    for (const resolution of ["revision", "range"] as const) {
      const boundary = { kind: "del" as const, info: { id: 12, author: "Reviewer" } };
      const membership = { revisionId: 12, author: "Reviewer" };
      const paragraph = schema.node(
        "paragraph",
        owner === "paragraph" ? { pPrMark: boundary } : null,
        [
          schema.text("gone", [
            deletion.create({ revisionId: 11, author: "Reviewer" }),
            comment.create({ commentId: 3 }),
          ]),
        ],
      );
      const prefix = schema.node(
        "paragraph",
        owner === "precedingParagraph" ? { pPrMark: boundary } : null,
        [schema.text("Prefix")],
      );
      const container =
        owner === "paragraph" || owner === "precedingParagraph"
          ? paragraph
          : schema.node("table", null, [
              schema.node("tableRow", owner === "tableRow" ? { trDel: membership } : null, [
                schema.node(
                  owner === "tableHeader" ? "tableHeader" : "tableCell",
                  owner === "tableCell" || owner === "tableHeader"
                    ? { cellMarker: { kind: "del", info: membership } }
                    : null,
                  [paragraph],
                ),
              ]),
            ]);
      const doc = schema.node("doc", null, [prefix, container]);
      let state = EditorState.create({ schema, doc });
      let textPosition = -1;
      doc.descendants((node, position) => {
        if (node.isText && deletion.isInSet(node.marks)) textPosition = position;
      });
      expect(textPosition).toBeGreaterThan(0);
      const command =
        resolution === "revision"
          ? acceptAIEditRevision(11)
          : acceptChange(textPosition, textPosition + 4);
      expect(
        command(state, (transaction) => {
          state = state.apply(transaction);
        }),
      ).toBe(true);
      const points: number[] = [];
      state.doc.descendants((node) => {
        if (node.type.name !== "rangeAnchor") return;
        const attrs = readRangeAnchorAttrs(node);
        if (attrs.ok && attrs.value.start.type === "commentRangeStart")
          points.push(attrs.value.start.id);
      });
      expect(points).toEqual([]);
      expect(state.doc.textContent).toBe("Prefix");
      const retained = state.doc.child(1);
      if (owner === "paragraph") expect(retained.attrs["pPrMark"]).toEqual(boundary);
      if (owner === "precedingParagraph")
        expect(state.doc.child(0).attrs["pPrMark"]).toEqual(boundary);
      if (owner === "tableRow") expect(retained.firstChild?.attrs["trDel"]).toEqual(membership);
      if (owner === "tableCell" || owner === "tableHeader")
        expect(retained.firstChild?.firstChild?.attrs["cellMarker"]).toEqual({
          kind: "del",
          info: membership,
        });
      expect(() => state.doc.check()).not.toThrow();
    }
  },
);
