/**
 * `w:initials` on tracked changes: serialization + PM round-trip.
 *
 * NOTE: `w:initials` is not part of ECMA-376 `CT_TrackChange` (it is defined on
 * `w:comment`). Folio emits it as an optional attribution extension when the
 * source carries it; Word ignores unknown attributes on `w:ins`/`w:del`.
 */

import { describe, expect, test } from "bun:test";
import type { Node as PMNode } from "prosemirror-model";

import { serializeParagraph } from "../../docx/serializer/paragraphSerializer";
import type { Document, Paragraph } from "../../types/document";
import { schema } from "../schema";
import { fromProseDoc } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";

const DATE = "2026-07-17T00:00:00.000Z";

const paragraphOf = (nodes: PMNode[]): PMNode =>
  schema.nodes["doc"]!.create({}, [schema.nodes["paragraph"]!.create({}, nodes)]);

const firstParagraph = (model: Document): Paragraph => {
  const block = model.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("expected a paragraph");
  }
  return block;
};

describe("tracked-change w:initials", () => {
  test("a run insertion serializes w:initials", () => {
    const mark = schema.marks["insertion"]!.create({
      revisionId: 5,
      author: "Ann Author",
      date: DATE,
      initials: "AA",
    });
    const xml = serializeParagraph(
      firstParagraph(fromProseDoc(paragraphOf([schema.text("hi", [mark])]))),
    );
    expect(xml).toContain("<w:ins");
    expect(xml).toContain('w:initials="AA"');
  });

  test("initials round-trip through the PM bridge (model -> PM -> model)", () => {
    const model: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              content: [
                {
                  type: "insertion",
                  info: { id: 7, author: "Ann", date: DATE, initials: "AA" },
                  content: [{ type: "run", content: [{ type: "text", text: "added" }] }],
                },
              ],
            },
          ],
        },
      },
    };

    const pmDoc = toProseDoc(model);
    // The PM insertion mark carries initials.
    let markInitials: unknown;
    pmDoc.descendants((node) => {
      const mark = node.marks.find((m) => m.type.name === "insertion");
      if (mark) {
        markInitials = mark.attrs["initials"];
      }
      return undefined;
    });
    expect(markInitials).toBe("AA");

    // And it survives the trip back to the model.
    const roundTripped = firstParagraph(fromProseDoc(pmDoc));
    const insertion = roundTripped.content.find((item) => item.type === "insertion");
    expect(insertion?.type === "insertion" ? insertion.info.initials : undefined).toBe("AA");
    expect(serializeParagraph(roundTripped)).toContain('w:initials="AA"');
  });

  test("no w:initials attribute when none is set", () => {
    const mark = schema.marks["insertion"]!.create({ revisionId: 8, author: "Ann", date: DATE });
    const xml = serializeParagraph(
      firstParagraph(fromProseDoc(paragraphOf([schema.text("x", [mark])]))),
    );
    expect(xml).toContain("<w:ins");
    expect(xml).not.toContain("w:initials");
  });
});
