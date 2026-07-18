/**
 * `w:initials` on tracked changes: standards-clean serialization + PM round-trip.
 *
 * `w:initials` is NOT part of ECMA-376 `CT_TrackChange` (it is defined on
 * `w:comment`). Folio therefore NEVER emits it on `w:ins`/`w:del`/`w:*PrChange`
 * or the table row/cell markers — output stays schema-strict for legal-document
 * validators. Initials are still carried in-model and on the ProseMirror marks
 * for UI attribution (hover, accept authoring), and the parser stays tolerant of
 * an `initials` attribute if some external document supplies one.
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
  test("initials are NEVER emitted on a tracked-change element (schema-strict guard)", () => {
    const mark = schema.marks["insertion"]!.create({
      revisionId: 5,
      author: "Ann Author",
      date: DATE,
      initials: "AA",
    });
    const xml = serializeParagraph(
      firstParagraph(fromProseDoc(paragraphOf([schema.text("hi", [mark])]))),
    );
    // The tracked change still serializes — just without the non-standard attr.
    expect(xml).toContain("<w:ins");
    expect(xml).toContain('w:author="Ann Author"');
    expect(xml).not.toContain("w:initials");
  });

  test("initials survive the PM bridge (model -> PM -> model) for UI attribution", () => {
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
    // The PM insertion mark carries initials (drives hover / accept authoring).
    let markInitials: unknown;
    pmDoc.descendants((node) => {
      const mark = node.marks.find((m) => m.type.name === "insertion");
      if (mark) {
        markInitials = mark.attrs["initials"];
      }
      return undefined;
    });
    expect(markInitials).toBe("AA");

    // And it survives the trip back to the model — but is NOT written to XML.
    const roundTripped = firstParagraph(fromProseDoc(pmDoc));
    const insertion = roundTripped.content.find((item) => item.type === "insertion");
    expect(insertion?.type === "insertion" ? insertion.info.initials : undefined).toBe("AA");
    expect(serializeParagraph(roundTripped)).not.toContain("w:initials");
  });
});
