/**
 * Inline references carry package semantics that do not affect their visible
 * text. Those facts still have to cross the editor projection: an unrelated
 * edit must not turn a named-frame link into a default link or forget where
 * inside the target document it points.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { Document, Hyperlink, NoteReferenceContent, Paragraph } from "../types/document";

const optionalString = fc.option(fc.string({ minLength: 1, maxLength: 32 }), { nil: undefined });
const optionalBoolean = fc.option(fc.boolean(), { nil: undefined });

const documentHolding = (content: Paragraph["content"]): Document => ({
  package: { document: { content: [{ type: "paragraph", content }] } },
});

const projectedParagraph = (source: Document): Paragraph => {
  const prose = toProseDoc(source);
  const cloned = prose.type.schema.nodeFromJSON(prose.toJSON());
  const block = fromProseDoc(cloned, source).package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("The editor projection lost its paragraph.");
  }
  return block;
};

const projectedHyperlink = (source: Hyperlink): Hyperlink => {
  const projected = projectedParagraph(documentHolding([source])).content.at(0);
  if (projected?.type !== "hyperlink") {
    throw new Error("The editor projection lost its hyperlink.");
  }
  return projected;
};

const projectedNoteReference = (source: NoteReferenceContent): NoteReferenceContent => {
  const projected = projectedParagraph(
    documentHolding([{ type: "run", content: [source] }]),
  ).content.at(0);
  if (projected?.type !== "run") {
    throw new Error("The editor projection lost its note-reference run.");
  }
  const reference = projected.content.at(0);
  if (reference?.type !== "footnoteRef" && reference?.type !== "endnoteRef") {
    throw new Error("The editor projection lost its note reference.");
  }
  return reference;
};

describe("inline reference metadata", () => {
  test(
    "every authored hyperlink attribute survives the editor projection",
    () => {
      fc.assert(
        fc.property(
          optionalString,
          optionalBoolean,
          optionalString,
          (target, history, docLocation) => {
            const source: Hyperlink = {
              type: "hyperlink",
              href: "https://example.test/source",
              rId: "rId7",
              tooltip: "Source link",
              ...(target === undefined ? {} : { target }),
              ...(history === undefined ? {} : { history }),
              ...(docLocation === undefined ? {} : { docLocation }),
              children: [{ type: "run", content: [{ type: "text", text: "source" }] }],
            };

            const projected = projectedHyperlink(source);

            expect(projected.target).toBe(target);
            expect(projected.history).toBe(history);
            expect(projected.docLocation).toBe(docLocation);
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(30_000),
  );

  test(
    "an explicit custom-mark decision survives on every note reference",
    () => {
      fc.assert(
        fc.property(
          fc.constantFrom("footnoteRef", "endnoteRef"),
          optionalBoolean,
          (type, customMarkFollows) => {
            const source: NoteReferenceContent = {
              type,
              id: 7,
              ...(customMarkFollows === undefined ? {} : { customMarkFollows }),
            };

            expect(projectedNoteReference(source)).toEqual(source);
          },
        ),
        propertyConfig({ numRuns: 100 }),
      );
    },
    propertyTestTimeout(30_000),
  );
});
