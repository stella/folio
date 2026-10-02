/**
 * A save writes back every `LISTNUM` field of a numbered paragraph, the ones
 * the reader folded into the list marker and the ones on the line, and what
 * the reader shows for the paragraph does not change.
 *
 * Two oracles that share nothing. The saved markup is read without the parser
 * and compared with the markup the paragraph was authored from: instruction,
 * field characters, cached result, the tab, the range markers between them and
 * the text around them, in order. The reopened document is compared with the
 * one first opened: which items the reader folded, the marker each paragraph
 * carries and is laid out with, and that the fields it shows are the fields
 * its markup holds.
 *
 * Each case goes out three ways, because each takes a different route through
 * the save: no paragraph of its own edited, text typed into every paragraph,
 * and the whole part rewritten.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { foldedListNumberPlugin } from "../prosemirror/foldedListNumber";
import type { Document } from "../types/document";
import {
  bodyParagraphs,
  cachedDisplay,
  documentXmlOf,
  expectedTokens,
  fieldResultsInFile,
  fieldResultsShown,
  type FieldSpec,
  foldedFieldsOf,
  contentShapes,
  type GapMarker,
  inlineTokens,
  layoutMarkers,
  listNumberFieldDocx,
  MARKER_KINDS,
  modelMarkers,
  openDocx,
  paragraphMarkupOf,
  type ParagraphSpec,
  PLAIN_PARAGRAPH_ID,
  saveDocx,
  typedInto,
  typeInto,
  withSettledTail,
} from "./__tests__/listNumberFieldFixture";
import { repackDocx } from "./rezip";
import { serializeParagraph } from "./serializer/paragraphSerializer";

setDefaultTimeout(propertyTestTimeout(30_000));

const GAP_MARKERS = [
  "bookmark",
  "bookmarkStart",
  "comment",
] as const satisfies readonly GapMarker[];
const RESULT_FORMATTING = ["plain", "bold", "symbol"] as const;

const fieldArbitrary: fc.Arbitrary<FieldSpec> = fc.record({
  instruction: fc.constantFrom(
    " LISTNUM ",
    "LISTNUM",
    " LISTNUM  LegalDefault \\l 3 ",
    " listnum NumberDefault \\s 2 ",
  ),
  // A literal percent sign among them; the fixed example below adds a tab
  // inside the cached result.
  result: fc.constantFrom("(a)", "(ii)", "50%", "%"),
  formatting: fc.constantFrom(...RESULT_FORMATTING),
  before: fc.constantFrom("", "", "and ", "x"),
  gap: fc.uniqueArray(fc.constantFrom(...GAP_MARKERS), { maxLength: 3 }),
  tab: fc.boolean(),
});

/**
 * A comment opens after the last field only, as the last marker ahead of its
 * tab: its range runs to the paragraph's end, and what else may stand inside
 * an open range is the editor's to order, not this save's.
 */
const foldable = (fields: FieldSpec[]): FieldSpec[] =>
  fields.map((field, index): FieldSpec => {
    const last = index === fields.length - 1;
    const bookmarks = field.gap.filter((marker) => marker !== "comment");
    const commented = last && field.gap.includes("comment");
    return { ...field, gap: commented ? [...bookmarks, "comment"] : bookmarks };
  });

const paragraphsArbitrary: fc.Arbitrary<ParagraphSpec[]> = fc
  .array(
    fc.record({
      marker: fc.constantFrom(...MARKER_KINDS),
      fields: fc.array(fieldArbitrary, { minLength: 1, maxLength: 3 }).map(foldable),
      body: fc.constantFrom("Body", "Tail text", "100% of it"),
    }),
    { minLength: 1, maxLength: 2 },
  )
  .map((paragraphs) =>
    paragraphs.map((paragraph, index) => ({ ...paragraph, paraId: `2000000${index + 1}` })),
  );

const TYPED = "!";

const foldedKinds = (model: Document): string[][] =>
  contentShapes(model).map((shape) => shape.filter((kind) => kind.startsWith("folded:")));

const expectSaved = async (
  saved: ArrayBuffer,
  original: Document,
  specs: readonly ParagraphSpec[],
): Promise<void> => {
  const xml = await documentXmlOf(saved);
  for (const spec of specs) {
    expect(withSettledTail(inlineTokens(paragraphMarkupOf(xml, spec.paraId)))).toEqual(
      withSettledTail(expectedTokens(spec)),
    );
  }

  const reopened = await openDocx(saved);
  expect(foldedKinds(reopened)).toEqual(foldedKinds(original));
  expect(modelMarkers(reopened)).toEqual(modelMarkers(original));
  expect(layoutMarkers(reopened)).toEqual(layoutMarkers(original));
  for (const paragraph of bodyParagraphs(reopened)) {
    expect(fieldResultsShown(paragraph)).toBe(
      fieldResultsInFile(inlineTokens(serializeParagraph(paragraph))),
    );
  }
};

/** What the reader owes each paragraph before any save is looked at. */
const expectFolded = (parsed: Document, specs: readonly ParagraphSpec[]): void => {
  const paragraphs = bodyParagraphs(parsed);
  for (const [index, spec] of specs.entries()) {
    const paragraph = paragraphs.at(index);
    if (!paragraph) {
      throw new Error(`The document has no paragraph ${spec.paraId}`);
    }
    // Only the fields that open the paragraph are behind the marker.
    const folded = foldedFieldsOf(spec);
    expect(foldedKinds(parsed).at(index)).toEqual(
      folded.flatMap(({ tab }) => (tab ? ["folded:field", "folded:tab"] : ["folded:field"])),
    );
    const marker = paragraph.listRendering?.marker ?? "";
    if (folded.length > 0) {
      expect(marker.endsWith(`\t${cachedDisplay(folded)}`)).toBe(true);
    } else if (spec.marker !== "symbol") {
      expect(marker.includes("\t")).toBe(false);
    }
    expect(fieldResultsShown(paragraph)).toBe(cachedDisplay(spec.fields));
  }
};

describe("saving paragraphs whose list markers hold LISTNUM fields", () => {
  test(
    "every field goes back where it stood and the markers stay as they were",
    async () => {
      await assertProperty(
        fc.asyncProperty(paragraphsArbitrary, async (specs) => {
          const buffer = await listNumberFieldDocx(specs);
          const parsed = await openDocx(buffer);
          expectFolded(parsed, specs);
          const opened = EditorState.create({
            doc: toProseDoc(parsed),
            plugins: [foldedListNumberPlugin()],
          });

          // No paragraph of its own edited.
          const beside = typeInto(opened, PLAIN_PARAGRAPH_ID, "Plain.", TYPED);
          await expectSaved(
            await saveDocx(fromProseDoc(beside.doc, parsed), buffer, [PLAIN_PARAGRAPH_ID]),
            parsed,
            specs,
          );

          // Text typed into every paragraph, after its fields.
          let edited = opened;
          for (const spec of specs) {
            edited = typeInto(edited, spec.paraId, spec.body, TYPED);
          }
          await expectSaved(
            await saveDocx(
              fromProseDoc(edited.doc, parsed),
              buffer,
              specs.map(({ paraId }) => paraId),
            ),
            parsed,
            specs.map((spec) => ({ ...spec, body: typedInto(spec.body, TYPED) })),
          );

          // The whole part rewritten.
          await expectSaved(
            await repackDocx(fromProseDoc(opened.doc, parsed), { updateModifiedDate: false }),
            parsed,
            specs,
          );
        }),
        {
          numRuns: 30,
          examples: [
            [
              [
                {
                  paraId: "20000001",
                  marker: "percent",
                  fields: [
                    {
                      instruction: " LISTNUM ",
                      result: "a\tb",
                      formatting: "symbol",
                      before: "",
                      gap: ["bookmark"],
                      tab: true,
                    },
                    {
                      instruction: "LISTNUM",
                      result: "50%",
                      formatting: "bold",
                      before: "",
                      gap: ["bookmarkStart"],
                      tab: true,
                    },
                    {
                      instruction: " LISTNUM  LegalDefault \\l 3 ",
                      result: "(a)",
                      formatting: "symbol",
                      before: "x",
                      gap: ["comment"],
                      tab: false,
                    },
                  ],
                  body: "Body",
                },
                {
                  paraId: "20000002",
                  marker: "symbol",
                  fields: [
                    {
                      instruction: " listnum NumberDefault \\s 2 ",
                      result: "(ii)",
                      formatting: "plain",
                      before: "",
                      gap: [],
                      tab: true,
                    },
                  ],
                  body: "Tail text",
                },
              ],
            ],
          ],
        },
      );
    },
    propertyTestTimeout(180_000),
  ); // Each case zips a package, then saves and reopens it three times.
});
