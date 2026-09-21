/**
 * What a `<w:numPr>` states has to survive the two round trips Folio owns: the
 * parse/serialize pair over OOXML, and the model/editor pair over ProseMirror
 * attrs. Both used to carry two optional slots with a reserved id inside them,
 * so "states no level", "states level zero" and "states no numbering" were one
 * shape apart and each tier decided for itself which it had.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  type ParagraphNumberingOverride,
  paragraphNumberingFromSlots,
  resolveParagraphNumbering,
} from "@stll/docx-core/model";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { resolveListState } from "../prosemirror/listState";
import type { Document, NumberFormat, NumberingDefinitions } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { createNumberingMap } from "./numberingParser";
import { readParagraphNumbering } from "./numberingReference";
import { serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { findChild, parseXmlDocument } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** One authored `<w:pPr>`, as the parser reads it. */
const statedNumbering = (innerXml: string): ParagraphNumberingOverride | undefined => {
  const pPr = parseXmlDocument(`<w:pPr xmlns:w="${W}">${innerXml}</w:pPr>`);
  if (pPr === null) {
    throw new Error("Expected the fixture to parse");
  }
  return readParagraphNumbering(findChild(pPr, "w", "numPr"));
};

/** What the serializer emits for a stated override, read back. */
const savedAndReread = (
  numbering: ParagraphNumberingOverride | undefined,
): ParagraphNumberingOverride | undefined => {
  const saved = serializeParagraphFormatting(numbering === undefined ? {} : { numPr: numbering });
  const inner = saved.startsWith("<w:pPr>")
    ? saved.slice("<w:pPr>".length, -"</w:pPr>".length)
    : saved;
  return statedNumbering(inner);
};

const numPrXml = ({ numId, ilvl }: { numId?: number; ilvl?: number }): string =>
  `<w:numPr>${ilvl === undefined ? "" : `<w:ilvl w:val="${String(ilvl)}"/>`}${
    numId === undefined ? "" : `<w:numId w:val="${String(numId)}"/>`
  }</w:numPr>`;

/**
 * The spellings §2.3.1 of the design inventories, plus the arm each reads as.
 * The three that state the reserved id differ only in what sits beside it, and
 * three committed tests used to disagree about what that meant.
 */
const SPELLINGS = [
  ["no element at all", {}, undefined],
  ["an element stating neither slot", { element: true }, undefined],
  ["a reference with a level", { numId: 3, ilvl: 2 }, { kind: "reference", numId: 3, ilvl: 2 }],
  ["a reference with no level", { numId: 3 }, { kind: "reference", numId: 3 }],
  [
    "a reference stating level zero",
    { numId: 3, ilvl: 0 },
    { kind: "reference", numId: 3, ilvl: 0 },
  ],
  ["a level with no reference", { ilvl: 2 }, { kind: "levelOnly", ilvl: 2 }],
  ["the reserved id alone", { numId: 0 }, { kind: "none" }],
  ["the reserved id with level zero", { numId: 0, ilvl: 0 }, { kind: "none" }],
  ["the reserved id with a level", { numId: 0, ilvl: 4 }, { kind: "none" }],
  ["the reserved id with a negative level", { numId: 0, ilvl: -1 }, { kind: "none" }],
] as const satisfies readonly (readonly [
  string,
  { numId?: number; ilvl?: number; element?: true },
  ParagraphNumberingOverride | undefined,
])[];

describe("every `<w:numPr>` spelling parses to one arm and saves back to itself", () => {
  test.each(SPELLINGS)("%s", (_name, authored, expected) => {
    const innerXml =
      "element" in authored ? "<w:numPr></w:numPr>" : numPrXml(authored as { numId?: number });
    const parsed = statedNumbering(innerXml);

    expect(parsed).toEqual(expected);
    // Save and parse again: whatever the arm is, stating it twice states the
    // same thing. An absent `w:ilvl` in particular must not acquire a zero.
    expect(savedAndReread(parsed)).toEqual(expected);
  });

  test("a cancellation still emits the reserved id, because deleting the element uncovers the style", () => {
    expect(serializeParagraphFormatting({ numPr: { kind: "none" } })).toContain(
      '<w:numId w:val="0"/>',
    );
  });
});

const overrides: fc.Arbitrary<ParagraphNumberingOverride | undefined> = fc
  .record({ ilvl: fc.option(fc.nat({ max: 8 }), { nil: undefined }), numId: fc.nat({ max: 9 }) })
  .map(({ numId, ilvl }) => paragraphNumberingFromSlots({ ilvl, numId }));

describe("a stated override is a fixed point of parse → save → parse", () => {
  test(
    "over every slot pair a package can carry",
    () => {
      fc.assert(
        fc.property(overrides, (stated) => {
          expect(savedAndReread(stated)).toEqual(stated);
        }),
        propertyConfig({ numRuns: 200 }),
      );
    },
    propertyTestTimeout(20_000),
  );
});

// That the `none` arm has no slot for a level at all is proved in
// `typecheck/model/paragraphNumbering.typecheck.ts` in `@stll/docx-core`.
describe("a cancellation carries no level", () => {
  test("whatever level the package stated beside the reserved id", () => {
    fc.assert(
      fc.property(fc.integer({ min: -2, max: 12 }), (ilvl) => {
        expect(paragraphNumberingFromSlots({ ilvl, numId: 0 })).toEqual({ kind: "none" });
      }),
      propertyConfig({ numRuns: 40 }),
    );
  });

  test("so nothing resolves the reserved id to a reference", () => {
    fc.assert(
      fc.property(overrides, (stated) => {
        const resolved = resolveParagraphNumbering(stated);
        if (resolved.kind === "reference") {
          expect(resolved.numId).not.toBe(0);
        }
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});

const NUMBERED_STYLES = {
  styles: [
    { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
    {
      type: "paragraph",
      styleId: "Numbered",
      name: "Numbered",
      pPr: { numPr: { kind: "reference", numId: 5, ilvl: 0 } },
    },
    {
      type: "paragraph",
      styleId: "Unnumbered",
      name: "Unnumbered",
      basedOn: "Numbered",
      pPr: { numPr: { kind: "none" } },
    },
  ],
} as const satisfies Document["package"]["styles"];

const NUMBERING = {
  abstractNums: [{ abstractNumId: 5, levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }] }],
  nums: [{ numId: 5, abstractNumId: 5 }],
} as const satisfies NumberingDefinitions;

type ParagraphBlock = Extract<
  Document["package"]["document"]["content"][number],
  { type: "paragraph" }
>;

const documentWith = (paragraph: Omit<ParagraphBlock, "type" | "content">): Document => {
  const document = createEmptyDocument();
  document.package.styles = structuredClone(NUMBERED_STYLES);
  document.package.numbering = structuredClone(NUMBERING);
  document.package.document.content = [
    { type: "paragraph", ...paragraph, content: [{ type: "run", content: [] }] },
  ];
  return document;
};

const paragraphNumberingOf = (document: Document): ParagraphNumberingOverride | undefined => {
  const block = fromProseDoc(toProseDoc(document), document).package.document.content.at(0);
  if (block?.type !== "paragraph") {
    throw new Error("Expected a paragraph");
  }
  return block.formatting?.numPr;
};

describe("a style that cancels the numbering its base supplies", () => {
  test("leaves the paragraph unnumbered", () => {
    const document = documentWith({ formatting: { styleId: "Unnumbered" } });
    const styles = document.package.styles?.styles ?? [];
    const derived = styles.find((style) => style.styleId === "Unnumbered");

    expect(derived?.pPr?.numPr).toEqual({ kind: "none" });
    expect(resolveParagraphNumbering(derived?.pPr?.numPr)).toEqual({ kind: "none" });
  });

  test("and a paragraph on the numbered base keeps its numbering through the editor", () => {
    expect(
      paragraphNumberingOf(documentWith({ formatting: { styleId: "Numbered" } })),
    ).toBeUndefined();
  });
});

describe("the editor round trip keeps what the paragraph stated", () => {
  test(
    "over every arm, direct and inside a recorded w:pPrChange",
    () => {
      fc.assert(
        fc.property(overrides, overrides, (stated, previous) => {
          const document = documentWith({
            formatting: stated === undefined ? {} : { numPr: stated },
            ...(previous === undefined
              ? {}
              : {
                  propertyChanges: [
                    {
                      type: "paragraphPropertyChange" as const,
                      info: { id: 1, author: "Reviewer", date: "2026-01-01" },
                      previousFormatting: { numPr: previous },
                    },
                  ],
                }),
          });
          const block = fromProseDoc(toProseDoc(document), document).package.document.content.at(0);
          if (block?.type !== "paragraph") {
            throw new Error("Expected a paragraph");
          }
          expect(block.formatting?.numPr).toEqual(stated);
          expect(block.propertyChanges?.at(0)?.previousFormatting?.numPr).toEqual(previous);
        }),
        propertyConfig({ numRuns: 60 }),
      );
    },
    propertyTestTimeout(30_000),
  );
});

const NUMBER_FORMATS = [
  "bullet",
  "decimal",
  "lowerLetter",
  "lowerRoman",
  "upperLetter",
  "upperRoman",
] as const satisfies readonly NumberFormat[];

describe("bullet detection reads the level, not the instance id", () => {
  test(
    "over generated numbering definitions",
    () => {
      fc.assert(
        fc.property(
          fc.nat({ max: 20 }).map((offset) => offset + 1),
          fc.nat({ max: 8 }),
          fc.constantFrom(...NUMBER_FORMATS),
          (numId, ilvl, numFmt) => {
            const numbering = createNumberingMap({
              abstractNums: [{ abstractNumId: 1, levels: [{ ilvl, numFmt, lvlText: "%1." }] }],
              nums: [{ numId, abstractNumId: 1 }],
            });

            expect(resolveListState(numbering, { kind: "reference", numId, ilvl })).toEqual({
              type: numFmt === "bullet" ? "bullet" : "numbered",
              level: ilvl,
              numId,
            });
          },
        ),
        propertyConfig({ numRuns: 120 }),
      );
    },
    propertyTestTimeout(20_000),
  );
});
