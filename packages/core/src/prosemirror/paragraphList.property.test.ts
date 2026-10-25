import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import type {
  AbstractNumbering,
  ListLevel,
  NumberingDefinitions,
  Paragraph,
} from "@stll/docx-core/model";
import {
  numberingLevelEntries,
  paragraphNumberingReference,
} from "@stll/docx-core/model";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { FolioDocxReviewer } from "../ai-edits/headless";
import { createNumberingMap } from "../docx/numberingParser";
import { createDocx } from "../docx/rezip";
import { toMarkdown } from "../markdown";
import { fromMarkdown } from "../markdown/fromMarkdown";
import { createListLabelCounter } from "./listLabels";
import { paragraphNumberingAttr } from "./numberingAttr";
import { resolveParagraphList } from "./paragraphList";

setDefaultTimeout(propertyTestTimeout(20_000));

const levelArb: fc.Arbitrary<ListLevel> = fc
  .record({
    ilvl: fc.integer({ min: 0, max: 8 }),
    numFmt: fc.constantFrom("decimal", "lowerLetter", "bullet", "none"),
    lvlText: fc.constantFrom("%1.", "%2)", "•", "", "   "),
    hidden: fc.boolean(),
    legal: fc.boolean(),
  })
  .map(({ ilvl, numFmt, lvlText, hidden, legal }) => {
    const candidate: ListLevel = { ilvl, numFmt, lvlText };
    if (hidden) candidate.rPr = { hidden: true };
    if (legal) candidate.isLgl = true;
    return candidate;
  });

const definitionsArb: fc.Arbitrary<NumberingDefinitions> = fc
  .record({
    abstracts: fc.array(
      fc.record({
        id: fc.integer({ min: 1, max: 3 }),
        levels: fc.array(levelArb, { maxLength: 3 }),
        link: fc.option(fc.constantFrom("ListA", "ListB"), { nil: undefined }),
        isTarget: fc.boolean(),
      }),
      { maxLength: 3 },
    ),
    nums: fc.array(
      fc.record({
        numId: fc.integer({ min: 1, max: 3 }),
        abstractNumId: fc.integer({ min: 1, max: 4 }),
        override: fc.option(levelArb, { nil: undefined }),
      }),
      { maxLength: 3 },
    ),
  })
  .map(({ abstracts, nums }) => ({
    abstractNums: [
      ...abstracts.map(({ id, levels, link, isTarget }) => {
        const abstract: AbstractNumbering = {
          abstractNumId: id,
          levels: link !== undefined && !isTarget ? [] : levels,
        };
        if (isTarget) abstract.styleLink = link;
        else abstract.numStyleLink = link;
        return abstract;
      }),
      { abstractNumId: 90, levels: [] },
      {
        abstractNumId: 91,
        levels: [
          { ilvl: 0, numFmt: "none", lvlText: "%1" },
          { ilvl: 1, numFmt: "decimal", lvlText: "%1", isLgl: true },
        ],
      },
    ],
    nums: [
      ...nums.map(({ numId, abstractNumId, override }) => ({
        numId,
        abstractNumId,
        ...(override === undefined
          ? {}
          : { levelOverrides: [{ ilvl: override.ilvl, lvl: override }] }),
      })),
      { numId: 90, abstractNumId: 90 },
      { numId: 91, abstractNumId: 91 },
    ],
  }));

const paragraph = (text: string, numId: number, ilvl: number): Paragraph => ({
  type: "paragraph",
  formatting: { numPr: paragraphNumberingReference({ numId, ilvl }) },
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const numPrXml = async (buffer: ArrayBuffer): Promise<string[]> => {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")?.async("string");
  return xml?.match(/<w:numPr>[\s\S]*?<\/w:numPr>/gu) ?? [];
};

type Expected = {
  text: string;
  kind: "paragraph" | "listItem";
  marker: boolean;
};

const assertReaders = (
  reviewer: FolioDocxReviewer,
  expected: Expected[],
): void => {
  const content = reviewer.getContent();
  const snapshot = reviewer.snapshot().blocks;
  const rows = reviewer.getContentAsText().split("\n");
  const markdown = toMarkdown(reviewer.toDocument(), {
    annotations: "strip",
    trackedChanges: "clean",
    comments: "strip",
  });
  const formats = new Set(
    reviewer
      .readNumberingDefinitions()
      .map(({ numId, level }) => `${numId}:${level}`),
  );

  for (const [index, result] of expected.entries()) {
    const block = content.find(({ text }) => text === result.text);
    const snap = snapshot.find(({ text }) => text === result.text);
    expect(block?.kind).toBe(result.kind);
    expect(snap?.kind).toBe(result.kind);
    expect(block?.displayLabel !== undefined).toBe(result.marker);
    expect(snap?.displayLabel !== undefined).toBe(result.marker);
    expect(rows[index]?.endsWith(result.text)).toBe(true);
    expect(markdown.includes(result.text)).toBe(true);
    if (block?.displayLabel !== undefined && block.listReference) {
      expect(
        formats.has(
          `${block.listReference.numId}:${block.listReference.level}`,
        ),
      ).toBe(true);
    }
  }
};

describe("paragraph list reader agreement", () => {
  test("empty and whitespace marker projections follow package numbering policy", () => {
    const reference = paragraphNumberingReference({ numId: 1, ilvl: 0 });
    const sources = { numPr: reference, numPrFromStyle: undefined };
    for (const numFmt of [
      "decimal",
      "lowerLetter",
      "bullet",
      "none",
    ] as const) {
      for (const isLegal of [false, true]) {
        for (const hidden of [false, true]) {
          for (const marker of ["", "   "]) {
            const numbering = createNumberingMap({
              abstractNums: [
                {
                  abstractNumId: 1,
                  levels: [
                    {
                      ilvl: 0,
                      numFmt,
                      lvlText: marker,
                      isLgl: isLegal,
                      rPr: { hidden },
                    },
                  ],
                },
              ],
              nums: [{ numId: 1, abstractNumId: 1 }],
            });
            const packaged = resolveParagraphList(sources, numbering);
            const rendering = {
              numId: 1,
              level: 0,
              marker,
              markerTemplate: marker,
              numFmt,
              isLegal,
              markerHidden: hidden,
              isBullet: numFmt === "bullet",
            };
            expect(resolveParagraphList(sources, null, rendering).kind).toBe(
              packaged.kind,
            );
            for (const markerAttr of [
              "listMarker",
              "listMarkerTemplate",
            ] as const) {
              const attrs = {
                numPr: paragraphNumberingAttr(reference),
                [markerAttr]: marker,
                listNumFmt: numFmt,
                listIsLegal: isLegal,
                listMarkerHidden: hidden,
              };
              const info = {
                id: 1,
                author: "Reviewer",
                date: "2026-01-01T00:00:00Z",
              };
              const current = createListLabelCounter()(attrs);
              const original = createListLabelCounter()({
                pPrMark: { kind: "del", info },
                _propertyChanges: [
                  {
                    type: "paragraphPropertyChange",
                    info,
                    previousFormatting: attrs,
                  },
                ],
              });
              for (const result of [current, original]) {
                expect(result.kind).toBe(packaged.kind);
                expect(result.label).toBe(
                  packaged.kind === "marker" ? marker : undefined,
                );
              }
            }
          }
        }
      }
    }
  });

  test("without a numbering part, a paragraph's own cancellation wins over its list projection", () => {
    const projection = { marker: "1.", level: 0, numId: 1, isBullet: false };
    expect(
      resolveParagraphList(
        { numPr: { kind: "none" }, numPrFromStyle: undefined },
        null,
        projection,
      ).kind,
    ).toBe("none");
    expect(
      resolveParagraphList(
        {
          numPr: { kind: "none" },
          numPrFromStyle: paragraphNumberingReference({ numId: 1, ilvl: 0 }),
        },
        null,
        projection,
      ).kind,
    ).toBe("none");
    expect(
      resolveParagraphList(
        { numPr: undefined, numPrFromStyle: undefined },
        null,
        projection,
      ).kind,
    ).toBe("marker");
  });

  test("a cached list projection never stands in for package numbering", () => {
    const projection = { marker: "1.", level: 0, numId: 1, isBullet: false };
    const numbering = createNumberingMap({
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
        },
      ],
      nums: [{ numId: 1, abstractNumId: 1 }],
    });
    const unnumbered = { numPr: undefined, numPrFromStyle: undefined };

    expect(resolveParagraphList(unnumbered, numbering, projection).kind).toBe(
      "none",
    );
    expect(
      resolveParagraphList(
        { ...unnumbered, inheritedReference: null },
        numbering,
        projection,
      ).kind,
    ).toBe("none");
    // Without a package, the projection is the list, whatever the style said,
    // but it names no authored numbering instance.
    expect(
      resolveParagraphList(
        { ...unnumbered, inheritedReference: null },
        null,
        projection,
      ),
    ).toMatchObject({ kind: "marker", source: "projection" });
  });

  test("an inherited style reference resolves through the numbering owner", () => {
    const definitions: NumberingDefinitions = {
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
        },
      ],
      nums: [{ numId: 1, abstractNumId: 1 }],
    };

    expect(
      resolveParagraphList(
        {
          numPr: undefined,
          numPrFromStyle: paragraphNumberingReference({ numId: 1, ilvl: 0 }),
        },
        createNumberingMap(definitions),
      ).kind,
    ).toBe("marker");

    expect(
      resolveParagraphList(
        {
          numPr: { kind: "levelOnly", ilvl: 0 },
          numPrFromStyle: undefined,
          inheritedReference: paragraphNumberingReference({
            numId: 1,
            ilvl: 0,
          }),
        },
        createNumberingMap(definitions),
      ).kind,
    ).toBe("marker");
  });

  test("a level-only override with no style numbering is no list, whatever its tracked predecessor numbered", async () => {
    const document = fromMarkdown("seed");
    document.package.numbering = {
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
        },
      ],
      nums: [{ numId: 1, abstractNumId: 1 }],
    };
    document.package.document.content = [
      {
        type: "paragraph",
        formatting: {
          numPr: { kind: "levelOnly", ilvl: 0 },
        },
        propertyChanges: [
          {
            type: "paragraphPropertyChange",
            info: { id: 1, author: "Reviewer", date: "2026-01-01T00:00:00Z" },
            previousFormatting: {
              numPr: paragraphNumberingReference({ numId: 1, ilvl: 0 }),
            },
          },
        ],
        content: [
          { type: "run", content: [{ type: "text", text: "Tracked level" }] },
        ],
      },
    ];

    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(document),
      {
        author: "Reviewer",
      },
    );

    // ECMA-376 §17.9.19: an `ilvl` without a `numId` takes the numId its
    // style supplies, and nothing here does. The `w:pPrChange` records what
    // rejecting restores; it numbers nothing now.
    assertReaders(reviewer, [
      { text: "Tracked level", kind: "paragraph", marker: false },
    ]);
    expect(reviewer.snapshot().blocks.at(0)?.displayLabel).toBeUndefined();
    expect(
      toMarkdown(reviewer.toDocument(), {
        annotations: "strip",
        trackedChanges: "clean",
        comments: "strip",
      }),
    ).not.toContain("1. Tracked level");
  });

  test("full level overrides use startOverride in definitions and rendered labels", async () => {
    const document = fromMarkdown("seed");
    document.package.numbering = {
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1.", start: 1 }],
        },
      ],
      nums: [
        {
          numId: 1,
          abstractNumId: 1,
          levelOverrides: [
            {
              ilvl: 0,
              lvl: { ilvl: 0, numFmt: "decimal", lvlText: "%1.", start: 1 },
              startOverride: 7,
            },
          ],
        },
      ],
    };
    document.package.document.content = [paragraph("Overridden", 1, 0)];

    const reviewer = await FolioDocxReviewer.fromBuffer(
      await createDocx(document),
      {
        author: "Reviewer",
      },
    );

    expect(reviewer.readNumberingDefinitions()).toContainEqual({
      numId: 1,
      level: 0,
      format: "decimal",
      levelText: "%1.",
      start: 7,
    });
    expect(reviewer.getContent().at(0)).toMatchObject({
      kind: "listItem",
      displayLabel: "7.",
    });
  });

  test("a reference to an undefined level retains provenance without becoming a list", () => {
    const numbering = createNumberingMap({
      abstractNums: [
        {
          abstractNumId: 1,
          levels: [{ ilvl: 0, numFmt: "decimal", lvlText: "%1." }],
        },
      ],
      nums: [{ numId: 5, abstractNumId: 1 }],
    });

    expect(
      resolveParagraphList(
        {
          numPr: paragraphNumberingReference({ numId: 5, ilvl: 8 }),
          numPrFromStyle: undefined,
        },
        numbering,
      ),
    ).toEqual({
      kind: "markerless",
      reason: "levelNotDefined",
      reference: paragraphNumberingReference({ numId: 5, ilvl: 8 }),
      source: "authored",
    });
  });

  test("all readers and save/reopen follow the numbering-level owner", async () => {
    await assertProperty(
      fc.asyncProperty(
        definitionsArb,
        fc.array(fc.boolean(), { maxLength: 5 }),
        async (definitions, includeCachedRendering) => {
          const references = [
            ...definitions.nums.map(({ numId }, index) => ({
              numId,
              ilvl: index % 10,
            })),
            { numId: 90, ilvl: 8 },
            { numId: 91, ilvl: 1 },
          ];
          const document = fromMarkdown("seed");
          document.package.numbering = definitions;
          const numbering = createNumberingMap(definitions);
          const expected = references.map(
            ({ numId, ilvl }, index): Expected => {
              const resolved = resolveParagraphList(
                {
                  numPr: paragraphNumberingReference({ numId, ilvl }),
                  numPrFromStyle: undefined,
                },
                numbering,
              );
              return {
                text: `P${index}`,
                kind: resolved.kind === "marker" ? "listItem" : "paragraph",
                marker: resolved.kind === "marker",
              };
            },
          );
          document.package.document.content = references.map(
            ({ numId, ilvl }, index) => {
              const candidate = paragraph(`P${index}`, numId, ilvl);
              const resolved = resolveParagraphList(
                {
                  numPr: candidate.formatting?.numPr,
                  numPrFromStyle: undefined,
                },
                numbering,
              );
              if (
                includeCachedRendering[index] === true &&
                resolved.kind === "marker"
              ) {
                candidate.listRendering = resolved.rendering;
              }
              return candidate;
            },
          );

          const source = await createDocx(document);
          const reviewer = await FolioDocxReviewer.fromBuffer(source, {
            author: "Reviewer",
          });
          assertReaders(reviewer, expected);
          const saved = await reviewer.toBuffer();
          const reopened = await FolioDocxReviewer.fromBuffer(saved, {
            author: "Reviewer",
          });
          assertReaders(reopened, expected);

          const undefinedLevels = references.filter(
            ({ numId, ilvl }) =>
              !numberingLevelEntries(definitions).some(
                (entry) => entry.numId === numId && entry.ilvl === ilvl,
              ),
          );
          if (undefinedLevels.length > 0)
            expect(await numPrXml(saved)).toEqual(await numPrXml(source));
        },
      ),
      { numRuns: 25, id: "all numbering readers agree" },
    );
  });
});
