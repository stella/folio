import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { paragraphNumberingReference } from "@stll/docx-core/model";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { expectParagraphBlock } from "../../../../test/paragraphBlock";
import {
  effectiveParagraphNumberingReference,
  readParagraphNumberingAttr,
} from "../prosemirror/numberingAttr";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import {
  createFolioAIEditSnapshotWithStyleResolver,
  remapFolioAIEditSnapshotNumberingReferences,
  remapFolioAIEditSnapshotStyleReferences,
  sourceDocumentOf,
  styleResolverOf,
} from "../ai-edits/snapshot";
import { createEmptyDocument } from "../utils/createDocument";
import type { FolioContentStatedNumbering } from "./content-types";

setDefaultTimeout(propertyTestTimeout(120_000));

const STYLE_SOURCE_IDS = [5, 6] as const;
const STYLE_ID_MAP = new Map([
  ["SourceFive", "ImportedFive"],
  ["SourceSix", "ImportedSix"],
]);

const REMAP_SCENARIOS = [
  {
    name: "chain",
    entries: [
      [5, 6],
      [6, 7],
    ],
  },
  {
    name: "swap",
    entries: [
      [5, 6],
      [6, 5],
    ],
  },
] as const;

const PARAGRAPH_CASES = [
  { paraId: "A1000001", styleId: "SourceFive", sourceNumId: 5, stated: { kind: "inherit" } },
  { paraId: "A1000002", styleId: "SourceSix", sourceNumId: 6, stated: { kind: "inherit" } },
  {
    paraId: "A1000003",
    styleId: "SourceFive",
    sourceNumId: 5,
    stated: { kind: "levelOnly", ilvl: 1 },
  },
  {
    paraId: "A1000004",
    styleId: "SourceFive",
    sourceNumId: 5,
    stated: paragraphNumberingReference({ numId: 6, ilvl: 0 }),
  },
  {
    paraId: "A1000005",
    styleId: "SourceSix",
    sourceNumId: 6,
    stated: paragraphNumberingReference({ numId: 5, ilvl: 1 }),
  },
] as const satisfies readonly {
  paraId: string;
  styleId: "SourceFive" | "SourceSix";
  sourceNumId: (typeof STYLE_SOURCE_IDS)[number];
  stated: FolioContentStatedNumbering;
}[];

const fixtureSnapshot = () => {
  const document = createEmptyDocument();
  const sourceStyles = [
    {
      type: "paragraph" as const,
      styleId: "SourceFive",
      name: "Source Five",
      pPr: { numPr: paragraphNumberingReference({ numId: 5 }) },
    },
    {
      type: "paragraph" as const,
      styleId: "SourceSix",
      name: "Source Six",
      pPr: { numPr: paragraphNumberingReference({ numId: 6 }) },
    },
  ];
  document.package.styles = {
    styles: [
      { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
      ...sourceStyles,
    ],
  };
  document.package.numbering = {
    abstractNums: STYLE_SOURCE_IDS.map((abstractNumId) => ({
      abstractNumId,
      levels: [{ ilvl: 0, numFmt: "decimal" as const, lvlText: "%1." }],
    })),
    nums: STYLE_SOURCE_IDS.map((numId) => ({ numId, abstractNumId: numId })),
  };
  document.package.document.content = PARAGRAPH_CASES.map(
    ({ paraId, styleId, sourceNumId, stated }) => ({
      type: "paragraph" as const,
      paraId,
      textId: paraId,
      formatting: {
        styleId,
        numPrFromStyle: paragraphNumberingReference({ numId: sourceNumId }),
        ...(stated.kind !== "inherit" && { numPr: stated }),
      },
      content: [{ type: "run" as const, content: [{ type: "text" as const, text: paraId }] }],
    }),
  );

  const styleResolver = createStyleResolver(document.package.styles);
  return {
    snapshot: createFolioAIEditSnapshotWithStyleResolver(toProseDoc(document), styleResolver),
    sourceStyles,
  };
};

const mapNumberOnce = (numId: number, entries: readonly (readonly [number, number])[]) =>
  new Map(entries).get(numId) ?? numId;

const reference = (numId: number, level: number) => ({ numId, level });

test("style rebinding then numbering remapping applies source ids once", async () => {
  await assertProperty(
    fc.asyncProperty(fc.constant(REMAP_SCENARIOS), async (scenarios) => {
      const { snapshot, sourceStyles } = fixtureSnapshot();
      const sourceDocument = sourceDocumentOf(snapshot);
      const sourceNodes: Array<{ attrs: Record<string, unknown>; text: string }> = [];
      sourceDocument.descendants((node) => {
        if (node.type.name !== "paragraph") return;
        sourceNodes.push({ attrs: node.attrs, text: node.textContent });
      });
      expect(sourceNodes).toHaveLength(PARAGRAPH_CASES.length);

      for (const { entries } of scenarios) {
        const numberingReferenceMap = new Map(entries);
        const importedStyles = sourceStyles.map((style) => {
          const sourceNumId = style.styleId === "SourceFive" ? 5 : 6;
          return {
            type: "paragraph" as const,
            styleId: STYLE_ID_MAP.get(style.styleId) ?? style.styleId,
            name: style.name,
            pPr: {
              numPr: paragraphNumberingReference({ numId: mapNumberOnce(sourceNumId, entries) }),
            },
          };
        });
        const importedStyleResolver = createStyleResolver({
          styles: [
            { type: "paragraph", styleId: "Normal", name: "Normal", default: true },
            ...importedStyles,
          ],
        });
        const styleRebound = remapFolioAIEditSnapshotStyleReferences({
          snapshot,
          styleIdMap: STYLE_ID_MAP,
          defaultParagraphStyleId: undefined,
          importedStyleResolver,
        });
        expect(styleResolverOf(styleRebound)).toBe(importedStyleResolver);
        const remapped = remapFolioAIEditSnapshotNumberingReferences(
          styleRebound,
          numberingReferenceMap,
        );
        const reboundNodes: Array<{ attrs: Record<string, unknown>; text: string }> = [];
        sourceDocumentOf(remapped).descendants((node) => {
          if (node.type.name !== "paragraph") return;
          reboundNodes.push({ attrs: node.attrs, text: node.textContent });
        });

        for (const [index, paragraph] of PARAGRAPH_CASES.entries()) {
          const sourceNode = sourceNodes.at(index);
          const reboundNode = reboundNodes.at(index);
          const sourceDirect = paragraph.stated.kind === "inherit" ? undefined : paragraph.stated;
          const expectedDirect =
            sourceDirect?.kind === "reference"
              ? paragraphNumberingReference({
                  numId: mapNumberOnce(sourceDirect.numId, entries),
                  ilvl: sourceDirect.ilvl,
                })
              : sourceDirect;
          const expectedStyleNumId = mapNumberOnce(paragraph.sourceNumId, entries);
          expect(sourceNode).toBeDefined();
          expect(reboundNode).toBeDefined();
          expect(sourceNode?.attrs["numPr"]).toEqual(
            paragraph.stated.kind === "inherit" ? null : paragraph.stated,
          );
          expect(sourceNode?.attrs["numPrFromStyle"]).toEqual(
            paragraphNumberingReference({ numId: paragraph.sourceNumId }),
          );
          expect(reboundNode?.attrs["styleId"]).toBe(STYLE_ID_MAP.get(paragraph.styleId));
          expect(reboundNode?.attrs["numPr"]).toEqual(expectedDirect ?? null);
          expect(reboundNode?.attrs["numPrFromStyle"]).toEqual(
            paragraphNumberingReference({ numId: expectedStyleNumId }),
          );

          const expectedEffective = (() => {
            switch (paragraph.stated.kind) {
              case "inherit":
                return reference(expectedStyleNumId, 0);
              case "levelOnly":
                return reference(expectedStyleNumId, paragraph.stated.ilvl);
              case "reference":
                return reference(
                  mapNumberOnce(paragraph.stated.numId, entries),
                  paragraph.stated.ilvl ?? 0,
                );
              case "none":
                return undefined;
              default: {
                const unhandled: never = paragraph.stated;
                return unhandled;
              }
            }
          })();
          if (expectedEffective === undefined) {
            throw new Error("Fixture paragraph unexpectedly has no effective numbering reference");
          }
          const expectedStated =
            paragraph.stated.kind === "reference"
              ? paragraphNumberingReference({
                  numId: mapNumberOnce(paragraph.stated.numId, entries),
                  ilvl: paragraph.stated.ilvl,
                })
              : paragraph.stated;
          expect(
            effectiveParagraphNumberingReference({
              numPr: readParagraphNumberingAttr(reboundNode?.attrs["numPr"]),
              numPrFromStyle: readParagraphNumberingAttr(reboundNode?.attrs["numPrFromStyle"]),
            }),
          ).toEqual(
            paragraphNumberingReference({
              numId: expectedEffective.numId,
              ilvl: expectedEffective.level,
            }),
          );
          const block = expectParagraphBlock(
            remapped.blocks.find(({ text }) => text === paragraph.paraId),
          );
          expect(block.statedNumbering).toEqual(expectedStated);
          expect(block.listReference).toEqual(expectedEffective);
        }
      }
    }),
    { numRuns: 1 },
  );
});
