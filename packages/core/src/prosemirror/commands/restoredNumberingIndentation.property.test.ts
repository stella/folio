import { expect, test } from "bun:test";
import fc from "fast-check";
import type {
  Document,
  Paragraph,
  ParagraphFormatting,
  StyleDefinitions,
} from "../../types/document";
import { EditorState } from "prosemirror-state";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";
import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { createNumberingMap } from "../../docx/numberingParser";
import { repackDocx } from "../../docx/rezip";
import { expectParagraphAttrs } from "../attrs";
import { fromProseDoc } from "../conversion/fromProseDoc";
import { toProseDoc } from "../conversion/toProseDoc";
import { createDocumentStylesPlugin, getDocumentStyleResolver } from "../plugins/documentStyles";
import { paragraphFormattingWithAuthoredIndentation } from "../../docx/paragraphPropertySource";
import { resolveParagraphChangeAttrs } from "./resolveParagraphProperties";

const CHANGE_INFO = { id: 42, author: "Reviewer", date: "2026-09-01T00:00:00Z" };

const SCENARIOS = [
  { type: "style-numbering-full" },
  { type: "style-numbering-partial" },
  { type: "direct-numbering-full" },
  { type: "direct-numbering-zero-left" },
  { type: "direct-numbering-direct-first-line" },
  { type: "resolver-free-remove-numbering" },
  { type: "resolver-free-change-numbering" },
  { type: "resolver-free-style-numbering" },
] as const;

type Scenario = (typeof SCENARIOS)[number];

type Indents = {
  styleLeft: number;
  styleFirstLine: number;
  oldLevelLeft: number;
  oldLevelFirstLine: number;
  restoredLevelLeft: number;
  restoredLevelFirstLine: number;
};

const reference = (numId: number) => ({ kind: "reference", numId, ilvl: 0 }) as const;

const currentNumberingForScenario = ({ type }: Scenario) => {
  switch (type) {
    case "resolver-free-remove-numbering":
    case "resolver-free-change-numbering":
      return reference(2);
    case "resolver-free-style-numbering":
      return { kind: "levelOnly", ilvl: 0 } as const;
    default:
      return undefined;
  }
};

const makeDocument = (scenario: Scenario, indents: Indents): Document => {
  const styleOwnsNumbering =
    scenario.type === "style-numbering-full" ||
    scenario.type === "style-numbering-partial" ||
    scenario.type === "resolver-free-style-numbering";
  const styleFormatting: ParagraphFormatting = {
    indentLeft:
      scenario.type === "resolver-free-remove-numbering" ? indents.oldLevelLeft : indents.styleLeft,
    ...(scenario.type === "style-numbering-partial" ||
    scenario.type === "resolver-free-remove-numbering"
      ? {}
      : { indentFirstLine: indents.styleFirstLine }),
    ...(styleOwnsNumbering ? { numPr: reference(1) } : {}),
  };
  const currentNumbering = currentNumberingForScenario(scenario);
  const currentStyle =
    scenario.type === "resolver-free-remove-numbering" ||
    scenario.type === "resolver-free-change-numbering" ||
    scenario.type === "resolver-free-style-numbering"
      ? "Styled"
      : "Current";

  let previousFormatting: ParagraphFormatting;
  switch (scenario.type) {
    case "style-numbering-full":
    case "style-numbering-partial":
    case "resolver-free-style-numbering":
      previousFormatting = { styleId: "Styled" };
      break;
    case "direct-numbering-full":
      previousFormatting = { styleId: "Styled", numPr: reference(1) };
      break;
    case "direct-numbering-zero-left":
      previousFormatting = { styleId: "Styled", numPr: reference(1), indentLeft: 0 };
      break;
    case "direct-numbering-direct-first-line":
      previousFormatting = {
        styleId: "Styled",
        numPr: reference(1),
        indentFirstLine: indents.styleFirstLine,
      };
      break;
    case "resolver-free-remove-numbering":
      previousFormatting = { styleId: "Styled" };
      break;
    case "resolver-free-change-numbering":
      previousFormatting = { styleId: "Styled", numPr: reference(1) };
      break;
    default: {
      const unhandled: never = scenario;
      throw new Error(`Unhandled scenario: ${JSON.stringify(unhandled)}`);
    }
  }

  const paragraph: Paragraph = {
    type: "paragraph",
    formatting: {
      styleId: currentStyle,
      ...(currentNumbering === undefined ? {} : { numPr: currentNumbering }),
    },
    propertyChanges: [
      {
        type: "paragraphPropertyChange",
        info: CHANGE_INFO,
        previousFormatting,
      },
    ],
    content: [{ type: "run", content: [{ type: "text", text: "Body" }] }],
  };
  const styles: StyleDefinitions = {
    styles: [
      { type: "paragraph", styleId: "Styled", pPr: styleFormatting },
      { type: "paragraph", styleId: "Current" },
    ],
  };

  return {
    package: {
      document: { content: [paragraph] },
      styles,
      numbering: {
        nums: [
          { numId: 1, abstractNumId: 1 },
          { numId: 2, abstractNumId: 2 },
        ],
        abstractNums: [
          {
            abstractNumId: 1,
            levels: [
              {
                ilvl: 0,
                numFmt: "decimal",
                lvlText: "%1.",
                pPr: {
                  indentLeft: indents.restoredLevelLeft,
                  indentFirstLine: indents.restoredLevelFirstLine,
                },
              },
            ],
          },
          {
            abstractNumId: 2,
            levels: [
              {
                ilvl: 0,
                numFmt: "decimal",
                lvlText: "%1.",
                pPr: {
                  indentLeft: indents.oldLevelLeft,
                  indentFirstLine: indents.oldLevelFirstLine,
                },
              },
            ],
          },
        ],
      },
    },
  };
};

const resolveAndRoundtrip = async (scenario: Scenario, indents: Indents): Promise<void> => {
  const source = makeDocument(scenario, indents);
  const parsed = await parseDocx(await createDocx(source));
  const doc = toProseDoc(parsed);
  const node = doc.child(0);
  const attrs = expectParagraphAttrs(node);
  const useResolver = !scenario.type.startsWith("resolver-free-");
  const state = EditorState.create({
    doc,
    plugins:
      useResolver && parsed.package.styles
        ? [createDocumentStylesPlugin(parsed.package.styles)]
        : [],
  });
  const patch = resolveParagraphChangeAttrs({
    node,
    boundaryCovered: true,
    mode: "reject",
    revisionSet: null,
    styleResolver: useResolver ? (getDocumentStyleResolver(state) ?? null) : null,
    numbering: parsed.package.numbering ? createNumberingMap(parsed.package.numbering) : null,
  });
  expect(patch).not.toBeNull();
  if (patch === null) throw new Error("Expected paragraph property restoration");

  const restored = node.type.create({ ...node.attrs, ...patch }, node.content);
  const outputDoc = doc.type.create(doc.attrs, [restored]);
  const saved = fromProseDoc(outputDoc, parsed, { stylesheetSource: { type: "package" } });
  const reopened = await parseDocx(await repackDocx(saved));
  const reopenedParagraph = reopened.package.document.content.at(0);
  expect(reopenedParagraph?.type).toBe("paragraph");
  if (reopenedParagraph?.type !== "paragraph") throw new Error("Expected a paragraph after save");
  const reopenedAttrs = expectParagraphAttrs(toProseDoc(reopened).child(0));
  const restoredAttrs = expectParagraphAttrs(restored);
  expect(restoredAttrs.indentLeft).toEqual(reopenedAttrs.indentLeft);
  expect(restoredAttrs.indentFirstLine).toEqual(reopenedAttrs.indentFirstLine);
  expect(restoredAttrs.hangingIndent).toEqual(reopenedAttrs.hangingIndent);

  switch (scenario.type) {
    case "style-numbering-full":
      expect(reopenedAttrs.indentLeft).toBe(indents.styleLeft);
      expect(reopenedAttrs.indentFirstLine).toBe(indents.styleFirstLine);
      break;
    case "style-numbering-partial":
      expect(reopenedAttrs.indentLeft).toBe(indents.styleLeft);
      expect(reopenedAttrs.indentFirstLine).toBe(indents.restoredLevelFirstLine);
      break;
    case "direct-numbering-full":
      expect(reopenedAttrs.indentLeft).toBe(indents.restoredLevelLeft);
      expect(reopenedAttrs.indentFirstLine).toBe(indents.restoredLevelFirstLine);
      break;
    case "direct-numbering-zero-left":
      expect(reopenedAttrs.indentLeft).toBe(0);
      expect(reopenedAttrs.indentFirstLine).toBe(indents.restoredLevelFirstLine);
      expect(reopenedParagraph.formatting?.indentLeft).toBe(0);
      break;
    case "direct-numbering-direct-first-line":
      expect(reopenedAttrs.indentLeft).toBe(indents.restoredLevelLeft);
      expect(reopenedAttrs.indentFirstLine).toBe(indents.styleFirstLine);
      expect(reopenedParagraph.formatting?.indentFirstLine).toBe(indents.styleFirstLine);
      break;
    case "resolver-free-remove-numbering":
      expect(attrs["_styleResolvedFormatting"]).toMatchObject({
        indentLeft: indents.oldLevelLeft,
      });
      expect(reopenedAttrs.numPr).toBeUndefined();
      expect(reopenedAttrs.indentLeft).toBe(indents.oldLevelLeft);
      expect(reopenedAttrs.indentFirstLine).toBeUndefined();
      expect(
        paragraphFormattingWithAuthoredIndentation(reopenedParagraph)?.indentLeft,
      ).toBeUndefined();
      expect(
        paragraphFormattingWithAuthoredIndentation(reopenedParagraph)?.indentFirstLine,
      ).toBeUndefined();
      break;
    case "resolver-free-change-numbering":
      expect(reopenedAttrs.numPr).toEqual(reference(1));
      expect(reopenedAttrs.indentLeft).toBe(indents.restoredLevelLeft);
      expect(reopenedAttrs.indentFirstLine).toBe(indents.restoredLevelFirstLine);
      expect(
        paragraphFormattingWithAuthoredIndentation(reopenedParagraph)?.indentLeft,
      ).toBeUndefined();
      break;
    case "resolver-free-style-numbering":
      expect(reopenedAttrs.numPrFromStyle).toEqual(reference(1));
      expect(reopenedAttrs.indentLeft).toBe(indents.styleLeft);
      expect(reopenedAttrs.indentFirstLine).toBe(indents.styleFirstLine);
      break;
    default: {
      const unhandled: never = scenario;
      throw new Error(`Unhandled scenario: ${JSON.stringify(unhandled)}`);
    }
  }
};

test(
  "restored numbering keeps style, direct, and level indentation provenance",
  async () => {
    await assertProperty(
      fc.asyncProperty(fc.integer({ min: 100, max: 10_000 }), async (base) => {
        const indents: Indents = {
          styleLeft: base,
          styleFirstLine: base + 10,
          oldLevelLeft: base + 20,
          oldLevelFirstLine: base + 30,
          restoredLevelLeft: base + 40,
          restoredLevelFirstLine: base + 50,
        };
        for (const scenario of SCENARIOS) {
          await resolveAndRoundtrip(scenario, indents);
        }
      }),
      {
        numRuns: 8,
        id: "restored numbering keeps style, direct, and level indentation provenance",
      },
    );
  },
  propertyTestTimeout(30_000),
);
