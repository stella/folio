/**
 * A no-op Document → ProseMirror → Document rebuild must keep
 * `listRendering.levelStarts`. Layout seeds list counters from it, so losing
 * it re-numbered a list defined to start at 5 as "1., 2." until the next
 * DOCX save and parse recomputed the field (issue #845).
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";

setDefaultTimeout(propertyTestTimeout(30_000));
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../../test/property-testing";

import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import { fromMarkdown } from "../../markdown/fromMarkdown";
import type { BlockContent, Document, StyleDefinitions } from "../../types/document";
import { updateDocumentContent } from "./fromProseDoc";
import { footnoteToProseDoc, headerFooterToProseDoc, toProseDoc } from "./toProseDoc";
import { paragraphNumberingReference, paragraphNumberingReferenceId } from "@stll/docx-core/model";
import {
  applyDocumentOps,
  DOCUMENT_OP_TYPES,
  normalizeForOps,
  paragraphVisibleText,
} from "@stll/docx-core/ops";
import { computeListRendering, getCachedNumberingMap } from "../../docx/numberingParser";
import { listRenderingAttrPatch } from "../listRenderingAttrs";
import { CLEARED_LIST_RENDERING_ATTRS } from "../listMarker";

type FixtureOptions = {
  /** `w:start` for abstract level 0. */
  start?: number;
  /** `w:start` for abstract level 1; the second item moves to level 1 when set. */
  nestedStart?: number;
  /** Instance-level `w:startOverride` for level 0. */
  startOverride?: number;
  foldedListNum?: boolean;
};

const numberedFixture = async ({
  start = 1,
  nestedStart,
  startOverride,
  foldedListNum = false,
}: FixtureOptions = {}): Promise<Document> => {
  const model = fromMarkdown("1. Alpha\n2. Beta\n\nTail.");
  // Operations address authored package ids; normalization does not allocate them.
  for (const [index, block] of model.package.document.content.entries()) {
    if (block.type === "paragraph") block.paraId = (index + 1).toString(16).padStart(8, "0");
  }
  const [alpha, beta] = model.package.document.content;
  if (alpha?.type !== "paragraph" || beta?.type !== "paragraph") {
    throw new Error("fixture must start with two paragraphs");
  }
  const numId = paragraphNumberingReferenceId(alpha.formatting?.numPr);
  const instance = model.package.numbering?.nums.find((num) => num.numId === numId);
  const abstract = model.package.numbering?.abstractNums.find(
    (definition) => definition.abstractNumId === instance?.abstractNumId,
  );
  if (!instance || !abstract) {
    throw new Error("fixture must carry a numbering definition");
  }
  const level0 = abstract.levels.at(0);
  if (!level0) {
    throw new Error("fixture must define numbering level 0");
  }
  level0.start = start;
  if (nestedStart !== undefined) {
    abstract.levels[1] = {
      ilvl: 1,
      start: nestedStart,
      numFmt: "lowerLetter",
      lvlText: "%2.",
      pPr: { indentLeft: 1440, indentFirstLine: -360, hangingIndent: true },
    };
    beta.formatting = { ...beta.formatting, numPr: { kind: "reference", numId, ilvl: 1 } };
  }
  if (startOverride !== undefined) {
    instance.levelOverrides = [{ ilvl: 0, startOverride }];
  }
  if (foldedListNum) {
    alpha.content.unshift(
      {
        type: "complexField",
        instruction: "LISTNUM",
        fieldType: "LISTNUM",
        fieldCode: [],
        fieldResult: [{ type: "run", content: [{ type: "text", text: "(a)" }] }],
      },
      { type: "run", content: [{ type: "tab" }] },
    );
  }
  return parseDocx(await createDocx(model), { preloadFonts: false, detectVariables: false });
};

const markers = (model: Document): (string | null)[] =>
  toFlowBlocks(toProseDoc(model))
    .filter((block) => block.kind === "paragraph")
    .map((block) => block.attrs?.listMarker ?? null);

const levelStartsOf = (model: Document): (number[] | undefined)[] =>
  model.package.document.content.map((block) =>
    block.type === "paragraph" ? block.listRendering?.levelStarts : undefined,
  );

describe("listRendering.levelStarts round-trip", () => {
  test("unchanged ordinary lists preserve each parsed counter marker and rendering field", async () => {
    const initial = await numberedFixture({ startOverride: 5 });
    const projection = toProseDoc(initial);
    for (const [index, paragraph] of initial.package.document.content.entries()) {
      if (paragraph.type !== "paragraph" || paragraph.listRendering === undefined) continue;
      const attrs = projection.child(index).attrs;
      for (const [key, value] of Object.entries(listRenderingAttrPatch(paragraph.listRendering))) {
        expect(attrs[key]).toEqual(value);
      }
    }
    expect(projection.child(0).attrs["listMarker"]).toBe("5.");
    expect(projection.child(1).attrs["listMarker"]).toBe("6.");
  });

  test("unchanged LISTNUM folding preserves its suffix, child advance and second-slot alignment", async () => {
    const initial = await numberedFixture({ start: 7, nestedStart: 1, foldedListNum: true });
    const paragraph = initial.package.document.content.at(0);
    if (paragraph?.type !== "paragraph" || paragraph.listRendering === undefined)
      throw new TypeError("Expected parsed LISTNUM rendering.");
    expect(paragraph.listRendering.marker).toBe("7.\t(a)");
    expect(paragraph.listRendering.markerTemplate).toBe("%1.");
    expect(paragraph.listRendering.foldedMarkerSuffix).toBe("(a)");
    expect(paragraph.listRendering.implicitChildLevelAdvances).toBe(1);
    expect(paragraph.listRendering.markerSecondSlotOffsetTwips).toBe(360);
    expect(paragraph.content.some((item) => item.type === "complexField")).toBe(false);
    const projection = toProseDoc(initial);
    for (const [key, value] of Object.entries(listRenderingAttrPatch(paragraph.listRendering))) {
      expect(projection.child(0).attrs[key]).toEqual(value);
    }
    const rebuilt = updateDocumentContent(initial, projection);
    const rebuiltFirst = rebuilt.package.document.content.at(0);
    expect(rebuiltFirst?.type === "paragraph" ? rebuiltFirst.listRendering : undefined).toEqual(
      paragraph.listRendering,
    );
    expect(markers(rebuilt).at(1)).toBe("b.");
    expect(markers(initial).at(0)).toBe("7.\t(a)");
    expect(markers(rebuilt).at(0)).toBe("7.\t(a)");
  });

  test("same reference recomputes cached rendering when its definition changes", async () => {
    const initial = await numberedFixture({ startOverride: 5 });
    for (const mutation of ["template", "start", "format", "marker", "tabs"] as const) {
      const changed = structuredClone(initial);
      const paragraph = changed.package.document.content.at(0);
      const numbering = changed.package.numbering;
      if (
        paragraph?.type !== "paragraph" ||
        numbering === undefined ||
        paragraph.formatting?.numPr?.kind !== "reference"
      )
        throw new TypeError("Expected numbered rendering fixture.");
      const numPr = paragraph.formatting.numPr;
      const instance = numbering.nums.find((num) => num.numId === numPr.numId);
      const definition = numbering.abstractNums.find(
        (abstract) => abstract.abstractNumId === instance?.abstractNumId,
      );
      const level = definition?.levels.at(0);
      if (!instance || !level) throw new TypeError("Missing numbering definition.");
      switch (mutation) {
        case "template":
          level.lvlText = "(%1)";
          break;
        case "start":
          instance.levelOverrides = [{ ilvl: 0, startOverride: 11 }];
          break;
        case "format":
          level.numFmt = "upperRoman";
          break;
        case "marker":
          level.rPr = { bold: true, allCaps: true };
          break;
        case "tabs":
          level.pPr = { tabs: [{ alignment: "left", position: 900 }] };
          break;
        default: {
          const unreachable: never = mutation;
          throw new TypeError(`Unknown rendering mutation ${unreachable}`);
        }
      }
      const computed = computeListRendering(numPr, getCachedNumberingMap(numbering));
      if (computed === null) throw new TypeError("Changed definition did not resolve.");
      const attrs = toProseDoc(changed).child(0).attrs;
      for (const [key, value] of Object.entries(listRenderingAttrPatch(computed))) {
        expect(attrs[key]).toEqual(value);
      }
      expect(attrs["listMarker"]).toBe(level.lvlText);
    }
  });

  test("generated secondary-story projections follow changing package definitions rather than cached rendering", async () => {
    await assertProperty(
      fc.property(
        fc.array(
          fc.record({
            start: fc.integer({ min: 1, max: 30 }),
            ownership: fc.constantFrom("direct", "cachedStyle", "resolvedStyle", "none"),
            template: fc.constantFrom("%1.", "(%1)", "%1)"),
            format: fc.constantFrom("decimal", "upperRoman", "lowerLetter"),
            bold: fc.boolean(),
            tab: fc.integer({ min: 360, max: 1440 }),
          }),
          { minLength: 8, maxLength: 16 },
        ),
        (trace) => {
          let current = fromMarkdown("1. Alpha");
          const paragraph = current.package.document.content.at(0);
          const originalNumbering = current.package.numbering;
          if (
            paragraph?.type !== "paragraph" ||
            originalNumbering === undefined ||
            paragraph.formatting?.numPr?.kind !== "reference"
          )
            throw new TypeError("Expected numbered paragraph fixture.");
          const cached = computeListRendering(
            paragraph.formatting.numPr,
            getCachedNumberingMap(originalNumbering),
          );
          if (cached === null) throw new TypeError("Expected cached list rendering.");
          paragraph.listRendering = cached;
          const numPr = paragraph.formatting.numPr;
          for (const mutation of trace) {
            current = structuredClone(current);
            const source = current.package.document.content.at(0);
            const numbering = current.package.numbering;
            if (source?.type !== "paragraph" || numbering === undefined)
              throw new TypeError("Expected authored numbering reference.");
            const ilvl = numPr.ilvl ?? 0;
            const instance = numbering.nums.find((num) => num.numId === numPr.numId);
            const level = numbering.abstractNums
              .find((definition) => definition.abstractNumId === instance?.abstractNumId)
              ?.levels.find((candidate) => candidate.ilvl === ilvl);
            if (instance === undefined || level === undefined)
              throw new TypeError("Expected owned numbering definition.");
            instance.levelOverrides = [{ ilvl, startOverride: mutation.start }];
            level.lvlText = mutation.template;
            level.numFmt = mutation.format;
            level.rPr = { bold: mutation.bold };
            level.pPr = { tabs: [{ alignment: "left", position: mutation.tab }] };
            const computed = computeListRendering(numPr, getCachedNumberingMap(numbering));
            if (computed === null)
              throw new TypeError("Changed numbering definition did not resolve.");
            const {
              numPr: _direct,
              numPrFromStyle: _cachedStyle,
              styleId: _styleId,
              ...formatting
            } = source.formatting ?? {};
            const styles = {
              styles: [{ type: "paragraph", styleId: "ListStyle", pPr: { numPr } }],
            } satisfies StyleDefinitions;
            current.package.styles = styles;
            switch (mutation.ownership) {
              case "direct":
                source.formatting = { ...formatting, numPr };
                break;
              case "cachedStyle":
                source.formatting = { ...formatting, numPrFromStyle: numPr };
                break;
              case "resolvedStyle":
                source.formatting = { ...formatting, styleId: "ListStyle" };
                break;
              case "none":
                source.formatting = {
                  ...formatting,
                  styleId: "ListStyle",
                  numPr: { kind: "none" },
                };
                break;
              default: {
                const unreachable: never = mutation.ownership;
                throw new TypeError(`Unknown numbering owner ${unreachable}`);
              }
            }
            const expected =
              mutation.ownership === "none"
                ? CLEARED_LIST_RENDERING_ATTRS
                : listRenderingAttrPatch(computed);
            const content: BlockContent[] = [
              source,
              {
                type: "table",
                rows: [
                  {
                    type: "tableRow",
                    cells: [{ type: "tableCell", content: [structuredClone(source)] }],
                  },
                ],
              },
            ];
            const projections = [
              { document: toProseDoc(current), expectedCount: 1 },
              {
                document: headerFooterToProseDoc(content, { numbering, styles }),
                expectedCount: 2,
              },
              { document: footnoteToProseDoc(content, { numbering, styles }), expectedCount: 2 },
            ];
            for (const projection of projections) {
              let count = 0;
              projection.document.descendants((node) => {
                if (
                  node.type.name !== "paragraph" ||
                  node.textContent !== paragraphVisibleText(source)
                )
                  return;
                count += 1;
                for (const [key, value] of Object.entries(expected)) {
                  expect(node.attrs[key]).toEqual(value);
                }
              });
              expect(count).toBe(projection.expectedCount);
            }
            expect(source.listRendering).toStrictEqual(cached);
          }
        },
      ),
      { numRuns: 30 },
    );
  });

  test("changing a LISTNUM alignment definition invalidates its paragraph-local rendering", async () => {
    const initial = await numberedFixture({ start: 7, nestedStart: 1, foldedListNum: true });
    const changed = structuredClone(initial);
    const nextLevel = changed.package.numbering?.abstractNums
      .at(0)
      ?.levels.find((level) => level.ilvl === 1);
    if (!nextLevel) throw new TypeError("Expected child numbering level.");
    nextLevel.pPr = { ...nextLevel.pPr, indentFirstLine: -720, hangingIndent: true };
    const attrs = toProseDoc(changed).child(0).attrs;
    expect(attrs["listMarker"]).toBe("%1.");
    expect(attrs["listImplicitChildLevelAdvances"]).toBeNull();
    expect(attrs["listMarkerSecondSlotOffsetTwips"]).toBeNull();
  });

  test("explicit no-numbering clears cached LISTNUM and ordinary-list rendering", async () => {
    for (const foldedListNum of [false, true]) {
      const initial = await numberedFixture({ nestedStart: 1, foldedListNum });
      const paragraph = initial.package.document.content.at(0);
      if (paragraph?.type !== "paragraph") throw new TypeError("Expected numbered paragraph.");
      paragraph.formatting = { ...paragraph.formatting, numPr: { kind: "none" } };
      const attrs = toProseDoc(initial).child(0).attrs;
      for (const [key, value] of Object.entries(CLEARED_LIST_RENDERING_ATTRS)) {
        expect(attrs[key]).toEqual(value);
      }
      expect(markers(initial).at(0)).toBeNull();
    }
  });

  test("authored numbering operations replace stale cached starts and clear removed lists", async () => {
    const initial = normalizeForOps(await numberedFixture({ startOverride: 5 }));
    const paragraph = initial.package.document.content.at(0);
    if (paragraph?.type !== "paragraph" || paragraph.paraId === undefined)
      throw new TypeError("Expected identified paragraph.");
    const numId = paragraphNumberingReferenceId(paragraph.formatting?.numPr);
    const original = initial.package.numbering?.nums.find((num) => num.numId === numId);
    if (original === undefined) throw new TypeError("Expected numbering instance.");
    const nextId =
      Math.max(0, ...(initial.package.numbering?.nums.map((num) => num.numId) ?? [])) + 1;
    const changed = applyDocumentOps(initial, [
      {
        type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
        num: { numId: nextId, abstractNumId: original.abstractNumId },
      },
      {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: "main",
        blockId: paragraph.paraId,
        patch: { numPr: paragraphNumberingReference({ numId: nextId, ilvl: 0 }) },
      },
    ]).unwrap();
    const projection = toProseDoc(changed.document);
    expect(projection.firstChild?.attrs["listStartOverride"]).toBeNull();
    expect(projection.firstChild?.attrs["listLevelStarts"]).toEqual([1]);
    expect(markers(changed.document).at(0)).toBe("1.");
    const cleared = applyDocumentOps(changed.document, [
      {
        type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
        story: "main",
        blockId: paragraph.paraId,
        patch: { numPr: { kind: "none" } },
      },
    ]).unwrap();
    expect(toProseDoc(cleared.document).firstChild?.attrs["listNumFmt"]).toBeNull();
    expect(markers(cleared.document).at(0)).toBeNull();
    expect(applyDocumentOps(changed.document, changed.inverse).unwrap().document).toEqual(initial);
  });

  const cases: { name: string; options: FixtureOptions; expected: (string | null)[] }[] = [
    { name: "default start", options: {}, expected: ["1.", "2.", null] },
    { name: "abstract level start 5", options: { start: 5 }, expected: ["5.", "6.", null] },
    {
      name: "instance start override 5",
      options: { startOverride: 5 },
      expected: ["5.", "6.", null],
    },
    {
      name: "multilevel per-level starts",
      options: { start: 5, nestedStart: 3 },
      expected: ["5.", "c.", null],
    },
  ];

  for (const { name, options, expected } of cases) {
    test(`keeps metadata and markers through a no-op rebuild: ${name}`, async () => {
      const initial = await numberedFixture(options);
      expect(markers(initial)).toEqual(expected);

      const rebuilt = updateDocumentContent(initial, toProseDoc(initial));

      expect(levelStartsOf(rebuilt)).toEqual(levelStartsOf(initial));
      expect(rebuilt.package.numbering).toEqual(initial.package.numbering);
      expect(markers(rebuilt)).toEqual(expected);
    });
  }
});
