/**
 * A no-op Document → ProseMirror → Document rebuild must keep
 * `listRendering.levelStarts`. Layout seeds list counters from it, so losing
 * it re-numbered a list defined to start at 5 as "1., 2." until the next
 * DOCX save and parse recomputed the field (issue #845).
 */

import { describe, expect, test } from "bun:test";

import { parseDocx } from "../../docx/parser";
import { createDocx } from "../../docx/rezip";
import { toFlowBlocks } from "../../layout-bridge/convert/toFlowBlocks";
import { fromMarkdown } from "../../markdown/fromMarkdown";
import type { Document } from "../../types/document";
import { updateDocumentContent } from "./fromProseDoc";
import { toProseDoc } from "./toProseDoc";
import { paragraphNumberingReferenceId } from "@stll/docx-core/model";

type FixtureOptions = {
  /** `w:start` for abstract level 0. */
  start?: number;
  /** `w:start` for abstract level 1; the second item moves to level 1 when set. */
  nestedStart?: number;
  /** Instance-level `w:startOverride` for level 0. */
  startOverride?: number;
};

const numberedFixture = async ({
  start = 1,
  nestedStart,
  startOverride,
}: FixtureOptions = {}): Promise<Document> => {
  const model = fromMarkdown("1. Alpha\n2. Beta\n\nTail.");
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
      pPr: { indentation: { left: 1440, hanging: 360 } },
    };
    beta.formatting = { ...beta.formatting, numPr: { kind: "reference", numId, ilvl: 1 } };
  }
  if (startOverride !== undefined) {
    instance.levelOverrides = [{ ilvl: 0, startOverride }];
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
