import { expect, test } from "bun:test";

import { resolveNumberingLevel } from "@stll/docx-core/model";

import type { NumberingDefinitions } from "../types/document";
import { cloneListInstanceWithFormat } from "./listNumberingInstances";

test("clones the effective level and allocates an independent numbering definition", () => {
  const definitions = {
    abstractNums: [
      {
        abstractNumId: 0,
        levels: [
          { ilvl: 0, numFmt: "decimal", lvlText: "%1." },
          { ilvl: 1, numFmt: "lowerLetter", lvlText: "%2." },
        ],
      },
    ],
    nums: [
      {
        numId: 1,
        abstractNumId: 0,
        levelOverrides: [
          {
            ilvl: 0,
            startOverride: 9,
            lvl: {
              ilvl: 0,
              start: 3,
              numFmt: "decimal",
              lvlText: "%1)",
              pPr: { indentLeft: 720 },
            },
          },
          { ilvl: 1, startOverride: 2 },
        ],
      },
    ],
    danglingNumIds: [2],
  } satisfies NumberingDefinitions;
  const elsewhere = { numIds: new Set([3]), abstractNumIds: new Set([1]) };
  const clone = cloneListInstanceWithFormat(
    definitions,
    { sourceNumId: 1, level: 0, numFmt: "upperRoman", start: 1 },
    elsewhere,
  );
  if (!clone) throw new Error("Expected a cloned numbering instance");

  expect([1, 2, 3]).not.toContain(clone.numId);
  expect(clone.definitions.abstractNums).toHaveLength(2);
  expect(clone.definitions.abstractNums.at(0)).toBe(
    definitions.abstractNums.at(0),
  );
  const copied = clone.definitions.abstractNums.at(-1);
  expect(copied?.abstractNumId).toBe(2);
  expect(copied?.levels).toEqual(definitions.abstractNums.at(0)?.levels);
  expect(clone.definitions.nums.at(-1)?.abstractNumId).toBe(
    copied?.abstractNumId,
  );
  expect(resolveNumberingLevel(clone.definitions, clone.numId, 0)).toEqual({
    kind: "level",
    source: "override",
    level: {
      ilvl: 0,
      numFmt: "upperRoman",
      lvlText: "%1)",
      start: 1,
      pPr: { indentLeft: 720 },
    },
  });
  expect(resolveNumberingLevel(clone.definitions, clone.numId, 1)).toEqual(
    resolveNumberingLevel(definitions, 1, 1),
  );
  expect(resolveNumberingLevel(clone.definitions, 1, 0)).toEqual({
    kind: "level",
    source: "override",
    level: {
      ilvl: 0,
      numFmt: "decimal",
      lvlText: "%1)",
      start: 9,
      pPr: { indentLeft: 720 },
    },
  });
  expect(
    cloneListInstanceWithFormat(
      definitions,
      { sourceNumId: 1, level: 2, numFmt: "upperRoman" },
      elsewhere,
    ),
  ).toBeUndefined();
});
