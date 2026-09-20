/**
 * A rebuilt `w:tblLook` states what the author stated, and means what it meant.
 *
 * `CT_TblLook` carries a legacy `w:val` bitmask and six `ST_OnOff` attributes
 * that restate the same bits, each in three possible states: absent, an
 * explicit off, an explicit on. The old serializer wrote an attribute only when
 * its flag was truthy, so an explicit `w:lastRow="0"` and a `w:val` the model
 * had no field for both disappeared the moment anything forced a rebuild.
 *
 * Two properties over that whole space. The first is fidelity: parse, force a
 * real save, parse again, and the model is the one the document stated. The
 * second is meaning: what a consumer resolves out of it is the flag where the
 * author gave one and the matching bit of `w:val` where they did not, compared
 * against a reference written straight from ECMA-376 §17.4.57.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { resolveTableLook, TABLE_LOOK_FLAGS, type TableLookFlag } from "./tableLook";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import type { TableFormatting, TableLook } from "../types/document";
import { serializeTableFormatting } from "./serializer/tableSerializer";
import { parseTableProperties } from "./tableParser";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Every `ST_OnOff` spelling, plus the absence that is neither polarity. */
const FLAG_STATES = [undefined, "0", "false", "off", "1", "true", "on"] as const;

type FlagState = (typeof FLAG_STATES)[number];

const statedPolarity = (state: FlagState): boolean | undefined => {
  switch (state) {
    case undefined:
      return undefined;
    case "0":
    case "false":
    case "off":
      return false;
    case "1":
    case "true":
    case "on":
      return true;
    default: {
      const exhaustive: never = state;
      throw new Error(`unhandled ST_OnOff spelling: ${String(exhaustive)}`);
    }
  }
};

type Authored = {
  val: string | undefined;
  flags: Readonly<Record<TableLookFlag, FlagState>>;
};

const authored = fc.record({
  val: fc.option(
    fc
      .integer({ min: 0, max: 0xff_ff })
      .map((bits) => bits.toString(16).toUpperCase().padStart(4, "0")),
    { nil: undefined },
  ),
  flags: fc.record({
    firstRow: fc.constantFrom(...FLAG_STATES),
    lastRow: fc.constantFrom(...FLAG_STATES),
    firstColumn: fc.constantFrom(...FLAG_STATES),
    lastColumn: fc.constantFrom(...FLAG_STATES),
    noHBand: fc.constantFrom(...FLAG_STATES),
    noVBand: fc.constantFrom(...FLAG_STATES),
  }),
});

const tblPrXml = ({ val, flags }: Authored): string => {
  const attrs: string[] = [];
  if (val !== undefined) {
    attrs.push(`w:val="${val}"`);
  }
  for (const flag of TABLE_LOOK_FLAGS) {
    const state = flags[flag];
    if (state !== undefined) {
      attrs.push(`w:${flag}="${state}"`);
    }
  }
  const look = attrs.length > 0 ? `<w:tblLook ${attrs.join(" ")}/>` : "<w:tblLook/>";
  return `<w:tblPr ${NS}>${look}</w:tblPr>`;
};

const parseTblPr = (xml: string): TableFormatting | undefined => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("w:tblPr did not parse");
  }
  return parseTableProperties(root);
};

/**
 * Parse, drop the captured markup, serialize, parse again.
 *
 * Dropping `sourceXml` is what forces the rebuild path. With it in place the
 * serializer replays the original bytes, which would make every one of these
 * cases pass without the model holding anything at all.
 */
const throughRealSave = (xml: string): TableLook | undefined => {
  const rebuilt: TableFormatting = { ...parseTblPr(xml) };
  delete rebuilt.sourceXml;
  const written = serializeTableFormatting(rebuilt);
  if (written === "") {
    return undefined;
  }
  return parseTblPr(written.replace("<w:tblPr>", `<w:tblPr ${NS}>`))?.look;
};

/** What the author stated, read off the generated case rather than the parser. */
const expectedLook = ({ val, flags }: Authored): TableLook | undefined => {
  const look: TableLook = {};
  if (val !== undefined) {
    look.val = val;
  }
  for (const flag of TABLE_LOOK_FLAGS) {
    const polarity = statedPolarity(flags[flag]);
    if (polarity !== undefined) {
      look[flag] = polarity;
    }
  }
  return Object.keys(look).length > 0 ? look : undefined;
};

/**
 * ECMA-376 §17.4.57, written out: the flag if stated, else the bit, else off.
 *
 * The bits are spelled again here rather than imported, so the property
 * compares the implementation against the format instead of against itself.
 */
const referenceEffective = ({ val, flags }: Authored): Record<TableLookFlag, boolean> => {
  const bits = val === undefined ? 0 : Number.parseInt(val, 16);
  const bitOf = {
    firstRow: 0x00_20,
    lastRow: 0x00_40,
    firstColumn: 0x00_80,
    lastColumn: 0x01_00,
    noHBand: 0x02_00,
    noVBand: 0x04_00,
  } as const satisfies Record<TableLookFlag, number>;

  const stated = (flag: TableLookFlag): boolean =>
    // oxlint-disable-next-line no-bitwise -- w:val is an OOXML bitmask
    statedPolarity(flags[flag]) ?? (bits & bitOf[flag]) !== 0;

  return {
    firstRow: stated("firstRow"),
    lastRow: stated("lastRow"),
    firstColumn: stated("firstColumn"),
    lastColumn: stated("lastColumn"),
    noHBand: stated("noHBand"),
    noVBand: stated("noVBand"),
  };
};

describe("w:tblLook survives a rebuild", () => {
  test("a rebuilt look states exactly what the document stated", () => {
    fc.assert(
      fc.property(authored, (sample) => {
        expect(throughRealSave(tblPrXml(sample))).toEqual(expectedLook(sample));
      }),
      propertyConfig({ numRuns: 400 }),
    );
  });

  test("and resolves to the same six answers as the format's own rule", () => {
    fc.assert(
      fc.property(authored, (sample) => {
        expect(resolveTableLook(throughRealSave(tblPrXml(sample)))).toEqual(
          referenceEffective(sample),
        );
      }),
      propertyConfig({ numRuns: 400 }),
    );
  });
});
