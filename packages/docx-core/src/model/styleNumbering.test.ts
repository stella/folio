import { describe, expect, test } from "bun:test";

import type { StyleDefinitions } from "./styles";
import {
  createParagraphStyleNumberingResolver,
  numberingIdentityEqual,
  verifiedNumberingLevelIndent,
} from "./styleNumbering";

describe("paragraph style numbering", () => {
  test("compares effective numbering identities", () => {
    expect(numberingIdentityEqual(null, undefined)).toBe(true);
    expect(numberingIdentityEqual({ numId: 7 }, { numId: 7, ilvl: 0 })).toBe(true);
    expect(numberingIdentityEqual({ numId: 7 }, { numId: 8 })).toBe(false);
    expect(numberingIdentityEqual({ numId: 7, ilvl: 1 }, { numId: 7 })).toBe(false);
    expect(numberingIdentityEqual({}, {})).toBe(true);
  });

  test("verifies numbering-owned indentation against its latent baseline", () => {
    const owned = { indentLeft: 720 };
    const baseline = {
      indentLeft: 720,
      indentFirstLine: -360,
      hangingIndent: true,
    };
    const formatting = {
      numPr: { numId: 7, ilvl: 0 },
      numberingLevelIndent: {
        type: "owned" as const,
        numId: 7,
        ilvl: 0,
        baseline,
        owned,
      },
    };

    expect(verifiedNumberingLevelIndent(formatting)).toBe(owned);
    expect(
      verifiedNumberingLevelIndent({
        ...formatting,
        numberingLevelIndent: { type: "latent", numId: 7, ilvl: 0, baseline },
      }),
    ).toBeUndefined();
    expect(
      verifiedNumberingLevelIndent({
        ...formatting,
        numberingLevelIndent: {
          ...formatting.numberingLevelIndent,
          owned: { indentLeft: 360 },
        },
      }),
    ).toBeUndefined();
    const malformed = { ...formatting };
    Reflect.deleteProperty(malformed.numberingLevelIndent, "owned");
    expect(verifiedNumberingLevelIndent(malformed)).toBeUndefined();
  });

  test("resolves partial and cyclic paragraph-style chains", () => {
    const styles = {
      styles: [
        {
          styleId: "Base",
          type: "paragraph",
          pPr: { numPr: { numId: 7, ilvl: 2 } },
        },
        {
          styleId: "Partial",
          type: "paragraph",
          basedOn: "Base",
          pPr: { numPr: { ilvl: 0 } },
        },
        { styleId: "CycleA", type: "paragraph", basedOn: "CycleB" },
        {
          styleId: "CycleB",
          type: "paragraph",
          basedOn: "CycleA",
          pPr: { numPr: { numId: 8 } },
        },
      ],
    } satisfies StyleDefinitions;
    const styleSuppliesNumbering = createParagraphStyleNumberingResolver(styles);

    expect(
      styleSuppliesNumbering({
        styleId: "Partial",
        numPrFromStyle: { numId: 7, ilvl: 0 },
      }),
    ).toBe(true);
    expect(
      styleSuppliesNumbering({
        styleId: "Partial",
        numPrFromStyle: { numId: 7, ilvl: 2 },
      }),
    ).toBe(false);
    expect(
      styleSuppliesNumbering({ styleId: "CycleA", numPrFromStyle: { numId: 8 } }),
    ).toBe(true);
    expect(
      styleSuppliesNumbering({ styleId: "Partial", numPrFromStyle: { ilvl: 0 } }),
    ).toBe(false);
  });
});
