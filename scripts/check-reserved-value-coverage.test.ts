/** The reserved-value coverage check, against a schema graph small enough to read. */

import { describe, expect, test } from "bun:test";

import {
  collectCandidates,
  enumerationMismatches,
  enumTokensOf,
  excludedSlots,
  expandSlot,
} from "./check-reserved-value-coverage";
import { expandSimpleType } from "./lib/narrowed-enum-schema-types";
import { buildIndex, type SchemaGraph, slotKey, WML_NAMESPACE } from "./lib/ooxml-schema-graph";

const qualified = (name: string): string => `{${WML_NAMESPACE}}${name}`;
const complexType = (name: string): string => `complexType:${qualified(name)}`;
const simpleType = (name: string): string => `simpleType:${qualified(name)}`;

/**
 * `w:document` -> `w:body` -> `w:shd` and `w:gridSpan`, plus a `w:noWrap` the
 * document cannot reach. Enough to exercise every branch of the derivation.
 */
const graph = (): SchemaGraph => ({
  namespaces: [{ uri: WML_NAMESPACE }],
  inheritance: [],
  symbols: [
    {
      id: `element:${qualified("document")}`,
      kind: "element",
      name: "document",
      namespace: WML_NAMESPACE,
      type: qualified("CT_Document"),
    },
    {
      id: complexType("CT_Document"),
      kind: "complexType",
      name: "CT_Document",
      namespace: WML_NAMESPACE,
    },
    { id: complexType("CT_Body"), kind: "complexType", name: "CT_Body", namespace: WML_NAMESPACE },
    { id: complexType("CT_Shd"), kind: "complexType", name: "CT_Shd", namespace: WML_NAMESPACE },
    {
      id: complexType("CT_DecimalNumber"),
      kind: "complexType",
      name: "CT_DecimalNumber",
      namespace: WML_NAMESPACE,
    },
    {
      id: complexType("CT_OnOff"),
      kind: "complexType",
      name: "CT_OnOff",
      namespace: WML_NAMESPACE,
    },
    {
      id: simpleType("ST_Shd"),
      kind: "simpleType",
      name: "ST_Shd",
      namespace: WML_NAMESPACE,
      base: "{http://www.w3.org/2001/XMLSchema}string",
      enumValues: ["clear", "solid", "nil"],
    },
    {
      id: simpleType("ST_DecimalNumber"),
      kind: "simpleType",
      name: "ST_DecimalNumber",
      namespace: WML_NAMESPACE,
      base: "{http://www.w3.org/2001/XMLSchema}integer",
    },
    {
      id: simpleType("ST_OnOff"),
      kind: "simpleType",
      name: "ST_OnOff",
      namespace: WML_NAMESPACE,
      memberTypes: [qualified("ST_OnOff1")],
    },
    {
      id: simpleType("ST_OnOff1"),
      kind: "simpleType",
      name: "ST_OnOff1",
      namespace: WML_NAMESPACE,
      base: "{http://www.w3.org/2001/XMLSchema}string",
      enumValues: ["on", "off"],
    },
  ],
  children: [
    {
      kind: "element",
      name: "body",
      namespace: WML_NAMESPACE,
      owner: complexType("CT_Document"),
      type: qualified("CT_Body"),
    },
    {
      kind: "element",
      name: "shd",
      namespace: WML_NAMESPACE,
      owner: complexType("CT_Body"),
      type: qualified("CT_Shd"),
    },
    {
      kind: "element",
      name: "gridSpan",
      namespace: WML_NAMESPACE,
      owner: complexType("CT_Body"),
      type: qualified("CT_DecimalNumber"),
    },
    {
      kind: "element",
      name: "noWrap",
      namespace: WML_NAMESPACE,
      owner: complexType("CT_Unreachable"),
      type: qualified("CT_OnOff"),
    },
  ],
  attributes: [
    {
      kind: "attribute",
      name: "val",
      owner: complexType("CT_Shd"),
      type: qualified("ST_Shd"),
    },
    {
      kind: "attribute",
      name: "space",
      owner: complexType("CT_Shd"),
      type: qualified("ST_DecimalNumber"),
      default: "0",
    },
    {
      kind: "attribute",
      name: "size",
      owner: complexType("CT_Shd"),
      type: qualified("ST_DecimalNumber"),
    },
    {
      kind: "attribute",
      name: "val",
      owner: complexType("CT_DecimalNumber"),
      type: qualified("ST_DecimalNumber"),
    },
    {
      kind: "attribute",
      name: "val",
      owner: complexType("CT_OnOff"),
      type: qualified("ST_OnOff"),
    },
  ],
});

describe("candidate derivation", () => {
  const candidates = collectCandidates(graph(), buildIndex(graph()));
  const bySlot = new Map(candidates.map((candidate) => [candidate.slot, candidate]));

  test("an enumeration carrying a reserved token is a candidate", () => {
    expect(bySlot.get(slotKey(WML_NAMESPACE, "shd", "val"))).toMatchObject({
      kind: "enum",
      detail: "clear, nil",
    });
  });

  test("an attribute with an XSD default is a candidate", () => {
    expect(bySlot.get(slotKey(WML_NAMESPACE, "shd", "space"))).toMatchObject({
      kind: "default",
      detail: 'default "0"',
    });
  });

  test("a plain attribute with neither is not", () => {
    expect(bySlot.has(slotKey(WML_NAMESPACE, "shd", "size"))).toBe(false);
  });

  test("an element no rebuilt part can reach is not", () => {
    expect(bySlot.has(slotKey(WML_NAMESPACE, "noWrap", "val"))).toBe(false);
  });

  test("a numeric sentinel is invisible to the derivation, so it comes from the prose list", () => {
    // `w:gridSpan/@w:val` is reachable and unfacetted: nothing in the schema
    // marks 1 or 0 as reserved, which is why PROSE_SLOTS exists.
    expect(bySlot.has(slotKey(WML_NAMESPACE, "gridSpan", "val"))).toBe(false);
  });
});

describe("enumTokensOf", () => {
  const index = buildIndex(graph());

  test("follows union members", () => {
    expect(enumTokensOf(index, qualified("ST_OnOff"))).toEqual(["on", "off"]);
  });

  test("returns nothing for a type with no enumeration", () => {
    expect(enumTokensOf(index, qualified("ST_DecimalNumber"))).toEqual([]);
  });
});

describe("expandSlot", () => {
  test("resolves the registry's prefix to a namespace URI", () => {
    expect(expandSlot("w:tcW@type")).toBe(slotKey(WML_NAMESPACE, "tcW", "type"));
  });

  test("rejects a prefix the registry does not declare", () => {
    expect(() => expandSlot("zz:tcW@type")).toThrow(/namespace prefix/u);
  });
});

describe("excludedSlots", () => {
  test("maps every slot in a group to that group's reason", () => {
    const excluded = excludedSlots([{ reason: "Preserved verbatim.", slots: ["w:a@x", "w:b@y"] }]);
    expect(excluded.get(slotKey(WML_NAMESPACE, "a", "x"))).toBe("Preserved verbatim.");
    expect(excluded.size).toBe(2);
  });

  test("rejects a slot excluded twice, because one entry is then dead", () => {
    expect(() =>
      excludedSlots([
        { reason: "First.", slots: ["w:a@x"] },
        { reason: "Second.", slots: ["w:a@x"] },
      ]),
    ).toThrow(/excluded twice/u);
  });
});

/**
 * The assertion that would have caught the stale `ST_ThemeColor` claim.
 *
 * The registry named `w:color@themeColor` and described what `none` meant
 * there, and the union behind it spelled six members the DrawingML way and
 * omitted seven. Naming a slot was all the check asked for, so the entry read
 * as covered for four releases.
 */
describe("enumeration bindings", () => {
  const themeColor = expandSimpleType("w:ST_ThemeColor");
  const schemaTokens = [
    "dark1",
    "light1",
    "dark2",
    "light2",
    "accent1",
    "accent2",
    "accent3",
    "accent4",
    "accent5",
    "accent6",
    "hyperlink",
    "followedHyperlink",
    "none",
    "background1",
    "text1",
    "background2",
    "text2",
  ];

  test("a bound enumeration that matches the schema reports nothing", () => {
    expect(
      enumerationMismatches(
        new Map([[themeColor, { slots: new Set(["w:color@themeColor"]), tokens: schemaTokens }]]),
      ),
    ).toEqual([]);
  });

  test("a schema enumeration nothing in the model is bound to is reported", () => {
    const unbound = expandSimpleType("w:ST_Invented");
    expect(
      enumerationMismatches(
        new Map([[unbound, { slots: new Set(["w:invented@val"]), tokens: ["a", "b"] }]]),
      ),
    ).toEqual([
      expect.stringContaining("ST_Invented (w:invented@val): no model enumeration is bound to it."),
    ]);
  });

  test("a bound enumeration the schema has grown past is reported member by member", () => {
    // The schema graph gaining a member the model's list does not carry is the
    // same failure as the model's list drifting, seen from the other side.
    const [reported] = enumerationMismatches(
      new Map([
        [
          themeColor,
          { slots: new Set(["w:color@themeColor"]), tokens: [...schemaTokens, "accent7"] },
        ],
      ]),
    );
    expect(reported).toContain("ST_ThemeColor (w:color@themeColor)");
    expect(reported).toContain("omits accent7");
  });
});
