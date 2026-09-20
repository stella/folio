/**
 * `word/numbering.xml` survives a rebuild from the model alone.
 *
 * A repack copies the part across and splices single definitions into it by
 * id, so nothing here runs on an ordinary save and the reader's gaps were
 * invisible: the survival census reported every pair in the part unmeasured
 * rather than lost. The census now forces the part serializer, and this is the
 * same assertion in the package that owns the reader, so a regression fails
 * `bun test packages/core` rather than waiting for a full sweep.
 *
 * The properties are over the declared children of each container rather than
 * over examples: the generated `CONTAINER_CHILDREN` set is the schema's, so a
 * child a schema refresh adds arrives in the test without anybody remembering
 * to add it, and the order assertion fails until somebody places it.
 */

import { readFileSync } from "node:fs";

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import { NUMBER_FORMAT_VALUES } from "../types/documentEnumValues";
import { CONTAINER_CHILDREN } from "./containerChildren.gen";
import { parseNumbering } from "./numberingParser";
import { serializeNumberingXml } from "./serializer/numberingSerializer";
import { getChildElements, getLocalName, parseXmlDocument } from "./xmlParser";

type NamedChild = readonly [name: string, xml: string];

/**
 * One instance of each child a container declares, in the order it declares
 * them.
 *
 * Written out rather than generated, because each needs the attributes its own
 * type requires; bound to the generated sets by the totality tests below, so
 * none can quietly fall behind the schema.
 */
const NUMBERING_CHILDREN: readonly NamedChild[] = [
  ["numPicBullet", '<w:numPicBullet w:numPicBulletId="0"><w:pict/></w:numPicBullet>'],
  ["abstractNum", '<w:abstractNum w:abstractNumId="1"/>'],
  ["num", '<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>'],
  ["numIdMacAtCleanup", '<w:numIdMacAtCleanup w:val="7"/>'],
];

const ABSTRACT_NUM_CHILDREN: readonly NamedChild[] = [
  ["nsid", '<w:nsid w:val="1B2C3D4E"/>'],
  ["multiLevelType", '<w:multiLevelType w:val="hybridMultilevel"/>'],
  ["tmpl", '<w:tmpl w:val="0409001D"/>'],
  ["name", '<w:name w:val="Bullets"/>'],
  ["styleLink", '<w:styleLink w:val="ListBullet"/>'],
  ["numStyleLink", '<w:numStyleLink w:val="ListNumber"/>'],
  ["lvl", '<w:lvl w:ilvl="0"/>'],
];

const LVL_CHILDREN: readonly NamedChild[] = [
  ["start", '<w:start w:val="3"/>'],
  ["numFmt", '<w:numFmt w:val="lowerRoman"/>'],
  ["lvlRestart", '<w:lvlRestart w:val="0"/>'],
  ["pStyle", '<w:pStyle w:val="ListParagraph"/>'],
  ["isLgl", "<w:isLgl/>"],
  ["suff", '<w:suff w:val="space"/>'],
  ["lvlText", '<w:lvlText w:val="%1."/>'],
  ["lvlPicBulletId", '<w:lvlPicBulletId w:val="0"/>'],
  ["legacy", '<w:legacy w:legacy="1" w:legacySpace="120" w:legacyIndent="240"/>'],
  ["lvlJc", '<w:lvlJc w:val="right"/>'],
  ["pPr", '<w:pPr><w:ind w:left="720"/></w:pPr>'],
  ["rPr", "<w:rPr><w:b/></w:rPr>"],
];

/**
 * `w:numFmt` and `w:lvlText` are written whether or not the source wrote one:
 * `ListLevel` holds both as required fields with a default, so a level that
 * states neither comes back stating both. Recorded here rather than worked
 * around, because it is what the sink's index is measured against.
 */
const ALWAYS_WRITTEN_LVL_CHILDREN: ReadonlySet<string> = new Set(["numFmt", "lvlText"]);

const NUM_CHILDREN: readonly NamedChild[] = [
  ["abstractNumId", '<w:abstractNumId w:val="1"/>'],
  ["lvlOverride", '<w:lvlOverride w:ilvl="0"/>'],
];

const LVL_OVERRIDE_CHILDREN: readonly NamedChild[] = [
  ["startOverride", '<w:startOverride w:val="5"/>'],
  ["lvl", '<w:lvl w:ilvl="0"/>'],
];

const partWith = (body: string, rootAttributes = ""): string =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
  ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"' +
  ` xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"${rootAttributes}>` +
  `${body}</w:numbering>`;

const abstractNumWith = (children: string, attributes = ""): string =>
  partWith(`<w:abstractNum w:abstractNumId="1"${attributes}>${children}</w:abstractNum>`);

const lvlWith = (children: string, attributes = ""): string =>
  abstractNumWith(`<w:lvl w:ilvl="0"${attributes}>${children}</w:lvl>`);

const numWith = (children: string, attributes = ""): string =>
  partWith(`<w:num w:numId="1"${attributes}>${children}</w:num>`);

const lvlOverrideWith = (children: string, attributes = ""): string =>
  numWith(
    `<w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"${attributes}>${children}</w:lvlOverride>`,
  );

const rebuild = (xml: string): string => serializeNumberingXml(parseNumbering(xml).definitions);

/** The child element names at `path`, in document order. */
const childNamesAt = (xml: string, path: readonly string[]): string[] => {
  const root = parseXmlDocument(xml);
  if (!root) {
    throw new Error(`the rebuilt part did not parse: ${xml}`);
  }
  let element = root;
  for (const step of path) {
    const next = getChildElements(element).find((node) => getLocalName(node.name) === step);
    if (!next) {
      throw new Error(`the rebuilt part has no ${path.join("/")}: ${xml}`);
    }
    element = next;
  }
  return getChildElements(element).map((node) => getLocalName(node.name));
};

const declaredNames = (children: readonly NamedChild[]): string[] =>
  children.map(([name]) => name).toSorted();

describe("the numbering reader covers what the schema declares", () => {
  test.each([
    ["w:numbering", NUMBERING_CHILDREN],
    ["w:abstractNum", ABSTRACT_NUM_CHILDREN],
    ["w:lvl", LVL_CHILDREN],
    ["w:num", NUM_CHILDREN],
    ["w:lvlOverride", LVL_OVERRIDE_CHILDREN],
  ] as const)("every declared child of %s has a case here", (container, children) => {
    expect(declaredNames(children)).toEqual([...CONTAINER_CHILDREN[container]].toSorted());
  });
});

type PlacementCase = {
  container: string;
  children: readonly NamedChild[];
  wrap: (children: string) => string;
  path: readonly string[];
  /** Children the writer emits whether or not the source carried them. */
  always?: ReadonlySet<string>;
  minLength: number;
};

const PLACEMENT_CASES: readonly PlacementCase[] = [
  {
    container: "w:numbering",
    children: NUMBERING_CHILDREN,
    wrap: partWith,
    path: [],
    minLength: 1,
  },
  {
    container: "w:abstractNum",
    children: ABSTRACT_NUM_CHILDREN,
    wrap: (children) => abstractNumWith(children),
    path: ["abstractNum"],
    minLength: 0,
  },
  {
    container: "w:lvl",
    children: LVL_CHILDREN,
    wrap: (children) => lvlWith(children),
    path: ["abstractNum", "lvl"],
    always: ALWAYS_WRITTEN_LVL_CHILDREN,
    minLength: 0,
  },
  {
    // `w:abstractNumId` is the one child `CT_Num` requires, so it is in every
    // subset rather than one of the ones varied: an instance that names no
    // template numbers nothing and the reader refuses it.
    container: "w:num",
    children: NUM_CHILDREN,
    wrap: (children) => numWith(children),
    path: ["num"],
    minLength: 2,
  },
  {
    container: "w:lvlOverride",
    children: LVL_OVERRIDE_CHILDREN,
    wrap: (children) => lvlOverrideWith(children),
    path: ["num", "lvlOverride"],
    minLength: 0,
  },
];

describe("a numbering definition survives a rebuild from the model", () => {
  test.each(PLACEMENT_CASES.map((placement) => [placement.container, placement] as const))(
    "any subset of %s's children comes back, in place",
    (_container, placement) => {
      fc.assert(
        fc.property(
          fc.subarray([...placement.children], { minLength: placement.minLength }),
          (children) => {
            const source = placement.wrap(children.map(([, xml]) => xml).join(""));
            const expected = placement.children
              .filter(
                ([name]) =>
                  placement.always?.has(name) === true ||
                  children.some(([chosen]) => chosen === name),
              )
              .map(([name]) => name);
            expect(childNamesAt(rebuild(source), placement.path)).toEqual(expected);
          },
        ),
        propertyConfig({ numRuns: 60 }),
      );
    },
  );

  test.each(PLACEMENT_CASES.map((placement) => [placement.container, placement] as const))(
    "a rebuilt part is a fixed point of the reader for %s",
    (_container, placement) => {
      fc.assert(
        fc.property(
          fc.subarray([...placement.children], { minLength: placement.minLength }),
          (children) => {
            const source = placement.wrap(children.map(([, xml]) => xml).join(""));
            const once = parseNumbering(rebuild(source)).definitions;
            const twice = parseNumbering(rebuild(rebuild(source))).definitions;
            expect(twice).toEqual(once);
          },
        ),
        propertyConfig({ numRuns: 60 }),
      );
    },
  );
});

/**
 * What a rebuild used to lose, one case per reason.
 *
 * Named apart from the properties because each is a different failure: an
 * element or attribute the model had no field for, a value the reader refused,
 * an element read as nothing, and a flag read off the wrong attribute.
 */
describe("what a rebuild used to lose", () => {
  test("a list template keeps the identity Word recognises it by", () => {
    const rebuilt = rebuild(
      abstractNumWith('<w:nsid w:val="1B2C3D4E"/><w:tmpl w:val="0409001D"/>'),
    );
    expect(rebuilt).toContain('<w:nsid w:val="1B2C3D4E"/>');
    expect(rebuilt).toContain('<w:tmpl w:val="0409001D"/>');
  });

  test("a level keeps its template code and its tentative flag", () => {
    const rebuilt = rebuild(lvlWith("", ' w:tplc="0409000F" w:tentative="1"'));
    expect(rebuilt).toContain('w:tplc="0409000F"');
    expect(rebuilt).toContain('w:tentative="1"');
  });

  test("a level keeps the paragraph style it numbers", () => {
    expect(rebuild(lvlWith('<w:pStyle w:val="Heading1"/>'))).toContain(
      '<w:pStyle w:val="Heading1"/>',
    );
  });

  test("a level keeps the picture bullet it draws from", () => {
    expect(rebuild(lvlWith('<w:lvlPicBulletId w:val="2"/>'))).toContain(
      '<w:lvlPicBulletId w:val="2"/>',
    );
  });

  test("a justification outside the three folio lays out is still carried", () => {
    // `both` is an `ST_Jc` member the marker has no layout for; refusing it
    // used to take the whole `w:lvlJc` with it.
    expect(rebuild(lvlWith('<w:lvlJc w:val="both"/>'))).toContain('<w:lvlJc w:val="both"/>');
  });

  test("an empty w:pPr is not an absent one", () => {
    expect(rebuild(lvlWith("<w:pPr/>"))).toContain("<w:pPr/>");
  });

  test("an empty w:rPr is not an absent one", () => {
    expect(rebuild(lvlWith("<w:rPr/>"))).toContain("<w:rPr/>");
  });

  test("an explicit w:isLgl of off stays off", () => {
    const rebuilt = rebuild(lvlWith('<w:isLgl w:val="0"/>'));
    expect(rebuilt).toContain("<w:isLgl");
    expect(parseNumbering(rebuilt).definitions.abstractNums.at(0)?.levels.at(0)?.isLgl).toBe(false);
  });

  test("w:legacy reads its own attribute rather than a w:val it has none of", () => {
    const rebuilt = rebuild(lvlWith('<w:legacy w:legacy="off"/>'));
    expect(rebuilt).toContain('w:legacy="0"');
  });

  test("a legacy measure spelled with its unit becomes the twips it counts", () => {
    expect(rebuild(lvlWith('<w:legacy w:legacySpace="72pt"/>'))).toContain('w:legacySpace="1440"');
  });

  test("w:lvlText keeps the null flag beside the text", () => {
    expect(rebuild(lvlWith('<w:lvlText w:val="" w:null="1"/>'))).toContain('w:null="1"');
  });

  test("a picture bullet keeps its place among the definitions", () => {
    const rebuilt = rebuild(
      partWith(
        '<w:numPicBullet w:numPicBulletId="0"><w:pict/></w:numPicBullet>' +
          '<w:abstractNum w:abstractNumId="1"/>',
      ),
    );
    expect(rebuilt).toContain("<w:numPicBullet");
    expect(rebuilt.indexOf("<w:numPicBullet")).toBeLessThan(rebuilt.indexOf("<w:abstractNum"));
  });

  test("the id high-water mark comes back after the definitions", () => {
    const rebuilt = rebuild(
      partWith('<w:abstractNum w:abstractNumId="1"/><w:numIdMacAtCleanup w:val="7"/>'),
    );
    expect(rebuilt).toContain("<w:numIdMacAtCleanup");
    expect(rebuilt.indexOf("<w:abstractNum")).toBeLessThan(rebuilt.indexOf("<w:numIdMacAtCleanup"));
  });

  test("an unmodelled attribute of w:abstractNum rides the element's remainder", () => {
    expect(rebuild(abstractNumWith("", ' w15:restartNumberingAfterBreak="0"'))).toContain(
      'w15:restartNumberingAfterBreak="0"',
    );
  });

  test("an unmodelled attribute of w:num rides the element's remainder", () => {
    expect(
      rebuild(numWith('<w:abstractNumId w:val="1"/>', ' w15:durableId="1451425873"')),
    ).toContain('w15:durableId="1451425873"');
  });

  test("an unmodelled attribute of w:lvlOverride rides the element's remainder", () => {
    expect(rebuild(lvlOverrideWith("", ' w14:unknown="3"'))).toContain('w14:unknown="3"');
  });

  test("a level whose ilvl names no level is still a definition the part carries", () => {
    const rebuilt = rebuild(
      partWith('<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="-1"/></w:abstractNum>'),
    );
    expect(rebuilt).toContain('w:ilvl="-1"');
    expect(parseNumbering(rebuilt).getAbstract(1)?.levels).toHaveLength(1);
    expect(parseNumbering(rebuilt).getLevel(1, -1)).toBeNull();
  });
});

/**
 * `mc:Ignorable` is stated by the rebuilt root, never replayed from the source.
 *
 * It lists which of the root's own `xmlns:*` bindings a consumer may skip, and
 * a rebuilt root binds the prefixes its body uses.
 */
describe("a part root states its ignorable prefixes rather than replaying them", () => {
  test("a source mc:Ignorable does not reach the rebuilt root", () => {
    expect(
      rebuild(partWith('<w:abstractNum w:abstractNumId="1"/>', ' mc:Ignorable="w14 w15"')),
    ).not.toContain('mc:Ignorable="w14 w15"');
  });

  test("another unmodelled root attribute does", () => {
    expect(
      rebuild(partWith('<w:abstractNum w:abstractNumId="1"/>', ' w:unknownRootFlag="1"')),
    ).toContain('w:unknownRootFlag="1"');
  });
});

/**
 * The model's number formats against the schema's, so the mirror cannot drift.
 *
 * `ST_NumberFormat` is the enumeration `w:numFmt@w:val` admits. folio models
 * every member but `custom`, which names a format string in `w:format` rather
 * than a format of its own, and adds three synthetic names for the pad widths
 * a custom zero-padded format resolves to.
 */
describe("the number formats folio models are the schema's", () => {
  const ST_NUMBER_FORMAT_VALUES: readonly string[] = (() => {
    const graph = JSON.parse(
      readFileSync(
        new URL(
          "../../../../specifications/generated/docx-transitional-schema.gen.json",
          import.meta.url,
        ),
        "utf8",
      ),
    ) as { symbols: { kind?: string; name?: string; namespace?: string; enumValues?: string[] }[] };
    const symbol = graph.symbols.find(
      (candidate) =>
        candidate.kind === "simpleType" &&
        candidate.name === "ST_NumberFormat" &&
        candidate.namespace === "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    );
    if (!symbol?.enumValues) {
      throw new Error("the committed schema graph declares no ST_NumberFormat");
    }
    return symbol.enumValues;
  })();

  /** Resolved through `w:format`, so it never reaches the model as a format. */
  const READER_OWNED = new Set(["custom"]);

  /** folio's names for the pad widths a `custom` zero-padded format resolves to. */
  const SYNTHETIC = new Set(["decimalZero3", "decimalZero4", "decimalZero5"]);

  test("every schema member but the reader-owned one is modelled", () => {
    const modelled = new Set<string>(NUMBER_FORMAT_VALUES);
    expect(
      ST_NUMBER_FORMAT_VALUES.filter((value) => !READER_OWNED.has(value) && !modelled.has(value)),
    ).toEqual([]);
  });

  test("no modelled format is outside the schema but the synthetic ones", () => {
    const schema = new Set(ST_NUMBER_FORMAT_VALUES);
    expect(
      NUMBER_FORMAT_VALUES.filter((value) => !schema.has(value) && !SYNTHETIC.has(value)),
    ).toEqual([]);
  });
});
