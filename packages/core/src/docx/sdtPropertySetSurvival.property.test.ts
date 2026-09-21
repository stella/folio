/**
 * `w:sdtPr` keeps what its author wrote.
 *
 * The handler map is total over `CT_SdtPr`'s declared children by
 * construction — the compiler refuses a partial one — so what needs a test is
 * the half a type cannot state: that the parse/save pair is the identity on
 * an arbitrary property set in schema order, that a *value* the reader refuses
 * keeps its element, and that an extension child the Transitional graph does
 * not declare survives too.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { propertyTestTimeout } from "../../../../test/property-testing";

import { CONTAINER_CHILDREN } from "./containerChildren.gen";
import { parseSdtProperties } from "./sdtProperties";
import { serializeSdtProperties } from "./serializer/sdtPropertiesSerializer";
import { parseXml } from "./xmlParser";

setDefaultTimeout(propertyTestTimeout(30_000));

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
  'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"',
].join(" ");

const parse = (body: string) => {
  const root = parseXml(`<w:sdtPr ${NS}>${body}</w:sdtPr>`);
  const sdtPr = root.elements?.[0];
  if (!sdtPr) {
    throw new TypeError("expected a w:sdtPr root");
  }
  return parseSdtProperties(sdtPr);
};

/** The markup folio writes back, with the namespace declarations dropped. */
const roundTrip = (body: string): string =>
  serializeSdtProperties(parse(body)).replaceAll(/ xmlns:[\w-]+="[^"]*"/gu, "");

/**
 * One authored spelling per declared child of `CT_SdtPr`, keyed by name so
 * the sample is total over the generated set rather than over a list beside
 * it: a child that joins the schema and not this table fails the test below.
 */
const AUTHORED = {
  rPr: '<w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr>',
  alias: '<w:alias w:val="Client name"/>',
  tag: '<w:tag w:val="client"/>',
  id: '<w:id w:val="1734"/>',
  lock: '<w:lock w:val="sdtContentLocked"/>',
  placeholder: '<w:placeholder><w:docPart w:val="DefaultPlaceholder"/></w:placeholder>',
  temporary: "<w:temporary/>",
  showingPlcHdr: "<w:showingPlcHdr/>",
  // The three attributes of `CT_DataBinding`. `w:prefixMappings` carries
  // quotes of its own, which the writer spells as entities, so it is written
  // here the way folio writes it back.
  dataBinding:
    '<w:dataBinding w:prefixMappings="xmlns:ns0=&apos;urn:x&apos;" w:xpath="/ns0:root[1]/ns0:a[1]" w:storeItemID="{DEAD-BEEF}"/>',
  label: '<w:label w:val="17"/>',
  tabIndex: '<w:tabIndex w:val="3"/>',
  equation: "<w:equation/>",
  comboBox:
    '<w:comboBox w:lastValue="b"><w:listItem w:displayText="A" w:value="a"/><w:listItem w:displayText="B" w:value="b"/></w:comboBox>',
  date: '<w:date w:fullDate="2026-06-02T00:00:00Z"><w:dateFormat w:val="d MMMM yyyy"/><w:lid w:val="cs-CZ"/><w:storeMappedDataAs w:val="dateTime"/><w:calendar w:val="gregorian"/></w:date>',
  docPartObj:
    '<w:docPartObj><w:docPartGallery w:val="Quick Parts"/><w:docPartCategory w:val="General"/><w:docPartUnique/></w:docPartObj>',
  docPartList:
    '<w:docPartList><w:docPartGallery w:val="Bibliographies"/><w:docPartCategory w:val="General"/><w:docPartUnique w:val="0"/></w:docPartList>',
  dropDownList:
    '<w:dropDownList w:lastValue="a"><w:listItem w:displayText="A" w:value="a"/></w:dropDownList>',
  picture: "<w:picture/>",
  richText: "<w:richText/>",
  text: '<w:text w:multiLine="1"/>',
  citation: "<w:citation/>",
  group: "<w:group/>",
  bibliography: "<w:bibliography/>",
} as const satisfies Record<
  (typeof CONTAINER_CHILDREN)["content-control-properties"][number],
  string
>;

/** The kind elements are a choice, so a sample may hold at most one. */
const KIND_CHILDREN = [
  "equation",
  "comboBox",
  "date",
  "docPartObj",
  "docPartList",
  "dropDownList",
  "picture",
  "richText",
  "text",
  "citation",
  "bibliography",
  "group",
] as const;

const SEQUENCE = CONTAINER_CHILDREN["content-control-properties"];

/** A property set: any subset of the declared children, in schema order. */
const propertySet = fc
  .record({
    subset: fc.subarray(SEQUENCE.filter((name) => !KIND_CHILDREN.some((kind) => kind === name))),
    kind: fc.option(fc.constantFrom(...KIND_CHILDREN), { nil: undefined }),
  })
  .map(({ subset, kind }) =>
    [...subset, ...(kind === undefined ? [] : [kind])]
      .toSorted((left, right) => SEQUENCE.indexOf(left) - SEQUENCE.indexOf(right))
      .map((name) => AUTHORED[name])
      .join(""),
  );

describe("w:sdtPr survives a rebuild", () => {
  test("parse → save is the identity on any subset of CT_SdtPr, in schema order", () => {
    fc.assert(
      fc.property(propertySet, (body) => {
        expect(roundTrip(body)).toBe(
          body.length === 0 ? "<w:sdtPr/>" : `<w:sdtPr>${body}</w:sdtPr>`,
        );
      }),
      { numRuns: 300 },
    );
  });

  test("parse(save(parse(x))) = parse(x)", () => {
    fc.assert(
      fc.property(propertySet, (body) => {
        expect(parse(roundTrip(body).slice("<w:sdtPr>".length, -"</w:sdtPr>".length))).toEqual(
          parse(body),
        );
      }),
      { numRuns: 300 },
    );
  });

  test("an extension child the Transitional graph does not declare survives", () => {
    // `w14:checkbox`, `w15:appearance`, `w15:color` and `w15:repeatingSection`
    // are not in the content model at all, so nothing but the sink can keep
    // them, and the sink is where an undeclared name goes by default.
    const body =
      '<w:id w:val="9"/><w14:checkbox><w14:checked w14:val="1"/><w14:checkedState w14:val="2612" w14:font="MS Gothic"/></w14:checkbox><w15:appearance w15:val="hidden"/><w15:color w15:val="4472C4"/><w15:repeatingSection/>';
    const props = parse(body);
    expect(props.sdtType).toBe("checkbox");
    expect(props.checked).toBe(true);
    expect(roundTrip(body)).toBe(`<w:sdtPr>${body}</w:sdtPr>`);
  });

  test("a value the reader refuses keeps its element rather than changing meaning", () => {
    // `ST_Lock` has four values. A fifth used to be read as `unlocked`, which
    // is a silent change of meaning, and is the case a name-keyed map cannot
    // express: the decision is the reader's outcome, not the child's name.
    const props = parse('<w:lock w:val="sdtUnlocked"/>');
    expect(props.lock).toBeUndefined();
    expect(roundTrip('<w:lock w:val="sdtUnlocked"/>')).toBe(
      '<w:sdtPr><w:lock w:val="sdtUnlocked"/></w:sdtPr>',
    );
    // `<w:id/>` states no number, and `<w:alias/>` no name.
    expect(roundTrip("<w:id/>")).toBe("<w:sdtPr><w:id/></w:sdtPr>");
    expect(roundTrip("<w:alias/>")).toBe("<w:sdtPr><w:alias/></w:sdtPr>");
  });

  test("an empty property set is written as an empty one, not an absent one", () => {
    expect(roundTrip("")).toBe("<w:sdtPr/>");
  });

  test("showingPlcHdr is tri-state: absent, on, off", () => {
    expect(parse("").showingPlaceholder).toBeUndefined();
    expect(parse("<w:showingPlcHdr/>").showingPlaceholder).toBe(true);
    expect(parse('<w:showingPlcHdr w:val="0"/>').showingPlaceholder).toBe(false);
    expect(roundTrip('<w:showingPlcHdr w:val="0"/>')).toBe(
      '<w:sdtPr><w:showingPlcHdr w:val="0"/></w:sdtPr>',
    );
  });
});
