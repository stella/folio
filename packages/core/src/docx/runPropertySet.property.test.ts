/**
 * A run property set survives a rebuild, in the order the schema declares.
 *
 * `w:rPr` was walked by a reader per property with no branch for the rest.
 * `w:bdr`, `w:fitText`, `w:eastAsianLayout`, `w:snapToGrid`, `w:webHidden` and
 * `w:oMath` had no model at all; `w:u`, `w:rFonts`, `w:lang` and `w:w` were
 * read and dropped whenever the reader took no typed value from them; and the
 * paragraph mark's whole `w:rPrChange` went with them. A save that rewrote the
 * element — which is every save after an edit — lost all of it.
 *
 * The universe is the generated declared-child list rather than a list
 * somebody kept: `Record<DeclaredChild<"run-properties">, string>` is total, so
 * a child the schema gains cannot reach the property test without a sample, and
 * the same list gives the order the assertions check.
 *
 * All three owners are exercised, because the point of one reader and one
 * writer is that they cannot answer differently: a run, the paragraph mark
 * inside `w:pPr`, and the snapshot inside a `w:rPrChange`.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";

import type { TextFormatting } from "../types/document";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import { CONTAINER_CHILDREN, type DeclaredChild } from "./containerChildren.gen";
import { parseParagraph } from "./paragraphParser";
import { parseRun, parseRunProperties, RUN_PROPERTY_OWNERS } from "./runParser";
import { serializeParagraphFormatting } from "./serializer/paragraphSerializer";
import { serializeRun } from "./serializer/runSerializer";
import { serializeTextFormatting } from "./serializer/textFormattingSerializer";
import { parseXmlDocument, type XmlElement } from "./xmlParser";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const DECLARED = CONTAINER_CHILDREN["run-properties"];

/**
 * One authored instance per declared child, each stating something.
 *
 * `satisfies Record<DeclaredChild<…>, string>` is what makes this total: a
 * child added to `EG_RPrBase` fails the compile here before it can be dropped
 * silently at run time.
 */
const SAMPLES = {
  ins: '<w:ins w:id="11" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>',
  del: '<w:del w:id="12" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>',
  moveFrom: '<w:moveFrom w:id="13" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>',
  moveTo: '<w:moveTo w:id="14" w:author="Reviewer" w:date="2026-05-15T12:00:00Z"/>',
  rStyle: '<w:rStyle w:val="Emphasis"/>',
  rFonts: '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>',
  b: "<w:b/>",
  bCs: "<w:bCs/>",
  i: "<w:i/>",
  iCs: "<w:iCs/>",
  caps: "<w:caps/>",
  smallCaps: "<w:smallCaps/>",
  strike: "<w:strike/>",
  dstrike: "<w:dstrike/>",
  outline: "<w:outline/>",
  shadow: "<w:shadow/>",
  emboss: "<w:emboss/>",
  imprint: "<w:imprint/>",
  noProof: "<w:noProof/>",
  snapToGrid: '<w:snapToGrid w:val="0"/>',
  vanish: "<w:vanish/>",
  webHidden: "<w:webHidden/>",
  color: '<w:color w:val="1F4E79"/>',
  spacing: '<w:spacing w:val="20"/>',
  w: '<w:w w:val="110"/>',
  kern: '<w:kern w:val="18"/>',
  position: '<w:position w:val="6"/>',
  sz: '<w:sz w:val="22"/>',
  szCs: '<w:szCs w:val="22"/>',
  highlight: '<w:highlight w:val="yellow"/>',
  u: '<w:u w:val="single"/>',
  effect: '<w:effect w:val="shimmer"/>',
  bdr: '<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>',
  shd: '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>',
  fitText: '<w:fitText w:val="1440" w:id="3"/>',
  vertAlign: '<w:vertAlign w:val="superscript"/>',
  rtl: "<w:rtl/>",
  cs: "<w:cs/>",
  em: '<w:em w:val="dot"/>',
  lang: '<w:lang w:val="cs-CZ" w:eastAsia="ja-JP"/>',
  eastAsianLayout: '<w:eastAsianLayout w:id="5" w:combine="1"/>',
  specVanish: "<w:specVanish/>",
  oMath: "<w:oMath/>",
  rPrChange:
    '<w:rPrChange w:id="7" w:author="Reviewer" w:date="2026-05-15T12:00:00Z">' +
    '<w:rPr><w:b/><w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:rPr>' +
    "</w:rPrChange>",
} as const satisfies Record<DeclaredChild<"run-properties">, string>;

/**
 * Children the reader states nothing about, one per way of stating nothing.
 *
 * An element with no attributes at all, one whose value the reader's
 * enumeration does not admit, and one from a namespace the content model does
 * not name. None of the three can be decided by a map keyed on the child's
 * name, which is why the handler answers with what it took.
 */
const UNREAD_CHILDREN = [
  "<w:rFonts/>",
  "<w:lang/>",
  "<w:sz/>",
  '<w:u w:val="squiggle"/>',
  '<w:highlight w:val="chartreuse"/>',
  '<w:vertAlign w:val="middle"/>',
  '<x:hint xmlns:x="urn:example:vendor" x:kind="type"/>',
] as const;

const parseOne = (xml: string): XmlElement => {
  const root = parseXmlDocument(xml) as XmlElement | null;
  if (!root) {
    throw new Error("the fixture did not parse");
  }
  return root;
};

/** A run's own `w:rPr`, parsed and written back from the model alone. */
const rebuildRun = (properties: string): string => {
  const run = parseRun(
    parseOne(`<w:r xmlns:w="${W}"><w:rPr>${properties}</w:rPr><w:t>x</w:t></w:r>`),
    null,
    null,
  );
  return serializeRun(run);
};

/**
 * The paragraph mark's `w:rPr`, parsed and written back from the model alone.
 *
 * `serializeParagraphFormatting` is called rather than `serializeParagraph`
 * because the paragraph replays its captured `w:pPr` whenever the model still
 * agrees with it; going through the formatting serializer is the same forcing
 * the survival law applies, and the only way this test can see the rebuild.
 */
const rebuildParagraphMark = (properties: string): string => {
  const paragraph = parseParagraph(
    parseOne(`<w:p xmlns:w="${W}"><w:pPr><w:rPr>${properties}</w:rPr></w:pPr></w:p>`),
    null,
    null,
    null,
  );
  return serializeParagraphFormatting(
    paragraph.formatting,
    paragraph.propertyChanges,
    paragraph.pPrMark,
  );
};

/** A `w:rPrChange` snapshot, parsed and written back from the model alone. */
const rebuildChangeOriginal = (properties: string): string =>
  serializeTextFormatting(
    parseRunProperties(
      parseOne(`<w:rPr xmlns:w="${W}">${properties}</w:rPr>`),
      null,
      RUN_PROPERTY_OWNERS.standalone,
    ),
  );

/**
 * Where a child starts in the saved markup, or `-1`.
 *
 * The delimiter matters: `w:b` is a prefix of `w:bCs` and `w:bdr`, and `w:w`
 * of `w:webHidden`, so a plain `indexOf` reports one child at another's place
 * and an order assertion built on it proves nothing.
 */
const childAt = (saved: string, name: string): number =>
  saved.search(new RegExp(`<w:${name}[ />]`, "u"));

/** The outer `w:rPr`'s own children: a `w:rPrChange` nests a second one. */
const propertiesOf = (saved: string): string =>
  saved.slice(saved.indexOf("<w:rPr>") + "<w:rPr>".length, saved.lastIndexOf("</w:rPr>"));

describe("a run property set survives a rebuild", () => {
  test("every declared child comes back on a run, and the save after it is a fixed point", () => {
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        const saved = rebuildRun(SAMPLES[name]);

        expect(childAt(saved, name)).toBeGreaterThan(-1);
        expect(rebuildRun(propertiesOf(saved))).toBe(saved);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("every declared child comes back on the paragraph mark", () => {
    fc.assert(
      fc.property(fc.constantFrom(...DECLARED), (name) => {
        const saved = rebuildParagraphMark(SAMPLES[name]);

        expect(childAt(saved, name)).toBeGreaterThan(-1);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("a child the reader takes nothing from keeps its bytes", () => {
    fc.assert(
      fc.property(
        fc.constantFrom(...UNREAD_CHILDREN),
        fc.constantFrom(rebuildRun, rebuildChangeOriginal),
        (child, rebuild) => {
          const saved = rebuild(`<w:b/>${child}`);

          expect(saved).toContain(child);
          expect(rebuild(propertiesOf(saved))).toBe(saved);
        },
      ),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("the children come back in the order the content model declares", () => {
    // Authored backwards: a serializer that wrote them in the order it read
    // them, or in the order of its own statements, would come back out of
    // sequence and a validating consumer would refuse the part. `w:noProof`
    // before `w:vanish` is the case the hand-kept order had wrong.
    //
    // The run is the owner with no mutually exclusive children: a paragraph
    // mark carries one of `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo` and never
    // four, so only a run can hold one of everything at once.
    const saved = rebuildRun(
      [...DECLARED]
        .reverse()
        .map((name) => SAMPLES[name])
        .join(""),
    );

    const positions = DECLARED.map((name) => childAt(saved, name));
    expect(positions.every((at) => at > -1)).toBe(true);
    expect(positions).toEqual([...positions].sort((left, right) => left - right));
  });

  test("a capture comes back between the same two modelled properties", () => {
    const saved = rebuildRun('<w:b/><w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>');

    expect(childAt(saved, "b")).toBeGreaterThan(-1);
    expect(childAt(saved, "b")).toBeLessThan(childAt(saved, "bdr"));
  });

  test("the snapshot inside a tracked property change keeps what no reader took", () => {
    const saved = rebuildRun(`<w:i/>${SAMPLES.rPrChange}`);

    expect(saved).toContain('<w:rPrChange w:id="7" w:author="Reviewer"');
    expect(saved).toContain('<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    expect(rebuildRun(propertiesOf(saved))).toBe(saved);
  });

  test("a tracked property change survives a run that states no properties of its own", () => {
    const saved = rebuildRun(SAMPLES.rPrChange);

    expect(saved).toContain('<w:rPrChange w:id="7" w:author="Reviewer"');
  });

  test("captured bytes belong to the element that was parsed and are never inherited", () => {
    // #873: a style-resolved value may not become direct formatting. The sink
    // is the strongest case of that rule, because it is not a value at all —
    // it is the bytes one `w:rPr` held, and a merge that carried them would
    // write a style's or a paragraph mark's markup into every run below it.
    const inherited = parseRunProperties(
      parseOne(`<w:rPr xmlns:w="${W}"><w:bdr w:val="single" w:sz="4"/></w:rPr>`),
      null,
      RUN_PROPERTY_OWNERS.standalone,
    );
    expect(inherited?.preserved).toBeDefined();

    const direct: TextFormatting = { bold: true };
    expect(mergeTextFormatting(inherited, direct)?.preserved).toBeUndefined();
    expect(mergeTextFormatting(direct, inherited)?.preserved).toBeUndefined();
    expect(serializeTextFormatting(mergeTextFormatting(inherited, direct))).not.toContain("w:bdr");
  });
});
