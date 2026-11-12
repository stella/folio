/**
 * `spliceXml` applies every splice of a part in one pass: it walks the splices
 * end to start, collects the kept source regions and the replacements, and
 * joins them once.
 *
 *   1. The result equals applying the splices one at a time to a rebuilt
 *      string, end to start (the per-splice algorithm, kept here only as the
 *      oracle), for every set of splices: non-overlapping ones in any order,
 *      empty, adjacent and at both ends of the part, and overlapping ones,
 *      which both refuse.
 *   2. Overlapping source regions are refused.
 *   3. A part is built once, however many splices it takes: no rebuilt string
 *      is sliced again, and the pieces are joined exactly once. The same probe
 *      run against the per-splice oracle reports a rebuild per splice.
 */

import { describe, expect, setDefaultTimeout, spyOn, test } from "bun:test";
import fc from "fast-check";

import {
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { patchBreaksCommentRangeBalance } from "./commentRangeIntegrity";
import { spliceXml, type XmlSplice } from "./selectiveXmlPatch";

setDefaultTimeout(propertyTestTimeout(30_000));

/** The per-splice algorithm `spliceXml` replaced: one rebuild per splice. */
const spliceOneAtATime = (
  xml: string,
  splices: readonly XmlSplice[],
): string | null => {
  let result = xml;
  let unpatchedEnd = xml.length;
  for (const { start, end, newXml } of [...splices].toSorted(
    (a, b) => b.start - a.start,
  )) {
    if (end > unpatchedEnd) return null;
    result = result.slice(0, start) + newXml + result.slice(end);
    unpatchedEnd = start;
  }
  return patchBreaksCommentRangeBalance(xml, result) ? null : result;
};

// Source and replacement text built from markup tokens, comment range halves
// among them, so the comment range refusal both algorithms share is reached.
const TOKENS = [
  "<w:p>",
  "</w:p>",
  "<w:r><w:t>text</w:t></w:r>",
  '<w:commentRangeStart w:id="1"/>',
  '<w:commentRangeEnd w:id="1"/>',
  '<w:commentRangeStart w:id="2"/>',
  '<w:commentRangeEnd w:id="2"/>',
  "a",
  "é",
  " ",
];

const markupArb = (maxLength: number): fc.Arbitrary<string> =>
  fc
    .array(fc.constantFrom(...TOKENS), { maxLength })
    .map((tokens) => tokens.join(""));

const replacementArb = fc.oneof(
  { weight: 1, arbitrary: fc.constant("") },
  { weight: 3, arbitrary: markupArb(3) },
);

/**
 * A part and splices that never overlap: cut points into the part, paired in
 * order, each pair either a replaced region or an insertion at one point. Equal
 * cut points give empty regions and adjacent splices; the part's two ends are
 * cut points as often as any other.
 */
const disjointCaseArb = markupArb(12).chain((xml) =>
  fc
    .array(
      fc.oneof(
        { weight: 1, arbitrary: fc.constant(0) },
        { weight: 1, arbitrary: fc.constant(xml.length) },
        { weight: 4, arbitrary: fc.nat({ max: xml.length }) },
      ),
      { maxLength: 12 },
    )
    .chain((cuts) => {
      const sorted = cuts.toSorted((left, right) => left - right);
      const regions: { start: number; end: number }[] = [];
      for (let index = 0; index + 1 < sorted.length; index += 2) {
        const [start, end] = [sorted[index], sorted[index + 1]];
        if (start !== undefined && end !== undefined)
          regions.push({ start, end });
      }
      return fc
        .tuple(
          fc.tuple(...regions.map(() => replacementArb)),
          fc.tuple(...regions.map(() => fc.boolean())),
        )
        .map(([replacements, insertOnly]) =>
          regions.map(({ start, end }, index) => ({
            start,
            end: insertOnly[index] ? start : end,
            newXml: replacements[index] ?? "",
          })),
        );
    })
    .chain((splices) =>
      fc
        .shuffledSubarray(splices, {
          minLength: splices.length,
          maxLength: splices.length,
        })
        .map((shuffled) => ({ xml, splices: shuffled })),
    ),
);

/** Any splices into the part: overlapping, nested, reversed, repeated. */
const anyCaseArb = markupArb(12).chain((xml) =>
  fc
    .array(
      fc.record({
        start: fc.nat({ max: xml.length }),
        end: fc.nat({ max: xml.length }),
        newXml: replacementArb,
      }),
      { maxLength: 8 },
    )
    .map((splices) => ({ xml, splices })),
);

/**
 * Two splices overlap when their source regions share a character, or when one
 * inserts at a point strictly inside the other's region. Splices that only
 * touch (an insertion at a region's edge, adjacent regions) do not.
 */
const overlaps = (splices: readonly XmlSplice[]): boolean =>
  splices.some((left, leftIndex) =>
    splices.some((right, rightIndex) => {
      if (leftIndex === rightIndex) return false;
      if (left.start > left.end || right.start > right.end) return false;
      const sharesCharacter =
        Math.max(left.start, right.start) < Math.min(left.end, right.end);
      const insertsInside =
        right.start === right.end &&
        left.start < right.start &&
        right.start < left.end;
      return sharesCharacter || insertsInside;
    }),
  );

describe("spliceXml", () => {
  test("non-overlapping splices in any order equal applying them one at a time", () => {
    fc.assert(
      fc.property(disjointCaseArb, ({ xml, splices }) => {
        expect(spliceXml(xml, splices)).toBe(spliceOneAtATime(xml, splices));
      }),
      propertyConfig({ numRuns: 600 }),
    );
  });

  test("any splices, overlapping ones included, equal applying them one at a time", () => {
    fc.assert(
      fc.property(anyCaseArb, ({ xml, splices }) => {
        expect(spliceXml(xml, splices)).toBe(spliceOneAtATime(xml, splices));
      }),
      propertyConfig({ numRuns: 600 }),
    );
  });

  test("overlapping source regions are refused", () => {
    fc.assert(
      fc.property(
        anyCaseArb.filter(({ splices }) => overlaps(splices)),
        ({ xml, splices }) => {
          expect(spliceXml(xml, splices)).toBeNull();
        },
      ),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("empty, adjacent and end splices, and an overlap", () => {
    const xml = "<w:p>abc</w:p>";
    expect(spliceXml(xml, [])).toBe(xml);
    expect(spliceXml(xml, [{ start: 5, end: 5, newXml: "" }])).toBe(xml);
    expect(
      spliceXml(xml, [
        { start: 6, end: 8, newXml: "Y" },
        { start: 5, end: 6, newXml: "X" },
      ]),
    ).toBe("<w:p>XY</w:p>");
    expect(
      spliceXml(xml, [
        { start: xml.length, end: xml.length, newXml: "!" },
        { start: 0, end: 0, newXml: "^" },
      ]),
    ).toBe("^<w:p>abc</w:p>!");
    expect(
      spliceXml(xml, [
        { start: 5, end: 7, newXml: "X" },
        { start: 6, end: 8, newXml: "Y" },
      ]),
    ).toBeNull();
  });
});

type BuildProbe = {
  /** Slices of a string other than the source part: rebuilt strings read again. */
  rebuiltSlices: number;
  /** `join` calls that assembled a string. */
  joins: number;
};

/**
 * Run `splice` while counting the strings it slices and the pieces it joins.
 * One build of the part is one join of pieces sliced from the source; every
 * slice of a string other than the source is a rebuilt part read again.
 */
const probeBuilds = (
  splice: (xml: string, splices: readonly XmlSplice[]) => string | null,
  xml: string,
  splices: readonly XmlSplice[],
): BuildProbe & { result: string | null } => {
  const slice = String.prototype.slice;
  let rebuiltSlices = 0;
  const sliceSpy = spyOn(String.prototype, "slice").mockImplementation(
    function (this: string, ...args: Parameters<typeof slice>) {
      if (String(this) !== xml) rebuiltSlices += 1;
      return slice.apply(this, args);
    },
  );
  const joinSpy = spyOn(Array.prototype, "join");
  try {
    const result = splice(xml, splices);
    return { rebuiltSlices, joins: joinSpy.mock.calls.length, result };
  } finally {
    sliceSpy.mockRestore();
    joinSpy.mockRestore();
  }
};

describe("spliceXml builds a part once", () => {
  // A part with a splice in every paragraph, as stamping paragraph ids does.
  const PARAGRAPHS = 50;
  const xml = Array.from(
    { length: PARAGRAPHS },
    (_unused, index) =>
      `<w:p><w:r><w:t>paragraph ${String(index)}</w:t></w:r></w:p>`,
  ).join("");
  const splices: XmlSplice[] = [];
  for (
    let at = xml.indexOf("<w:p>");
    at >= 0;
    at = xml.indexOf("<w:p>", at + 1)
  ) {
    splices.push({
      start: at,
      end: at + "<w:p>".length,
      newXml: '<w:p w14:paraId="1">',
    });
  }

  test("one join of source pieces, whatever the number of splices", () => {
    expect(splices).toHaveLength(PARAGRAPHS);
    const probe = probeBuilds(spliceXml, xml, splices);
    expect(probe.result).toBe(spliceOneAtATime(xml, splices));
    expect(probe.rebuiltSlices).toBe(0);
    expect(probe.joins).toBe(1);
  });

  test("the probe reports a rebuild per splice for the per-splice algorithm", () => {
    const probe = probeBuilds(spliceOneAtATime, xml, splices);
    expect(probe.joins).toBe(0);
    // The first splice slices the source; each later one slices the string the
    // previous splice built, twice.
    expect(probe.rebuiltSlices).toBe(2 * (PARAGRAPHS - 1));
  });
});
