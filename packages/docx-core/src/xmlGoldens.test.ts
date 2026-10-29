import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { resolve } from "node:path";

import { scanTestTimeout } from "../../../test/scan-timeout";
import {
  attempt,
  attemptAsync,
  expectEdgesMatch,
  expectFixturesMatch,
  readGoldens,
  refusedWhen,
  SAMPLE_FIXTURE,
  withDocument,
} from "../../../test/xmlGoldens";
import {
  XML_EDGE_FRAGMENTS,
  XML_EDGE_INPUTS,
} from "../../../tests/support/xmlEdgeInputs";
import {
  answerDigest,
  readFixturePackages,
} from "../../../tests/support/xmlFixtureParts";
import { isSafePreservedChildXml } from "./schema/sequenceChildren";
import { assertValidOoxmlNumericIds } from "./serialize/numericIds";
import { validateDocxPackage } from "./validate/docx";
import { validatePackageGraph } from "./validate/packageGraph";
import { validateSchemaAttributes } from "./validate/schemaAttributes";
import {
  getXmlSourceRange,
  parseStreamingXmlWithSourceRanges,
  type XmlElement,
} from "./xml/streamingXmlParser";

setDefaultTimeout(scanTestTimeout(120_000));

type Site =
  | "validateSchemaAttributes"
  | "validatePackageGraph"
  | "validateDocxPackage"
  | "assertValidOoxmlNumericIds"
  | "isSafePreservedChildXml";

const GOLDENS = readGoldens<Site>(
  resolve(import.meta.dir, "__fixtures__/xml-goldens.json"),
);
const packages = await readFixturePackages();
const base = packages.find(({ path }) => path === SAMPLE_FIXTURE);
if (base === undefined) throw new Error(`missing ${SAMPLE_FIXTURE}`);

/** Up to 300 element sources of at most 4 KiB per part, in document order. */
const sourceSlices = (xml: string): string[] => {
  const parsed = parseStreamingXmlWithSourceRanges(xml);
  if (parsed.status !== "parsed") return [];
  const slices: string[] = [];
  const visit = (element: XmlElement): void => {
    if (slices.length >= 300) return;
    if (element.type === "element") {
      const range = getXmlSourceRange(element);
      if (range && range.end - range.start <= 4096)
        slices.push(xml.slice(range.start, range.end));
    }
    for (const child of element.elements ?? []) visit(child);
  };
  visit(parsed.value);
  return slices;
};

const perPart = (
  answer: (xml: string, part: string) => unknown,
): Record<string, string> =>
  Object.fromEntries(
    packages.map(({ path, parts }) => [
      path,
      answerDigest([...parts].map(([part, xml]) => [part, answer(xml, part)])),
    ]),
  );

/** Markup no XML reader accepts, which the earlier validator did not look for. */
const NOT_WELL_FORMED = "not well-formed XML 1.0";
const DEPTH = "nesting is bounded at the reader's resource limit";

describe("validateSchemaAttributes", () => {
  const refused = refusedWhen((value) => value !== null);
  test("answers every fixture part as before", () => {
    expectFixturesMatch(
      GOLDENS.validateSchemaAttributes,
      perPart((xml) => attempt(() => validateSchemaAttributes(xml))),
    );
  });
  test("refuses what it refused", () => {
    expectEdgesMatch({
      golden: GOLDENS.validateSchemaAttributes,
      answers: Object.fromEntries(
        XML_EDGE_INPUTS.map(([name, xml]) => [
          name,
          attempt(() => validateSchemaAttributes(xml)),
        ]),
      ),
      refused,
      tightened: {
        "less-than in attribute": NOT_WELL_FORMED,
        "raw control character": NOT_WELL_FORMED,
        "CDATA terminator in text": NOT_WELL_FORMED,
        "double hyphen in comment": NOT_WELL_FORMED,
        "declaration inside the root": NOT_WELL_FORMED,
        "depth 101": DEPTH,
      },
    });
  });
});

describe("validatePackageGraph", () => {
  const refused = refusedWhen((value) => value !== null);
  test("answers every fixture package as before", () => {
    expectFixturesMatch(
      GOLDENS.validatePackageGraph,
      Object.fromEntries(
        packages.map(({ path, parts, paths }) => [
          path,
          attempt(() => validatePackageGraph(parts, paths)),
        ]),
      ),
    );
  });
  test("refuses what it refused", () => {
    expectEdgesMatch({
      golden: GOLDENS.validatePackageGraph,
      answers: Object.fromEntries(
        XML_EDGE_INPUTS.map(([name, xml]) => {
          const parts = new Map(base.parts);
          parts.set("word/document.xml", xml);
          return [name, attempt(() => validatePackageGraph(parts, base.paths))];
        }),
      ),
      refused,
      tightened: {
        "two roots": NOT_WELL_FORMED,
        "less-than in attribute": NOT_WELL_FORMED,
        "raw control character": NOT_WELL_FORMED,
        "CDATA terminator in text": NOT_WELL_FORMED,
        "double hyphen in comment": NOT_WELL_FORMED,
        "declaration inside the root": NOT_WELL_FORMED,
        "depth 101": DEPTH,
      },
    });
  });
});

describe("validateDocxPackage", () => {
  const refused = refusedWhen(
    (value) =>
      typeof value === "object" &&
      value !== null &&
      "valid" in value &&
      value.valid === false,
  );
  test("answers every fixture package as before", async () => {
    const answers: Record<string, unknown> = {};
    for (const { path, bytes } of packages) {
      answers[path] = await attemptAsync(() => validateDocxPackage(bytes));
    }
    expectFixturesMatch(GOLDENS.validateDocxPackage, answers);
  });
  test("refuses what it refused", async () => {
    const answers: Record<string, unknown> = {};
    for (const [name, xml] of XML_EDGE_INPUTS) {
      answers[name] = await attemptAsync(async () =>
        validateDocxPackage(await withDocument(base, xml)),
      );
    }
    expectEdgesMatch({
      golden: GOLDENS.validateDocxPackage,
      answers,
      refused,
    });
  });
});

describe("assertValidOoxmlNumericIds", () => {
  test("answers every fixture part as before", () => {
    expectFixturesMatch(
      GOLDENS.assertValidOoxmlNumericIds,
      perPart((xml, part) =>
        attempt(() => assertValidOoxmlNumericIds(xml, part)),
      ),
    );
  });
  test("answers every edge input as before", () => {
    expectEdgesMatch({
      golden: GOLDENS.assertValidOoxmlNumericIds,
      answers: Object.fromEntries(
        XML_EDGE_INPUTS.map(([name, xml]) => [
          name,
          attempt(() => assertValidOoxmlNumericIds(xml, "word/document.xml")),
        ]),
      ),
      refused: refusedWhen(() => false),
    });
  });
});

describe("isSafePreservedChildXml", () => {
  test("answers every fixture element as before", () => {
    expectFixturesMatch(
      GOLDENS.isSafePreservedChildXml,
      perPart((xml) => sourceSlices(xml).map(isSafePreservedChildXml)),
    );
    expect(
      packages.some(({ parts }) =>
        [...parts.values()].some((xml) =>
          sourceSlices(xml).some(isSafePreservedChildXml),
        ),
      ),
    ).toBe(true);
  });
  test("answers every edge fragment as before", () => {
    expectEdgesMatch({
      golden: GOLDENS.isSafePreservedChildXml,
      answers: Object.fromEntries(
        XML_EDGE_FRAGMENTS.map(([name, xml]) => [
          name,
          attempt(() => isSafePreservedChildXml(xml)),
        ]),
      ),
      refused: refusedWhen((value) => value === false),
    });
  });
});
