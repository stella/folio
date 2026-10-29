import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { resolve } from "node:path";

import { scanTestTimeout } from "../../../../test/scan-timeout";
import {
  attempt,
  expectEdgesMatch,
  expectFixturesMatch,
  NEVER,
  readGoldens,
  refusedWhen,
} from "../../../../test/xmlGoldens";
import {
  XML_EDGE_FRAGMENTS,
  XML_EDGE_INPUTS,
} from "../../../../tests/support/xmlEdgeInputs";
import {
  answerDigest,
  describeXmlTree,
  readFixturePackages,
} from "../../../../tests/support/xmlFixtureParts";
import { parseXml } from "./xmlParser";
import { getDocxXmlSafetyIssue } from "./xmlSafety";

setDefaultTimeout(scanTestTimeout(60_000));

const GOLDENS = readGoldens<"parseXml" | "safety">(
  resolve(import.meta.dir, "__fixtures__/xml-goldens.json"),
);
const packages = await readFixturePackages();
const EDGES = [...XML_EDGE_INPUTS, ...XML_EDGE_FRAGMENTS];

/** Markup no XML reader accepts; earlier releases repaired it into some tree. */
const REPAIRED_BEFORE =
  "earlier releases rebuilt a tree from broken markup; it is refused now";

describe("parseXml", () => {
  test("reads every fixture part as before", () => {
    expect(packages.length).toBeGreaterThan(30);
    expectFixturesMatch(
      GOLDENS.parseXml,
      Object.fromEntries(
        packages.map(({ path, parts }) => [
          path,
          answerDigest(
            [...parts].map(([part, xml]) => [
              part,
              attempt(() => describeXmlTree(parseXml(xml))),
            ]),
          ),
        ]),
      ),
    );
  });

  test("refuses what it refused, reads the rest as before", () => {
    const answers = Object.fromEntries(
      EDGES.map(([name, xml]) => {
        const answer = attempt(() => describeXmlTree(parseXml(xml)));
        return [
          name,
          "value" in answer ? { digest: answerDigest(answer.value) } : answer,
        ];
      }),
    );
    expectEdgesMatch({
      golden: GOLDENS.parseXml,
      answers,
      refused: refusedWhen(NEVER),
      tightened: {
        "mismatched end tag": REPAIRED_BEFORE,
        "unclosed element": REPAIRED_BEFORE,
        "stray end tag": REPAIRED_BEFORE,
        "attribute without value": REPAIRED_BEFORE,
        "unquoted attribute": REPAIRED_BEFORE,
        mismatched: REPAIRED_BEFORE,
        unclosed: REPAIRED_BEFORE,
        DOCTYPE: "a DOCTYPE is refused; DOCX parts never declare one",
        "DOCTYPE entity expansion":
          "document-declared entities are refused rather than expanded",
        ENTITY: "markup declarations are refused",
        "depth 101":
          "nesting is bounded at the resource limit the package reader applies",
      },
      reread: {
        "text before root":
          "text outside the root element is dropped, whitespace included",
        "CDATA only": "text outside the root element is dropped",
      },
    });
  });
});

describe("getDocxXmlSafetyIssue", () => {
  test("judges every fixture part as before", () => {
    expectFixturesMatch(
      GOLDENS.safety,
      Object.fromEntries(
        packages.map(({ path, parts }) => [
          path,
          answerDigest(
            [...parts].map(([part, xml]) => [part, getDocxXmlSafetyIssue(xml)]),
          ),
        ]),
      ),
    );
  });

  test("refuses what it refused, and markup no XML reader accepts", () => {
    const notWellFormed =
      "not well-formed XML; the earlier validator did not look for it";
    expectEdgesMatch({
      golden: GOLDENS.safety,
      answers: Object.fromEntries(
        EDGES.map(([name, xml]) => [name, getDocxXmlSafetyIssue(xml)]),
      ),
      refused: (issue) => issue !== null,
      tightened: {
        "two roots": notWellFormed,
        "two elements": notWellFormed,
        "element and text": notWellFormed,
        "prototype attribute": "names the reader reserves are refused",
        "constructor element": "names the reader reserves are refused",
        "constructor attribute": "names the reader reserves are refused",
        "depth 101": "nesting is bounded at the reader's resource limit",
        "depth 150": "nesting is bounded at the reader's resource limit",
      },
    });
  });
});
