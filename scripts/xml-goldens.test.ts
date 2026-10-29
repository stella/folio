import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { resolve } from "node:path";

import { readXml } from "@stll/docx-core/xml";

import { scanTestTimeout } from "../test/scan-timeout";
import {
  attempt,
  attemptAsync,
  expectEdgesMatch,
  expectFixturesMatch,
  NEVER,
  readGoldens,
  refusedWhen,
  SAMPLE_FIXTURE,
  withDocument,
} from "../test/xmlGoldens";
import { XML_EDGE_INPUTS } from "../tests/support/xmlEdgeInputs";
import {
  answerDigest,
  describeXmlTree,
  readFixturePackages,
} from "../tests/support/xmlFixtureParts";
import { classifyCorpusFile } from "./lib/corpus-classify";
import {
  loadSchemaGraph,
  validateOoxmlPart,
} from "./lib/corpus-schema-validator";
import { prepareXmlPruning } from "./lib/corpus-xml-prune";

setDefaultTimeout(scanTestTimeout(300_000));

const GOLDENS = readGoldens<
  "prepareXmlPruning" | "validateOoxmlPart" | "classifyCorpusFile"
>(resolve(import.meta.dir, "lib/__fixtures__/xml-goldens.json"));
const packages = await readFixturePackages();
const base = packages.find(({ path }) => path === SAMPLE_FIXTURE);
if (base === undefined) throw new Error(`missing ${SAMPLE_FIXTURE}`);

const perPart = (answer: (xml: string) => unknown): Record<string, string> =>
  Object.fromEntries(
    packages.map(({ path, parts }) => [
      path,
      answerDigest([...parts].map(([part, xml]) => [part, answer(xml)])),
    ]),
  );

const REPAIRED_BEFORE =
  "the earlier parser rebuilt a tree from broken markup; it is refused now";
const RESERVED = "names the reader reserves are refused";
const NOT_WELL_FORMED = "not well-formed XML 1.0";

const pruneAnswer = (xml: string) => {
  const prunable = prepareXmlPruning(xml);
  const all = new Set(prunable.addresses);
  const even = new Set(
    prunable.addresses.filter((address) => address % 2 === 0),
  );
  return {
    count: prunable.addresses.length,
    all: prunable.render(all),
    even: prunable.render(even),
  };
};

/** The rebuild writes what it read: the root element, read again, is unchanged. */
const rootTree = (xml: string): string => {
  const read = readXml(xml, { conformance: "tolerant" });
  expect(read.status).toBe("parsed");
  return read.status === "parsed"
    ? describeXmlTree(
        read.value.elements?.find(({ type }) => type === "element") ?? {},
      )
    : "";
};

/**
 * The rebuild keeps a declaration as written, drops a byte-order mark beside
 * the root, and writes a character reference as the character it names
 * rather than escaping its ampersand into literal text.
 */
const REWRITTEN =
  "the rebuild keeps the declaration as written and decodes references";

describe("prepareXmlPruning", () => {
  test("a rebuild with nothing removed reads back as the part", () => {
    for (const { path, parts } of packages) {
      for (const [part, xml] of parts) {
        const prunable = prepareXmlPruning(xml);
        expect(
          rootTree(prunable.render(new Set(prunable.addresses))),
          `${path}:${part}`,
        ).toBe(rootTree(xml));
      }
    }
  });

  test("rebuilds every fixture part as before, but where it now reads correctly", () => {
    expectFixturesMatch(
      GOLDENS.prepareXmlPruning,
      perPart((xml) => attempt(() => pruneAnswer(xml))),
      {
        "packages/core/src/docx/__tests__/__fixtures__/corpus/checkbox-val-false.docx":
          REWRITTEN,
        "packages/core/src/docx/__tests__/__fixtures__/corpus/checkbox-val-true.docx":
          REWRITTEN,
        "packages/core/src/docx/__tests__/__fixtures__/corpus/inline-sdt-checkbox.docx":
          REWRITTEN,
        "packages/core/src/docx/__tests__/__fixtures__/regressions/repack-paragraph-sectpr.docx":
          REWRITTEN,
        [SAMPLE_FIXTURE]: REWRITTEN,
      },
    );
  });

  test("refuses what it refused", () => {
    const answers = Object.fromEntries(
      XML_EDGE_INPUTS.map(([name, xml]) => {
        const answer = attempt(() => pruneAnswer(xml));
        return [
          name,
          "value" in answer ? { digest: answerDigest(answer.value) } : answer,
        ];
      }),
    );
    expectEdgesMatch({
      golden: GOLDENS.prepareXmlPruning,
      answers,
      refused: refusedWhen(NEVER),
      tightened: {
        "mismatched end tag": REPAIRED_BEFORE,
        "unclosed element": REPAIRED_BEFORE,
        "stray end tag": REPAIRED_BEFORE,
        "attribute without value": REPAIRED_BEFORE,
        "unquoted attribute": REPAIRED_BEFORE,
        DOCTYPE: "a DOCTYPE is refused",
        "DOCTYPE entity expansion": "a DOCTYPE is refused",
        "prototype attribute": RESERVED,
        "constructor attribute": RESERVED,
      },
      reread: Object.fromEntries(
        [
          "character references",
          "CRLF and tab attribute",
          "comment and PI",
          "text before root",
          "HTML entity in text",
          "HTML entity in attribute",
          "raw control character",
          "declaration after whitespace",
          "byte-order mark",
          "space before tag name",
          "space after self-closing slash",
          "numeric identifier with reference",
          "numeric identifier with HTML entity",
        ].map((name) => [name, REWRITTEN]),
      ),
    });
  });
});

describe("validateOoxmlPart", () => {
  test("reports every fixture part as before", async () => {
    const graph = await loadSchemaGraph();
    expectFixturesMatch(
      GOLDENS.validateOoxmlPart,
      perPart((xml) => attempt(() => validateOoxmlPart({ graph, xml }))),
    );
  });

  test("reports every edge part as before, but where the part is malformed", async () => {
    const graph = await loadSchemaGraph();
    const malformed =
      "a part the reader refuses has no schema signature to report";
    const unreadable =
      "a part the reader refuses is reported unreadable rather than passed";
    expectEdgesMatch({
      golden: GOLDENS.validateOoxmlPart,
      answers: Object.fromEntries(
        XML_EDGE_INPUTS.map(([name, xml]) => [
          name,
          attempt(() => validateOoxmlPart({ graph, xml })),
        ]),
      ),
      refused: refusedWhen(NEVER),
      reread: {
        "unquoted attribute": malformed,
        "depth 101": malformed,
        "space after self-closing slash": malformed,
        ...Object.fromEntries(
          [
            "mismatched end tag",
            "unclosed element",
            "stray end tag",
            "attribute without value",
            "DOCTYPE",
            "DOCTYPE entity expansion",
            "prototype attribute",
            "constructor element",
            "constructor attribute",
            "unclosed CDATA",
            "unclosed comment",
            "depth 150",
          ].map((name) => [name, unreadable]),
        ),
      },
    });
  });
});

describe("classifyCorpusFile", () => {
  const refused = refusedWhen(
    (value) =>
      typeof value === "object" &&
      value !== null &&
      "kind" in value &&
      value.kind === "not-a-docx",
  );

  test("classifies every fixture package as before", async () => {
    const answers: Record<string, unknown> = {};
    for (const { path, bytes } of packages) {
      answers[path] = await attemptAsync(() => classifyCorpusFile(bytes));
    }
    expectFixturesMatch(GOLDENS.classifyCorpusFile, answers);
  });

  test("refuses what it refused, and what a conformant XML parser refuses", async () => {
    const answers: Record<string, unknown> = {};
    for (const [name, xml] of XML_EDGE_INPUTS) {
      const bytes = await withDocument(base, xml);
      answers[name] = await attemptAsync(() => classifyCorpusFile(bytes));
    }
    expectEdgesMatch({
      golden: GOLDENS.classifyCorpusFile,
      answers,
      refused,
      tightened: Object.fromEntries(
        [
          "two roots",
          "less-than in attribute",
          "HTML entity in text",
          "HTML entity in attribute",
          "unknown entity",
          "null character reference",
          "C0 character reference",
          "surrogate character reference",
          "out-of-range character reference",
          "raw control character",
          "CDATA terminator in text",
          "double hyphen in comment",
          "DOCTYPE entity expansion",
          "declaration inside the root",
          "numeric identifier with HTML entity",
        ].map((name) => [name, NOT_WELL_FORMED]),
      ),
    });
  });
});
