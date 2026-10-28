import { describe, setDefaultTimeout, test } from "bun:test";
import { resolve } from "node:path";

import { scanTestTimeout } from "../../../test/scan-timeout";
import {
  attemptAsync,
  expectEdgesMatch,
  expectFixturesMatch,
  readGoldens,
  refusedWhen,
  SAMPLE_FIXTURE,
  withDocument,
} from "../../../test/xmlGoldens";
import { XML_EDGE_INPUTS } from "../../../tests/support/xmlEdgeInputs";
import { readFixturePackages } from "../../../tests/support/xmlFixtureParts";
import { checkWordprocessingPackage } from "./main-document-part";
import { checkPackageIntegrity } from "./package-parts";

setDefaultTimeout(scanTestTimeout(300_000));

const GOLDENS = readGoldens<
  "checkWordprocessingPackage" | "checkPackageIntegrity"
>(resolve(import.meta.dir, "__tests__/xml-goldens.json"));
const packages = await readFixturePackages();
const base = packages.find(({ path }) => path === SAMPLE_FIXTURE);
if (base === undefined) throw new Error(`missing ${SAMPLE_FIXTURE}`);

type CheckResult = {
  isOk(): boolean;
  isErr(): boolean;
  error?: { message: string };
};
const resultAnswer = (result: CheckResult) =>
  result.isOk() ? "ok" : { error: result.error?.message };
const refused = refusedWhen((value) => value !== "ok");
const NOT_WELL_FORMED = "not well-formed XML 1.0";

describe("checkWordprocessingPackage", () => {
  test("answers every fixture package as before", async () => {
    const answers: Record<string, unknown> = {};
    for (const { path, bytes } of packages) {
      answers[path] = await attemptAsync(async () =>
        resultAnswer(await checkWordprocessingPackage(path, bytes)),
      );
    }
    expectFixturesMatch(GOLDENS.checkWordprocessingPackage, answers);
  });
  test("answers every edge document as before", async () => {
    const answers: Record<string, unknown> = {};
    for (const [name, xml] of XML_EDGE_INPUTS) {
      const bytes = await withDocument(base, xml);
      answers[name] = await attemptAsync(async () =>
        resultAnswer(await checkWordprocessingPackage("edge.docx", bytes)),
      );
    }
    expectEdgesMatch({
      golden: GOLDENS.checkWordprocessingPackage,
      answers,
      refused,
    });
  });
});

describe("checkPackageIntegrity", () => {
  test("answers every fixture package as before", async () => {
    const answers: Record<string, unknown> = {};
    for (const { path, bytes, parts } of packages) {
      const changed = [...parts.keys()].map((part) => ({
        part,
        change: "modified" as const,
      }));
      answers[path] = await attemptAsync(async () =>
        resultAnswer(await checkPackageIntegrity(bytes, changed)),
      );
    }
    expectFixturesMatch(GOLDENS.checkPackageIntegrity, answers);
  });
  test("refuses what it refused", async () => {
    const answers: Record<string, unknown> = {};
    for (const [name, xml] of XML_EDGE_INPUTS) {
      const bytes = await withDocument(base, xml);
      answers[name] = await attemptAsync(async () =>
        resultAnswer(
          await checkPackageIntegrity(bytes, [
            { part: "word/document.xml", change: "modified" },
          ]),
        ),
      );
    }
    expectEdgesMatch({
      golden: GOLDENS.checkPackageIntegrity,
      answers,
      refused,
      tightened: {
        "two roots": NOT_WELL_FORMED,
        "less-than in attribute": NOT_WELL_FORMED,
        "raw control character": NOT_WELL_FORMED,
        "CDATA terminator in text": NOT_WELL_FORMED,
        "double hyphen in comment": NOT_WELL_FORMED,
        "declaration inside the root": NOT_WELL_FORMED,
      },
    });
  });
});
