// Merge the per-document files written by the second-engine layout spec, and
// the Playwright JSON results, into one report and a markdown summary.
//
// Report-only: differences never change the exit code. It exits 1 only for
// harness problems: no results file, runner-level errors, no tests executed,
// a layout-record test that failed (it fails only when a document cannot be
// loaded or read), or a document that wrote no output.
//
// Usage: bun scripts/engine-parity-summary.ts <playwright-results.json> <out-dir>
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import path from "node:path";

import { summarizeDocument, type Difference } from "../tests/visual/engineLayoutDiff";

const LAYOUT_SPEC = "engine-layout-parity.spec.ts";

type DocumentReport = {
  fixture: string;
  engine: string;
  pageCount: { expected: number | null; actual: number };
  differences: Difference[];
  fonts: { expected: unknown; actual: { stack: string; weight: string; loaded: boolean }[] };
};

type Kerning = {
  engine: string;
  result: { family: string; loaded: boolean; delta: number; changes: boolean }[];
};

type TestRun = { file: string; title: string; status: string };

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

/** Flatten Playwright's JSON reporter tree into one row per test. */
const collectTests = (node: unknown, out: TestRun[] = []): TestRun[] => {
  if (!isRecord(node)) return out;
  for (const suite of Array.isArray(node["suites"]) ? node["suites"] : []) collectTests(suite, out);
  for (const spec of Array.isArray(node["specs"]) ? node["specs"] : []) {
    if (!isRecord(spec)) continue;
    for (const test of Array.isArray(spec["tests"]) ? spec["tests"] : []) {
      if (!isRecord(test)) continue;
      const results = Array.isArray(test["results"]) ? test["results"] : [];
      const last = results.at(-1);
      out.push({
        file: String(spec["file"] ?? ""),
        title: String(spec["title"] ?? ""),
        status: isRecord(last) ? String(last["status"]) : String(test["status"] ?? "unknown"),
      });
    }
  }
  return out;
};

const [resultsPath, outDir] = process.argv.slice(2);
if (resultsPath === undefined || outDir === undefined) {
  console.error("usage: bun scripts/engine-parity-summary.ts <playwright-results.json> <out-dir>");
  process.exit(2);
}

const harnessErrors: string[] = [];
let tests: TestRun[] = [];
if (!existsSync(resultsPath)) {
  harnessErrors.push(`no Playwright results at ${resultsPath}`);
} else {
  const results: unknown = JSON.parse(readFileSync(resultsPath, "utf8"));
  if (isRecord(results) && Array.isArray(results["errors"]) && results["errors"].length > 0) {
    harnessErrors.push(`runner reported ${String(results["errors"].length)} error(s)`);
  }
  tests = collectTests(results);
  if (tests.length === 0) harnessErrors.push("no tests executed");
}

for (const test of tests) {
  if (test.file.endsWith(LAYOUT_SPEC) && test.status !== "passed") {
    harnessErrors.push(`${test.title}: ${test.status}`);
  }
}

const readJson = (file: string): unknown =>
  JSON.parse(readFileSync(path.join(outDir, file), "utf8"));
const files = existsSync(outDir) ? readdirSync(outDir).toSorted() : [];
const documents = files
  .filter((file) => file.endsWith(".docx.json"))
  // SAFETY: written by the spec this script summarises.
  .map((file) => readJson(file) as DocumentReport);
const kerning = files.includes("_kerning.json") ? (readJson("_kerning.json") as Kerning) : null;

const documentTests = tests.filter(
  (test) =>
    test.file.endsWith(LAYOUT_SPEC) &&
    test.title.startsWith("records ") &&
    test.title.endsWith(".docx"),
);
if (documents.length < documentTests.length) {
  harnessErrors.push(
    `${String(documentTests.length - documents.length)} document(s) wrote no output`,
  );
}
if (kerning === null) harnessErrors.push("kerning probe wrote no output");

// Outcome of the measure-backend spec step, which runs as-is under the second
// engine with its assertions intact: "failure" is a finding, not a harness error.
const backendOutcome = process.env["BACKEND_OUTCOME"] ?? "unknown";
const differences = documents.flatMap((doc) => doc.differences);

writeFileSync(
  path.join(outDir, "report.json"),
  `${JSON.stringify({ differences, documents, kerning, backendOutcome, harnessErrors }, null, 2)}\n`,
);

const lines = [
  "# Layout under a second engine",
  "",
  `Engine: ${documents.at(0)?.engine ?? kerning?.engine ?? "unknown"}. Report-only: differences do not fail the run.`,
  "",
  "| Document | Differences | Kinds |",
  "| --- | --- | --- |",
  ...documents.map((doc) => summarizeDocument(doc.fixture, doc.differences)),
  "",
  "## Resolved fonts",
  "",
  ...documents.map(
    (doc) =>
      `- ${doc.fixture}: ${
        doc.fonts.actual
          .map((font) => `${font.weight} ${font.stack}${font.loaded ? "" : " (not loaded)"}`)
          .join("; ") || "none"
      }`,
  ),
  "",
  "## fontKerning none vs normal (canvas measureText)",
  "",
  ...(kerning?.result.map(
    (row) =>
      `- ${row.family}: ${row.changes ? "width changes" : "no change"} (delta ${row.delta.toFixed(3)}px${row.loaded ? "" : ", face not loaded"})`,
  ) ?? ["- not recorded"]),
  "",
  "## Measure backend parity under this engine",
  "",
  backendOutcome === "failure"
    ? "- assertion failed under second engine"
    : `- step outcome: ${backendOutcome}`,
  "",
  ...(harnessErrors.length > 0
    ? ["## Harness errors", "", ...harnessErrors.map((e) => `- ${e}`), ""]
    : []),
];
writeFileSync(path.join(outDir, "summary.md"), `${lines.join("\n")}\n`);

if (harnessErrors.length > 0) {
  console.error(harnessErrors.join("\n"));
  process.exit(1);
}
console.log(
  `${String(differences.length)} difference(s) across ${String(documents.length)} document(s)`,
);
