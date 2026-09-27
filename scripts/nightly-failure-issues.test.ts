import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import {
  issueBody,
  issueTitle,
  parseFailures,
  replayFor,
  seedEntry,
  unparsedFailure,
} from "./nightly-failure-issues";

const fixture = (name: string): string =>
  readFileSync(path.join(import.meta.dir, "fixtures", "nightly-failure-issues", name), "utf8");

const packageOf = (file: string): string | null =>
  file.startsWith("src/") ? "packages/core" : null;

const context = {
  kind: "property" as const,
  runUrl: "https://github.com/stella/folio/actions/runs/36297187033",
  sha: "2bf761d01ae4fe6cf72024227916919581f1389b",
  factor: 10,
  date: "2026-09-27",
};

describe("nightly failure issues", () => {
  test("read fast-check's own report from a `gh run view --log` capture (run 36297187033)", () => {
    const failures = parseFailures(fixture("property-run-36297187033.log"), packageOf);
    expect(failures).toHaveLength(1);
    const [failure] = failures;
    expect(failure).toMatchObject({
      name: "a tracked or suggested replacement redlines only the characters it changes > over generated paragraphs and edits, accepted and rejected",
      file: "packages/core/src/ai-edits/minimalDirectReplacement.property.test.ts",
      seed: -449189980,
      path: "350:1:1:1:1:1:11:10:10:10:11:1:1:1:1:1:1:1:1:5:6:6:4:5:5",
      counterexample:
        '[[{"kind":"field","result":"3.6"},{"kind":"tab"}],{"whole":false,"sliceStart":0,"sliceLength":0.5,"mutations":[{"at":0,"remove":0,"insert":" and"}]},"tracked-changes","word"]',
      pinned: false,
    });
    expect(failure?.error).toContain('Expected: " and3.6"');
    expect(replayFor("property", failure!, 10)).toBe(
      "cd packages/core && PROPERTY_TEST_SEED=-449189980 PROPERTY_TEST_PATH='350:1:1:1:1:1:11:10:10:10:11:1:1:1:1:1:1:1:1:5:6:6:4:5:5' PROPERTY_TEST_NUM_RUNS_FACTOR=10 bun test src/ai-edits/minimalDirectReplacement.property.test.ts -t 'over generated paragraphs and edits, accepted and rejected'",
    );
    expect(seedEntry(failure!, "2026-09-27", null)).toBe(
      [
        '"packages/core/src/ai-edits/minimalDirectReplacement.property.test.ts::over generated paragraphs and edits, accepted and rejected": [',
        '  {"seed":-449189980,"path":"350:1:1:1:1:1:11:10:10:10:11:1:1:1:1:1:1:1:1:5:6:6:4:5:5","note":"nightly 2026-09-27: <what it caught>","date":"2026-09-27"}',
        "]",
      ].join("\n"),
    );
  });

  test("read the PROPERTY_FAILURE line and keep other failures apart", () => {
    const failures = parseFailures(fixture("property-marker.log"), packageOf);
    expect(failures.map(({ name }) => name)).toEqual([
      "a random batch with overlapping, nested and duplicate targets > refuses each conflict and applies the rest as one at a time would",
      "element serialization > escapes attribute values",
    ]);
    const [property, example] = failures;
    expect(property).toMatchObject({
      file: "packages/core/src/ai-edits/batchOverlap.test.ts",
      title: "refuses each conflict and applies the rest as one at a time would",
      seed: -440111207,
      path: "136:9:8:8",
      pinned: true,
      replay:
        "cd packages/core && PROPERTY_TEST_SEED=-440111207 PROPERTY_TEST_PATH='136:9:8:8' bun test src/ai-edits/batchOverlap.test.ts -t 'refuses each conflict and applies the rest as one at a time would'",
    });
    expect(example).toMatchObject({
      file: "packages/core/src/docx/xmlSerialize.property.test.ts",
      seed: null,
      counterexample: null,
    });
    expect(example?.error).toContain("Received:");
  });

  test("an issue carries the replay line, the counterexample and the entry to pin", () => {
    const [failure] = parseFailures(fixture("property-run-36297187033.log"), packageOf);
    const body = issueBody(failure!, context, false);
    expect(issueTitle("property", failure!)).toBe(
      "Nightly property failure: packages/core/src/ai-edits/minimalDirectReplacement.property.test.ts::a tracked or suggested replacement redlines only the characters it changes > over generated paragraphs and edits, accepted and rejected",
    );
    expect(body).toContain("commit `2bf761d01ae4`, factor 10");
    expect(body).toContain("### Replay\n```sh\ncd packages/core && PROPERTY_TEST_SEED=-449189980");
    expect(body).toContain("### Counterexample");
    expect(body).toContain("### Pin it");
    expect(issueBody(failure!, context, true)).toStartWith("Failed again in [the nightly run]");
  });

  test("a pinned seed failing again is not offered for pinning", () => {
    const [pinned] = parseFailures(fixture("property-marker.log"), packageOf);
    const body = issueBody(pinned!, context, false);
    expect(body).toContain("already pins");
    expect(body).not.toContain("### Pin it");
  });

  test("a conformance failure replays its case alone", () => {
    const failures = parseFailures(fixture("conformance.log"), packageOf);
    expect(failures).toHaveLength(1);
    expect(failures[0]?.file).toBe("packages/core/src/__tests__/editorCommandConformance.test.ts");
    expect(replayFor("conformance", failures[0]!, null)).toBe(
      "cd packages/core && FOLIO_CONFORMANCE=full FOLIO_CONFORMANCE_FILTER='^table-merged-cells › deleteRow @ across-cells$' bun test src/__tests__/editorCommandConformance.test.ts -t 'table-merged-cells › deleteRow @ across-cells'",
    );
    const body = issueBody(failures[0]!, { ...context, kind: "conformance", factor: null }, false);
    expect(body).toStartWith("A test failed in");
    expect(body).not.toContain("### Pin it");
  });

  test("a failed run with no failing test in its log still gets an issue", () => {
    expect(parseFailures("error: something crashed\n")).toEqual([]);
    expect(issueTitle("property", unparsedFailure("property"))).toBe(
      "Nightly property failure: <unknown file>::property sweep failed without a failing test in the log",
    );
  });

  test("same test name in different files produces separate issues", () => {
    const failures = parseFailures(
      [
        "scripts/first.test.ts:",
        "(fail) same property",
        "scripts/second.test.ts:",
        "(fail) same property",
      ].join("\n"),
    );
    expect(failures.map(({ file }) => file)).toEqual([
      "scripts/first.test.ts",
      "scripts/second.test.ts",
    ]);
    expect(new Set(failures.map((failure) => issueTitle("property", failure))).size).toBe(2);
  });
});
