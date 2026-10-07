import { expect, test } from "bun:test";
import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";

import { apiReportsWithCarriageReturns } from "./lib/api-report-line-endings";

const report = "export type Name = string;\n";

test("the report guard detects every CR position, including CRLF and bare CR in stale nested reports", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-api-report-lines-"));
  try {
    const nested = path.join(directory, "obsolete", "nested");
    mkdirSync(nested, { recursive: true });
    writeFileSync(path.join(directory, "index.api.md"), report);
    writeFileSync(path.join(nested, "ignored.md"), "ignored\r\n");
    const candidate = path.join(nested, "stale.api.md");
    const expected = [path.join("obsolete", "nested", "stale.api.md")];
    for (let offset = 0; offset <= report.length; offset++) {
      for (const ending of ["\r", "\r\n", "\n\r"] as const) {
        writeFileSync(candidate, report.slice(0, offset) + ending + report.slice(offset));
        expect(apiReportsWithCarriageReturns(directory)).toEqual(expected);
      }
    }
    writeFileSync(candidate, report);
    expect(apiReportsWithCarriageReturns(directory)).toEqual([]);
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});

test("every committed API report uses LF", () => {
  expect(apiReportsWithCarriageReturns(path.resolve(import.meta.dir, "../api-reports"))).toEqual(
    [],
  );
});
