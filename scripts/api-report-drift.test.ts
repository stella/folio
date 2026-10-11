import { expect, setDefaultTimeout, test } from "bun:test";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../test/property-testing";
import { renderApiReportDrift } from "./lib/api-report-drift";

setDefaultTimeout(propertyTestTimeout(5_000));

test("new API entries report their commit path and full generated surface without a baseline", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-new-api-report-"));
  const baselinePath = path.join(directory, "committed.api.md");
  const candidatePath = path.join(directory, "generated.api.md");
  try {
    assertProperty(
      fc.property(
        fc.array(fc.string({ minLength: 1, maxLength: 30 }), { minLength: 2, maxLength: 40 }),
        fc.integer({ min: 1, max: 10 }),
        (lines, maxLines) => {
          const candidate = `## API Report File for "@stll/docx-core"\n\n${lines.join("\n")}\n`;
          writeFileSync(candidatePath, candidate);
          expect(
            renderApiReportDrift({
              baselinePath,
              candidatePath,
              reportPath: "api-reports/docx-core/zip.api.md",
              maxLines,
            }),
          ).toBe(`new entry: commit api-reports/docx-core/zip.api.md\n${candidate}`);
        },
      ),
      {
        numRuns: 30,
        id: "new API entries report their commit path and full generated surface without a baseline",
      },
    );
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});

test("existing API entries retain the bounded drift diagnostic", () => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-existing-api-report-"));
  const baselinePath = path.join(directory, "committed.api.md");
  const candidatePath = path.join(directory, "generated.api.md");
  try {
    writeFileSync(baselinePath, "old\nshared");
    writeFileSync(candidatePath, "new\nshared");
    expect(
      renderApiReportDrift({
        baselinePath,
        candidatePath,
        reportPath: "api-reports/docx-core/zip.api.md",
        maxLines: 1,
      }),
    ).toBe("-old\n… 1 more diff line(s) truncated");
  } finally {
    rmSync(directory, { recursive: true, force: true });
  }
});
