import { afterEach, expect, test } from "bun:test";
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { reportsWithCarriageReturns } from "./lib/api-report-line-endings";

const directories: string[] = [];
afterEach(() => {
  for (const directory of directories.splice(0)) rmSync(directory, { recursive: true });
});
const reportTree = () => {
  const directory = mkdtempSync(join(tmpdir(), "folio-api-report-lf-"));
  directories.push(directory);
  return directory;
};

test("all nested reports are checked, while unrelated files are ignored", () => {
  const directory = reportTree();
  const files = {
    "core/index.api.md": "alpha\r\nbeta\r\n",
    "core/ai-edits.api.md": "alpha\nbeta\n",
    "vue/nested/index.api.md": "alpha\rbeta",
    "README.md": "ignored\r\n",
  };
  for (const [file, contents] of Object.entries(files)) {
    const target = join(directory, file);
    mkdirSync(dirname(target), { recursive: true });
    writeFileSync(target, contents);
  }
  expect(reportsWithCarriageReturns(directory)).toEqual([
    "core/index.api.md",
    "vue/nested/index.api.md",
  ]);
});

test("a carriage return at every position in a report is rejected", () => {
  const directory = reportTree();
  const report = join(directory, "index.api.md");
  for (const content of ["", "alpha\nbeta\n", "契約\nمرحبا\n😀"]) {
    writeFileSync(report, content);
    expect(reportsWithCarriageReturns(directory)).toEqual([]);
    for (let position = 0; position <= content.length; position++) {
      writeFileSync(report, `${content.slice(0, position)}\r${content.slice(position)}`);
      expect(reportsWithCarriageReturns(directory)).toEqual(["index.api.md"]);
    }
  }
});
