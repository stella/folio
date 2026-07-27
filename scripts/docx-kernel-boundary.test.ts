import { describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import path from "node:path";
import ts from "typescript";

const projectionPath = path.join(
  import.meta.dir,
  "..",
  "packages",
  "docx-core",
  "src",
  "projection.ts",
);

describe("DOCX projection architecture boundary", () => {
  test("keeps the TypeScript surface limited to the generated WASM binding", async () => {
    const source = await readFile(projectionPath, "utf8");
    const importedModules = ts.preProcessFile(source).importedFiles.map(({ fileName }) => fileName);

    expect(importedModules).toEqual(["better-result", "./generated/docx_kernel.js"]);
    expect(source).not.toContain("jszip");
    expect(source).not.toContain("fast-xml-parser");
    expect(source).not.toContain("DOMParser");
  });
});
