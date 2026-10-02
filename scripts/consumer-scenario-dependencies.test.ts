import { expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";
import ts from "typescript";

import { CONSUMER_DEPENDENCIES } from "./consumer-scenario-dependencies";

test("every third-party import in staged scenarios has an explicit consumer install", async () => {
  const root = path.resolve(import.meta.dir, "../test/consumer-scenarios");
  const installed = new Set(
    CONSUMER_DEPENDENCIES.map((dependency) => dependency.slice(0, dependency.lastIndexOf("@"))),
  );
  const undeclared: string[] = [];
  let corpusChecked = false;
  for (const directory of ["scenarios", "support"]) {
    for await (const file of new Bun.Glob("**/*.ts").scan(path.join(root, directory))) {
      const source = readFileSync(path.join(root, directory, file), "utf8");
      if (file === "public-corpus.ts") corpusChecked = true;
      for (const { fileName: specifier } of ts.preProcessFile(source).importedFiles) {
        if (
          specifier.startsWith(".") ||
          specifier.startsWith("node:") ||
          specifier.startsWith("@stll/")
        )
          continue;
        const parts = specifier.split("/");
        const name = parts.slice(0, specifier.startsWith("@") ? 2 : 1).join("/");
        if (!installed.has(name)) undeclared.push(`${directory}/${file}: ${specifier}`);
      }
    }
  }
  expect(corpusChecked).toBe(true);
  expect(undeclared).toEqual([]);
});
