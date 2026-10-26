import { expect, setDefaultTimeout, test } from "bun:test";
import { readFileSync } from "node:fs";
import path from "node:path";

import { scanTestTimeout } from "../test/scan-timeout";
import { CONSUMER_DEPENDENCIES } from "./consumer-scenario-dependencies";
import { importedModuleNames, parseSourceModule } from "./lib/oxc-ast";

setDefaultTimeout(scanTestTimeout(30_000));

test("every third-party import in staged scenarios has an explicit consumer install", async () => {
  const root = path.resolve(import.meta.dir, "../test/consumer-scenarios");
  const installed = new Set(
    CONSUMER_DEPENDENCIES.map((dependency) =>
      dependency.slice(0, dependency.lastIndexOf("@")),
    ),
  );
  const undeclared: string[] = [];
  let corpusChecked = false;
  for (const directory of ["scenarios", "support"]) {
    for await (const file of new Bun.Glob("**/*.ts").scan(
      path.join(root, directory),
    )) {
      const source = readFileSync(path.join(root, directory, file), "utf8");
      if (file === "public-corpus.ts") corpusChecked = true;
      for (const specifier of importedModuleNames(
        parseSourceModule(file, source),
      )) {
        if (
          specifier.startsWith(".") ||
          specifier.startsWith("node:") ||
          specifier.startsWith("@stll/")
        )
          continue;
        const parts = specifier.split("/");
        const name = parts
          .slice(0, specifier.startsWith("@") ? 2 : 1)
          .join("/");
        if (!installed.has(name))
          undeclared.push(`${directory}/${file}: ${specifier}`);
      }
    }
  }
  expect(corpusChecked).toBe(true);
  expect(undeclared).toEqual([]);
});
