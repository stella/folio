/**
 * Fail when a parse-boundary normaliser neither warns nor says why it need not.
 *
 * Usage: bun scripts/check-normalization-warnings.ts
 */

import { readdir } from "node:fs/promises";
import path from "node:path";

import {
  NORMALIZATION_MODULE_SUFFIX,
  type NormalizationModule,
  findNormalizationWarningViolations,
} from "./lib/normalization-warnings";

const REPOSITORY_ROOT = path.resolve(import.meta.dir, "..");
const SCANNED_DIRECTORIES = ["packages/core/src/docx", "packages/core/src/style-sets"];

const loadModules = async (): Promise<NormalizationModule[]> => {
  const modules: NormalizationModule[] = [];
  for (const directory of SCANNED_DIRECTORIES) {
    const entries = await readdir(path.join(REPOSITORY_ROOT, directory));
    for (const entry of entries.sort()) {
      if (!entry.endsWith(NORMALIZATION_MODULE_SUFFIX)) {
        continue;
      }
      const relative = `${directory}/${entry}`;
      modules.push({
        path: relative,
        source: await Bun.file(path.join(REPOSITORY_ROOT, relative)).text(),
      });
    }
  }
  return modules;
};

const main = async (): Promise<void> => {
  const modules = await loadModules();
  const violations = findNormalizationWarningViolations(modules);
  if (violations.length === 0) {
    process.stdout.write(
      `check-normalization-warnings: ${String(modules.length)} normalisers, all accounted for\n`,
    );
    return;
  }
  process.stderr.write(
    `A parse-boundary normaliser must report what it normalised:\n${violations
      .map(({ path: modulePath, detail }) => `- ${modulePath} ${detail}`)
      .join("\n")}\n`,
  );
  process.exitCode = 1;
};

if (import.meta.main) {
  await main();
}
