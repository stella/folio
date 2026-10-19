import { readdir } from "node:fs/promises";
import { join } from "node:path";

import { exportedConstAssertionsWithoutTypes } from "./lib/exported-const-assertions";

const root = join(import.meta.dir, "..");
const packagesDirectory = join(root, "packages");
const findings = new Map<string, string>();

for (const packageEntry of await readdir(packagesDirectory, { withFileTypes: true })) {
  if (!packageEntry.isDirectory()) continue;
  const packageDirectory = join(packagesDirectory, packageEntry.name);
  const manifestFile = Bun.file(join(packageDirectory, "package.json"));
  if (!(await manifestFile.exists())) continue;
  const manifest = (await manifestFile.json()) as { private?: boolean };
  if (manifest.private) continue;

  const sourceDirectory = join(packageDirectory, "src");
  const glob = new Bun.Glob("**/*.{ts,tsx}");
  for await (const relativeFile of glob.scan({ cwd: sourceDirectory, onlyFiles: true })) {
    if (
      /\.(?:test|typecheck)\.[^.]+$/u.test(relativeFile) ||
      relativeFile.startsWith("__tests__/") ||
      relativeFile.includes("/__tests__/")
    ) {
      continue;
    }
    const absoluteFile = join(sourceDirectory, relativeFile);
    const source = await Bun.file(absoluteFile).text();
    for (const finding of exportedConstAssertionsWithoutTypes(source, absoluteFile)) {
      const relativeSource = join("packages", packageEntry.name, "src", relativeFile);
      findings.set(
        `${relativeSource}:${finding.name}`,
        `${relativeSource}:${finding.line}:${finding.column} ${finding.name} needs an explicit type annotation`,
      );
    }
  }
}

const baseline = new Set(
  (await Bun.file(join(import.meta.dir, "exported-const-assertions.baseline.txt")).text())
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean),
);
const failures = [...findings].filter(([key]) => !baseline.has(key)).map(([, message]) => message);
const stale = [...baseline].filter((key) => !findings.has(key));

if (failures.length > 0 || stale.length > 0) {
  console.error(
    "Exported const-asserted arrays and objects in published packages need explicit types.\n" +
      failures.join("\n") +
      (stale.length > 0 ? `\nRemove resolved baseline entries:\n${stale.join("\n")}` : ""),
  );
  process.exit(1);
}

console.log("No unbaselined const-asserted aggregate exports lack explicit types.");
