import { join } from "node:path";

import {
  SUGGESTION_INPUT_DRIVERS,
  SUGGESTION_INPUT_KINDS,
} from "../packages/core/src/__tests__/suggestionInputKinds";

// Playwright's JSON reporter removes the leading @ from registered tags.
const tagPrefix = "browser-input:";
const root = join(import.meta.dir, "..");

type PlaywrightSpec = { file: string; title: string; tags: string[] };
type PlaywrightSuite = { specs: PlaywrightSpec[]; suites?: PlaywrightSuite[] };
type PlaywrightReport = { suites: PlaywrightSuite[]; errors: unknown[] };

const result = Bun.spawnSync(
  ["bunx", "playwright", "test", "--project=interactions", "--list", "--reporter=json"],
  { cwd: root, stdout: "pipe", stderr: "pipe" },
);
if (result.exitCode !== 0) {
  throw new Error(`Playwright test discovery failed:\n${result.stderr.toString()}`);
}

// SAFETY: This JSON is produced by the pinned Playwright reporter invoked above.
const report = JSON.parse(result.stdout.toString()) as PlaywrightReport;
if (report.errors.length > 0) {
  throw new Error(`Playwright test discovery reported errors: ${JSON.stringify(report.errors)}`);
}

const registered = new Map<string, string[]>();
const collect = (suite: PlaywrightSuite) => {
  for (const spec of suite.specs) {
    for (const tag of spec.tags) {
      if (!tag.startsWith(tagPrefix)) continue;
      const kind = tag.slice(tagPrefix.length);
      const locations = registered.get(kind) ?? [];
      locations.push(`${spec.file}: ${spec.title}`);
      registered.set(kind, locations);
    }
  }
  for (const child of suite.suites ?? []) collect(child);
};
for (const suite of report.suites) collect(suite);

const declared = SUGGESTION_INPUT_KINDS.filter(
  (kind) => SUGGESTION_INPUT_DRIVERS[kind].type === "browser",
);
const missing = declared.filter((kind) => !registered.has(kind));
const unexpected = [...registered.keys()].filter(
  (kind) => !declared.some((declaredKind) => declaredKind === kind),
);
const duplicated = [...registered].filter(([, locations]) => locations.length !== 1);

if (missing.length || unexpected.length || duplicated.length) {
  throw new Error(
    [
      `Browser input coverage differs from discovered Playwright tests.`,
      `Missing: ${missing.join(", ") || "none"}`,
      `Unexpected: ${unexpected.join(", ") || "none"}`,
      `Registered more than once: ${duplicated.map(([kind, locations]) => `${kind} (${locations.join("; ")})`).join(", ") || "none"}`,
    ].join("\n"),
  );
}

console.log(`Browser input coverage: ${declared.length} kinds registered in Playwright.`);
