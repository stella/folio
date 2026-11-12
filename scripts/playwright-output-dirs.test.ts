/**
 * A browser job's failure evidence (traces, screenshots, error context) is
 * only useful when CI uploads it. Every Playwright config writes its results
 * inside a directory a workflow uploads; `tests/support/playwrightOutput.ts`
 * owns that directory for configs that name their own.
 */

import { expect, setDefaultTimeout, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

import {
  expectExited,
  runChild,
  subprocessTestTimeout,
} from "../test/subprocess";
import { PLAYWRIGHT_RESULTS_ROOT } from "../tests/support/playwrightOutput";

setDefaultTimeout(subprocessTestTimeout());

const REPO_ROOT = path.resolve(import.meta.dir, "..");
const WORKFLOWS = path.join(REPO_ROOT, ".github", "workflows");
const PLAYWRIGHT_CONFIG = /(?:^|\/)playwright[^/]*\.config\.[cm]?[jt]s$/u;
/** Configs a bun test runs in a scratch directory; their output is never uploaded. */
const UNUPLOADED_FIXTURES = new Set([
  "scripts/__fixtures__/protocol-capture/playwright.config.ts",
]);

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

type ConfigOutput = { config: string; outputDir: string };

/** Every config whose results land outside the uploaded directories. */
const uncoveredOutputDirs = (
  outputs: readonly ConfigOutput[],
  uploaded: readonly string[],
): string[] =>
  outputs.flatMap(({ config, outputDir }) =>
    uploaded.some((root) => {
      const inside = path.relative(root, outputDir);
      return !inside.startsWith("..") && !path.isAbsolute(inside);
    })
      ? []
      : [
          `${config}: writes ${path.relative(REPO_ROOT, outputDir)}, which no workflow uploads`,
        ],
  );

/** The directories `actions/upload-artifact` steps upload, as absolute paths. */
const uploadedDirectories = (): string[] =>
  readdirSync(WORKFLOWS)
    .filter((file) => /\.ya?ml$/u.test(file))
    .flatMap((file) => {
      const workflow: unknown = Bun.YAML.parse(
        readFileSync(path.join(WORKFLOWS, file), "utf8"),
      );
      if (!isRecord(workflow) || !isRecord(workflow["jobs"])) return [];
      return Object.values(workflow["jobs"]).flatMap((job) => {
        if (!isRecord(job) || !Array.isArray(job["steps"])) return [];
        return job["steps"].flatMap((step: unknown) => {
          if (
            !isRecord(step) ||
            typeof step["uses"] !== "string" ||
            !step["uses"].startsWith("actions/upload-artifact@") ||
            !isRecord(step["with"]) ||
            typeof step["with"]["path"] !== "string"
          )
            return [];
          return step["with"]["path"]
            .split("\n")
            .map((entry) => entry.trim())
            .filter((entry) => entry !== "" && !entry.includes("${{"))
            .map((entry) => path.resolve(REPO_ROOT, entry));
        });
      });
    });

const trackedConfigs = (): string[] => {
  const listed = expectExited(
    runChild(["git", "ls-files", "-z"], { cwd: REPO_ROOT }),
  );
  if (listed.exitCode !== 0)
    throw new TypeError(`git ls-files failed: ${listed.stderr}`);
  return listed.stdout
    .toString()
    .split("\0")
    .filter((file) => PLAYWRIGHT_CONFIG.test(file));
};

/** Where a config writes results: its `outputDir`, or Playwright's default beside it. */
const configOutput = async (config: string): Promise<ConfigOutput> => {
  const absolute = path.join(REPO_ROOT, config);
  const loaded: unknown = await import(absolute);
  const exported = isRecord(loaded) ? loaded["default"] : undefined;
  const outputDir =
    isRecord(exported) && typeof exported["outputDir"] === "string"
      ? exported["outputDir"]
      : "test-results";
  return {
    config,
    outputDir: path.resolve(path.dirname(absolute), outputDir),
  };
};

test("every Playwright config writes its results where CI uploads them", async () => {
  const uploaded = uploadedDirectories();
  expect(uploaded).toContain(PLAYWRIGHT_RESULTS_ROOT);
  const configs = trackedConfigs().filter(
    (config) => !UNUPLOADED_FIXTURES.has(config),
  );
  expect(configs).toContain("packages/editor-web/playwright.config.ts");
  const outputs = await Promise.all(configs.map(configOutput));
  expect(uncoveredOutputDirs(outputs, uploaded)).toEqual([]);
});

test("a config writing outside the uploaded directories is reported", () => {
  const uploaded = [PLAYWRIGHT_RESULTS_ROOT];
  expect(
    uncoveredOutputDirs(
      [
        {
          config: "packages/example/playwright.config.ts",
          outputDir: path.join(REPO_ROOT, "packages/example/test-results"),
        },
        {
          config: "playwright.config.ts",
          outputDir: path.join(PLAYWRIGHT_RESULTS_ROOT, "example"),
        },
      ],
      uploaded,
    ),
  ).toEqual([
    "packages/example/playwright.config.ts: writes packages/example/test-results, which no workflow uploads",
  ]);
});

test("each uncovered fixture still exists", () => {
  const configs = new Set(trackedConfigs());
  for (const fixture of UNUPLOADED_FIXTURES)
    expect(configs.has(fixture)).toBe(true);
});
