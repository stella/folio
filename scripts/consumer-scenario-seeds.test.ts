import assert from "node:assert/strict";
import { existsSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { afterEach, test } from "node:test";

import frozenFlows from "./fixtures/consumer-scenario-seeds.snapshot.json" with { type: "json" };
import {
  loadScenarioSeedRegistry,
  scenarioSeedFileName,
  type PinnedScenarioFlow,
} from "../test/consumer-scenarios/support/scenario-seeds.ts";

const seedDirectory = path.join(import.meta.dir, "../test/consumer-scenarios/seeds");
const temporaryDirectories: string[] = [];
const flow = {
  seed: 41,
  steps: 3,
  kind: "random",
  generation: "targeted",
  title: "keeps a stable example",
} as const satisfies PinnedScenarioFlow;

const makeRegistry = (...entries: Array<{ fileName: string; value: unknown }>): string => {
  const directory = mkdtempSync(path.join(tmpdir(), "folio-scenario-seeds-"));
  temporaryDirectories.push(directory);
  for (const { fileName, value } of entries) {
    writeFileSync(path.join(directory, fileName), `${JSON.stringify(value)}\n`);
  }
  return directory;
};

afterEach(() => {
  for (const directory of temporaryDirectories.splice(0)) {
    rmSync(directory, { recursive: true, force: true });
  }
});

test("the seed directory contains exactly the frozen migrated registry", () => {
  const flows = loadScenarioSeedRegistry(seedDirectory);

  assert.equal(flows.length, 64);
  assert.deepEqual(
    [...flows].sort((left, right) =>
      scenarioSeedFileName(left).localeCompare(scenarioSeedFileName(right)),
    ),
    [...frozenFlows].sort((left, right) =>
      scenarioSeedFileName(left).localeCompare(scenarioSeedFileName(right)),
    ),
  );
});

test("the registry loads files in stable file-name order", () => {
  const second = { ...flow, seed: 9, title: "second" };
  const first = { ...flow, seed: 100, title: "first" };
  const directory = makeRegistry(
    { fileName: scenarioSeedFileName(second), value: second },
    { fileName: scenarioSeedFileName(first), value: first },
  );

  assert.deepEqual(
    loadScenarioSeedRegistry(directory).map(({ seed }) => seed),
    [100, 9],
  );
});

test("the registry rejects malformed entries", () => {
  const directory = makeRegistry({ fileName: "41-random-invalid.json", value: { seed: "bad" } });

  assert.throws(
    () => loadScenarioSeedRegistry(directory),
    /Invalid pinned consumer scenario flow/u,
  );
});

test("the registry rejects duplicate seed and kind pairs", () => {
  const duplicate = { ...flow, title: "a different title" };
  const directory = makeRegistry(
    { fileName: scenarioSeedFileName(flow), value: flow },
    { fileName: scenarioSeedFileName(duplicate), value: duplicate },
  );

  assert.throws(() => loadScenarioSeedRegistry(directory), /duplicate seed and kind/u);
});

test("the registry rejects a file name that does not match its content", () => {
  const directory = makeRegistry({ fileName: "41-random-wrong-title.json", value: flow });

  assert.throws(() => loadScenarioSeedRegistry(directory), /file name/u);
});

const filesReferencing = (fileNames: string[], removedName: string): string[] =>
  fileNames.filter(
    (fileName) =>
      existsSync(fileName) &&
      statSync(fileName).isFile() &&
      readFileSync(fileName, "utf8").includes(removedName),
  );

test("the legacy-reference guard detects a leftover reference", () => {
  const removedName = ["scenario", "seeds.json"].join("-");
  const directory = makeRegistry({ fileName: "reference.txt", value: removedName });

  assert.deepEqual(filesReferencing([path.join(directory, "reference.txt")], removedName), [
    path.join(directory, "reference.txt"),
  ]);
});

test("the repository has no references to the removed aggregate registry", () => {
  const repository = path.join(import.meta.dir, "..");
  const removedName = ["scenario", "seeds.json"].join("-");
  const tracked = spawnSync("git", ["ls-files", "-co", "--exclude-standard"], {
    cwd: repository,
    encoding: "utf8",
  });
  assert.equal(tracked.status, 0, tracked.stderr);
  const fileNames = tracked.stdout
    .trim()
    .split("\n")
    .filter(Boolean)
    .map((fileName) => path.join(repository, fileName));

  assert.deepEqual(filesReferencing(fileNames, removedName), []);
});
