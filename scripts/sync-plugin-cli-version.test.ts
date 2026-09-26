import { describe, expect, test } from "bun:test";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";

import { cliVersion, PLUGIN_VERSION_FILES, pluginVersions } from "./sync-plugin-cli-version";

describe("plugin CLI versions", () => {
  test("every plugin manifest and CLI spec names the CLI package's version", async () => {
    const version = await cliVersion();

    const found = await pluginVersions();

    expect(found.length).toBe(4);
    for (const { file, versions } of found) {
      expect([file, versions.every((entry) => entry === version)]).toEqual([file, true]);
    }
  });

  // The version PR is committed through the GitHub API, which refuses
  // executable files; the synced files must stay plain.
  test("no synced file is executable", async () => {
    for (const { file } of PLUGIN_VERSION_FILES) {
      const { mode } = await stat(file);
      expect([file, mode & 0o111]).toEqual([file, 0]);
    }
  });

  // The release policy only lets the version PR touch listed generated
  // paths; every synced plugin file must be one of them.
  test("the version PR policy allows every synced file", async () => {
    const repoRoot = path.resolve(import.meta.dir, "..");
    const ci = await readFile(path.join(repoRoot, ".github", "workflows", "ci.yml"), "utf8");
    const allowed = new Set(
      (ci.split("generated-paths: |")[1] ?? "")
        .split("\n")
        .slice(1)
        .map((line) => line.trim())
        .filter((line, index, lines) => line !== "" && !lines.slice(0, index).includes("")),
    );
    for (const { file } of PLUGIN_VERSION_FILES) {
      const relative = path.relative(repoRoot, file);
      expect([relative, allowed.has(relative)]).toEqual([relative, true]);
    }
  });
});
