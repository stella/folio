import { describe, expect, test } from "bun:test";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";

import { cliVersion, PLUGIN_VERSION_FILES, pluginVersions } from "./sync-plugin-cli-version";

describe("plugin CLI versions", () => {
  test("every plugin manifest and CLI spec names the CLI package's version", async () => {
    const version = await cliVersion();

    const found = await pluginVersions();

    expect(found.map(({ file }) => file)).toContain("plugins/vscode/package.json");
    expect(found.length).toBe(5);
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
    // The block scalar's entries are the lines indented past its key.
    const lines = ci.split("\n");
    const key = lines.findIndex((line) => line.trimStart() === "generated-paths: |");
    const keyIndent = (lines[key] ?? "").search(/\S/u);
    const allowed = new Set<string>();
    for (const line of lines.slice(key + 1)) {
      if (line.trim() !== "" && line.search(/\S/u) <= keyIndent) break;
      allowed.add(line.trim());
    }

    expect(key).toBeGreaterThan(-1);
    for (const { file } of PLUGIN_VERSION_FILES) {
      const relative = path.relative(repoRoot, file);
      expect([relative, allowed.has(relative)]).toEqual([relative, true]);
    }
  });
});
