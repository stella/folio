import { describe, expect, test } from "bun:test";
import { stat } from "node:fs/promises";

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
});
