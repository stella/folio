import { describe, expect, test } from "bun:test";

import { cliVersion, pluginVersions } from "./sync-plugin-cli-version";

describe("plugin CLI versions", () => {
  test("every plugin manifest and CLI spec names the CLI package's version", async () => {
    const version = await cliVersion();

    const found = await pluginVersions();

    expect(found.length).toBe(4);
    for (const { file, versions } of found) {
      expect([file, versions.every((entry) => entry === version)]).toEqual([file, true]);
    }
  });
});
