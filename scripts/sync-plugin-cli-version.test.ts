import { describe, expect, test } from "bun:test";

import { cliVersion, pluginVersions } from "./sync-plugin-cli-version";

describe("Claude Code plugin CLI version", () => {
  test("the plugin manifest and its MCP server run the CLI package's version", async () => {
    const version = await cliVersion();

    const found = await pluginVersions();

    expect(found.manifest).toBe(version);
    expect(found.mcp).toEqual([version]);
  });
});
