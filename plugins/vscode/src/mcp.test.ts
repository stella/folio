import { describe, expect, test } from "bun:test";

import { buildMcpLaunch, configuredAuthor, FOLIO_MCP_LABEL } from "./mcp";

const runtime = { nodePath: "/editor/node", cliEntry: "/ext/dist/cli/folio.mjs" };

describe("buildMcpLaunch", () => {
  test("runs the bundled CLI's mcp command with the editor's Node.js", () => {
    const launch = buildMcpLaunch({
      runtime,
      roots: ["/work/a"],
      author: "Ada Lovelace",
      version: "1.2.3",
    });

    expect(launch).toEqual({
      label: FOLIO_MCP_LABEL,
      command: "/editor/node",
      args: ["/ext/dist/cli/folio.mjs", "mcp", "--root", "/work/a", "--author", "Ada Lovelace"],
      env: { ELECTRON_RUN_AS_NODE: "1" },
      cwd: "/work/a",
      version: "1.2.3",
    });
  });

  test("passes every workspace folder as a root, once, in order", () => {
    const launch = buildMcpLaunch({
      runtime,
      roots: ["/work/b", "/work/a", "/work/b"],
      author: "Ada",
      version: "1.0.0",
    });

    expect(launch?.args).toEqual([
      "/ext/dist/cli/folio.mjs",
      "mcp",
      "--root",
      "/work/b",
      "--root",
      "/work/a",
      "--author",
      "Ada",
    ]);
    expect(launch?.cwd).toBe("/work/b");
  });

  test("keeps a root with spaces as one argument", () => {
    const launch = buildMcpLaunch({
      runtime,
      roots: ["/Users/me/My Documents"],
      author: "Ada",
      version: "1.0.0",
    });

    expect(launch?.args).toContain("/Users/me/My Documents");
  });

  test("always passes the resolved author", () => {
    const launch = buildMcpLaunch({
      runtime,
      roots: ["/work"],
      author: "Git Name",
      version: "1.0.0",
    });

    expect(launch?.args).toEqual([
      "/ext/dist/cli/folio.mjs",
      "mcp",
      "--root",
      "/work",
      "--author",
      "Git Name",
    ]);
  });

  test("offers no server without a folder to root it in", () => {
    expect(buildMcpLaunch({ runtime, roots: [], author: "Ada", version: "1.0.0" })).toBeNull();
  });
});

describe("configuredAuthor", () => {
  test("trims the setting", () => {
    expect(configuredAuthor("  Ada Lovelace ")).toBe("Ada Lovelace");
  });

  test("treats blank as unset", () => {
    expect(configuredAuthor(" \t")).toBeUndefined();
    expect(configuredAuthor(undefined)).toBeUndefined();
  });
});
